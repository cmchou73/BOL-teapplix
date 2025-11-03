# app.py — Streamlit BOL 產生器
# 變更：
# - Sidebar 新增多筆 PO 搜尋：「搜尋HD PO(一行一個PO):」多行輸入 + 「查詢 PO」按鈕
# - 搜尋包含：已出貨/未出貨、所有 ShipClass（含 UNSP_CG）
# - 搜尋結果與日期抓單結果合併（OID 去重），沿用同一個表格與輸出流程
# - 仍保留：只允許編輯「Select、Warehouse」、批次修改倉庫、訂單日期顯示 mm/dd/yy

import os
import io
import zipfile
from datetime import datetime, timedelta

import requests
import streamlit as st

try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo

from dotenv import load_dotenv
import fitz  # PyMuPDF

APP_TITLE = "Teapplix HD BOL 產生器"
TEMPLATE_PDF = "BOL.pdf"
OUTPUT_DIR = "output_bols"
BASE_URL  = "https://api.teapplix.com/api2/OrderNotification"
STORE_KEY = "HD"
PAGE_SIZE = 500

CHECKBOX_FIELDS   = {"MasterBOL", "Term_Pre", "Term_Collect", "Term_CustChk", "FromFOB", "ToFOB"}
FORCE_TEXT_FIELDS = {"PrePaid", "Collect", "3rdParty"}

BILL_NAME         = "THE HOME DEPOT"
BILL_ADDRESS      = "2455 PACES FERRY RD"
BILL_CITYSTATEZIP = "ATLANTA, GA 30339"

# ---------- secrets / env ----------
load_dotenv(override=False)
def _get_secret(name, default=""):
    return st.secrets.get(name, os.getenv(name, default))

TEAPPLIX_TOKEN = _get_secret("TEAPPLIX_TOKEN", "")

# UI 倉庫代號：「CA 91789」「NJ 08816」
WAREHOUSES = {
    "CA 91789": {
        "name": _get_secret("W1_NAME", "Festival Neo CA"),
        "addr": _get_secret("W1_ADDR", "5500 Mission Blvd"),
        "citystatezip": _get_secret("W1_CITYSTATEZIP", "Montclair, CA 91763"),
        "sid": _get_secret("W1_SID", "CA-001"),
    },
    "NJ 08816": {
        "name": _get_secret("W2_NAME", "Festival Neo NJ"),
        "addr": _get_secret("W2_ADDR", "10 Main St"),
        "citystatezip": _get_secret("W2_CITYSTATEZIP", "East Brunswick, NJ 08816"),
        "sid": _get_secret("W2_SID", "NJ-001"),
    },
}

# ---------- utils ----------
def phoenix_range_days(days=3):
    tz = ZoneInfo("America/Phoenix")
    now = datetime.now(tz)
    start = (now - timedelta(days=days)).replace(hour=0, minute=0, second=0, microsecond=0)
    end   = now.replace(hour=0, minute=0, second=0, microsecond=0)
    fmt = "%Y/%m/%d"
    return start.strftime(fmt), end.strftime(fmt)

def get_headers():
    return {"APIToken": TEAPPLIX_TOKEN, "Content-Type": "application/json;charset=UTF-8", "Accept": "application/json"}

def oz_to_lb(oz):
    try: return round(float(oz)/16.0, 2)
    except Exception: return None

def summarize_packages(order):
    details = order.get("ShippingDetails") or []
    total_pkgs = 0
    total_lb = 0.0
    for sd in details:
        pkg = sd.get("Package") or {}
        count = int(pkg.get("IdenticalPackageCount") or 1)
        wt = pkg.get("Weight") or {}
        lb = oz_to_lb(wt.get("Value")) or 0.0
        total_pkgs += max(1, count)
        total_lb   += lb * max(1, count)
    return total_pkgs, int(round(total_lb))

def override_carrier_name_by_scac(scac: str, current_name: str) -> str:
    s = (scac or "").strip().upper()
    mapping = {
        "EXLA": "Estes Express Lines",
        "AACT": "AAA Cooper Transportation",
        "CTII": "Central Transport Inc.",
        "CETR": "Central Transport Inc.",
        "ABF":  "ABF",
        "PITD": "PITT Ohio",
    }
    return mapping.get(s, current_name)

def group_by_original_txn(orders):
    grouped = {}
    for order in orders:
        oid = (order.get("OriginalTxnId") or "").strip()
        grouped.setdefault(oid, []).append(order)
    return grouped

def _first_item(order):
    items = order.get("OrderItems") or []
    if isinstance(items, list) and items:
        return items[0]
    if isinstance(items, dict):
        return items
    return {}

def _desc_value_from_order(order):
    sku = (_first_item(order).get("ItemSKU") or "")
    return f"{sku}Electric Fireplace".strip()

def _sku8_from_order(order):
    sku = (_first_item(order).get("ItemSKU") or "")
    return sku[:8] if sku else ""

def _qty_from_order(order):
    it = _first_item(order)
    try: return int(it.get("Quantity") or 0)
    except Exception: return 0

def _sum_group_totals(group):
    total_pkgs = 0
    total_lb = 0.0
    for od in group:
        pkgs, lb = summarize_packages(od)
        total_pkgs += int(pkgs or 0)
        total_lb   += float(lb or 0.0)
    return total_pkgs, int(round(total_lb))

def set_widget_value(widget, name, value):
    try:
        is_checkbox_type  = (widget.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX)
        is_checkbox_named = (name in CHECKBOX_FIELDS)
        is_forced_text    = (name in FORCE_TEXT_FIELDS)
        is_checkbox       = (is_checkbox_type or is_checkbox_named) and not is_forced_text
        if is_checkbox:
            v = str(value).strip().lower()
            widget.field_value = "Yes" if v in {"on","yes","1","true","x","✔"} else "Off"
        else:
            widget.field_value = "" if value is None else str(value)
        widget.update()
        return True
    except Exception as e:
        st.warning(f"填欄位 {name} 失敗：{e}")
        return False

# 訂單日期（只顯示日期 mm/dd/yy）
def _parse_order_date_str(first_order):
    tz_phx = ZoneInfo("America/Phoenix")
    od = first_order.get("OrderDetails") or {}
    candidates = [
        od.get("PaymentDate"),
        od.get("OrderDate"),
        first_order.get("PaymentDate"),
        first_order.get("Created"),
        first_order.get("CreateDate"),
    ]
    raw = next((v for v in candidates if v), None)
    if not raw:
        return ""
    val = str(raw).strip()

    dt = None
    try:
        if "T" in val:
            dt = datetime.fromisoformat(val.replace("Z", "+00:00"))
        else:
            for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d", "%Y-%m-%d"):
                try:
                    dt = datetime.strptime(val, fmt)
                    break
                except Exception:
                    continue
    except Exception:
        dt = None

    if dt is None:
        try:
            dt = datetime.fromisoformat(val[:19])
        except Exception:
            return ""

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=tz_phx)
    dt_phx = dt.astimezone(tz_phx)
    return dt_phx.strftime("%m/%d/%y")

# ---------- API：日期範圍抓單（未出貨、排除 UNSP_CG） ----------
def fetch_orders_by_days(days: int):
    ps, pe = phoenix_range_days(days)
    page = 1
    all_orders = []
    while True:
        params = {
            "PaymentDateStart": ps,
            "PaymentDateEnd": pe,
            "Shipped": "0",  # 僅未出貨
            "StoreKey": STORE_KEY,
            "PageSize": str(PAGE_SIZE),
            "PageNumber": str(page),
            "Combine": "combine",
            "DetailLevel": "shipping|inventory|marketplace",
        }
        r = requests.get(BASE_URL, headers=get_headers(), params=params, timeout=45)
        if r.status_code != 200:
            st.error(f"[抓單] API 錯誤: {r.status_code}\n{r.text}")
            break
        try:
            data = r.json()
        except Exception:
            st.error(f"[抓單] JSON 解析錯誤：{r.text[:1000]}")
            break

        orders = data.get("orders") or data.get("Orders") or []
        if not orders: break

        for o in orders:
            od = o.get("OrderDetails", {})
            # 依你既有邏輯：排除 UNSP_CG
            if (od.get("ShipClass") or "").strip().upper() != "UNSP_CG":
                all_orders.append(o)

        if len(orders) < PAGE_SIZE: break
        page += 1
    return all_orders

# ---------- API：以 OID 搜尋（含已出貨/未出貨、含所有 ShipClass） ----------
def fetch_orders_by_oids(oids):
    """對每個 OriginalTxnId 嘗試查詢；如果 API 不支援 OID 直查，就 fallback 近 90 天已/未出貨各抓一輪再 client 過濾。"""
    results = []
    oids = [o.strip() for o in oids if o.strip()]
    if not oids:
        return results

    # 方案 A：假設 API 支援 OriginalTxnId 參數
    for oid in oids:
        ok = False
        for shipped in ("0", "1"):  # 未出貨 / 已出貨 各查一輪
            params = {
                "StoreKey": STORE_KEY,
                "Combine": "combine",
                "DetailLevel": "shipping|inventory|marketplace",
                "PageSize": "500",
                "PageNumber": "1",
                "Shipped": shipped,
                "OriginalTxnId": oid,   # ★ 嘗試直查 OID
            }
            try:
                r = requests.get(BASE_URL, headers=get_headers(), params=params, timeout=45)
                if r.status_code == 200:
                    data = r.json()
                    orders = data.get("orders") or data.get("Orders") or []
                    # 不過濾 ShipClass，UNSP_CG 也保留
                    for o in orders:
                        if (o.get("OriginalTxnId") or "").strip() == oid:
                            results.append(o)
                    ok = True
                else:
                    # 可能不支援 OID 參數，交給 B 方案
                    ok = False
                    break
            except Exception:
                ok = False
                break
        if not ok:
            # 方案 B：90 天窗口 + 已/未出貨；client 端 OID 過濾
            ps, pe = phoenix_range_days(90)
            for shipped in ("0", "1"):
                page = 1
                while True:
                    params = {
                        "PaymentDateStart": ps,
                        "PaymentDateEnd": pe,
                        "Shipped": shipped,
                        "StoreKey": STORE_KEY,
                        "PageSize": str(PAGE_SIZE),
                        "PageNumber": str(page),
                        "Combine": "combine",
                        "DetailLevel": "shipping|inventory|marketplace",
                    }
                    r = requests.get(BASE_URL, headers=get_headers(), params=params, timeout=45)
                    if r.status_code != 200:
                        break
                    try:
                        data = r.json()
                    except Exception:
                        break
                    orders = data.get("orders") or data.get("Orders") or []
                    if not orders:
                        break
                    for o in orders:
                        if (o.get("OriginalTxnId") or "").strip() == oid:
                            results.append(o)
                    if len(orders) < PAGE_SIZE:
                        break
                    page += 1
    return results

# ---------- PDF 欄位建構 ----------
def build_row_from_group(oid, group, wh_key: str):
    first = group[0]
    to = first.get("To") or {}
    od = first.get("OrderDetails") or {}

    ship_details = (first.get("ShippingDetails") or [{}])[0] or {}
    pkg = ship_details.get("Package") or {}
    tracking = pkg.get("TrackingInfo") or {}

    scac_from_shipclass = (od.get("ShipClass") or "").strip()
    carrier_name_raw = (tracking.get("CarrierName") or "").strip()
    carrier_name_final = override_carrier_name_by_scac(scac_from_shipclass, carrier_name_raw)

    street  = (to.get("Street") or "")
    street2 = (to.get("Street2") or "")
    to_address = (street + (" " + street2 if street2 else "")).strip()
    custom_code = (od.get("Custom") or "").strip()

    total_pkgs, total_lb = _sum_group_totals(group)
    bol_num = (od.get("Invoice") or "").strip() or (oid or "").strip()

    WH = WAREHOUSES.get(wh_key, list(WAREHOUSES.values())[0])

    row = {
        "BillName": BILL_NAME,
        "BillAddress": BILL_ADDRESS,
        "BillCityStateZip": BILL_CITYSTATEZIP,
        "ToName": to.get("Name", ""),
        "ToAddress": to_address,
        "ToCityStateZip": f"{to.get('City','')}, {to.get('State','')} {to.get('ZipCode','')}".strip().strip(", "),
        "ToCID": to.get("PhoneNumber", ""),
        "FromName": WH["name"],
        "FromAddr": WH["addr"],
        "FromCityStateZip": WH["citystatezip"],
        "FromSIDNum": WH["sid"],
        "3rdParty": "X", "PrePaid": "", "Collect": "",
        "BOLnum": bol_num,
        "CarrierName": carrier_name_final,
        "SCAC": scac_from_shipclass,
        "PRO": tracking.get("TrackingNumber", ""),
        "CustomerOrderNumber": custom_code,
        "BillInstructions": f"PO#{oid or bol_num}",
        "OrderNum1": custom_code,
        "SpecialInstructions": "",
        "TotalPkgs": str(total_pkgs) if total_pkgs else "",
        "Total_Weight": str(total_lb) if total_lb else "",
        "Date": datetime.now().strftime("%Y/%m/%d"),
        "Page_ttl": "1",
        "NMFC1": "69420",
        "Class1": "125",
    }

    total_qty_sum = 0
    for idx, od_item in enumerate(group, start=1):
        desc_val = _desc_value_from_order(od_item)
        qty = _qty_from_order(od_item)
        if desc_val:
            row[f"Desc_{idx}"] = desc_val
            row[f"HU_Type_{idx}"]  = "piece"
            row[f"Pkg_Type_{idx}"] = "piece"
            row[f"HU_QTY_{idx}"]   = str(qty) if qty else ""
            row[f"Pkg_QTY_{idx}"]  = str(qty) if qty else ""
            total_qty_sum += qty

    row["NumPkgs1"] = str(total_qty_sum)
    row["Weight1"] = "130 lbs" if total_qty_sum <= 1 else f"{130 + (total_qty_sum - 1) * 30} lbs"
    return row, WH

def fill_pdf(row: dict, out_path: str):
    if not os.path.exists(TEMPLATE_PDF):
        raise FileNotFoundError(f"找不到 BOL 模板：{TEMPLATE_PDF}")
    doc = fitz.open(TEMPLATE_PDF)
    for page in doc:
        for w in (page.widgets() or []):
            name = w.field_name
            if name and name in row:
                set_widget_value(w, name, row[name])
    try: doc.need_appearances = True
    except Exception: pass
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path, deflate=True, incremental=False, encryption=fitz.PDF_ENCRYPT_KEEP)
    doc.close()

# ---------- Streamlit UI ----------
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

# 說明（沿用你上版）
st.markdown("""
**說明：**
1. 可能會錯, 請仔細核對
2. ABCD
""")

if not TEAPPLIX_TOKEN:
    st.error("找不到 TEAPPLIX_TOKEN，請在 .env 或 Streamlit Secrets 設定。")
    st.stop()

# Sidebar：天數 & PO 搜尋
days = st.sidebar.selectbox("抓取天數", options=[1,2,3,4,5,6,7], index=2, help="預設 3 天")
st.sidebar.markdown("**搜尋HD PO(一行一個PO):**")
po_text = st.sidebar.text_area("在此貼上 OriginalTxnId（每行一個）", height=150, label_visibility="collapsed")
if st.sidebar.button("查詢 PO"):
    oids = list({line.strip() for line in po_text.splitlines() if line.strip()})
    st.session_state["search_oids"] = oids
    if oids:
        st.session_state["orders_search_raw"] = fetch_orders_by_oids(oids)
    else:
        st.session_state["orders_search_raw"] = []

# 主操作：抓單
if st.button("抓取訂單", use_container_width=True):
    st.session_state["orders_days_raw"] = fetch_orders_by_days(days)
    st.session_state.pop("table_rows_override", None)

orders_days_raw   = st.session_state.get("orders_days_raw", None)
orders_search_raw = st.session_state.get("orders_search_raw", None)

# 合併來源：日期抓單 + PO 搜尋（OID 去重）
all_orders = []
if orders_days_raw:
    all_orders.extend(orders_days_raw)
if orders_search_raw:
    all_orders.extend(orders_search_raw)

if all_orders:
    grouped_all = group_by_original_txn(all_orders)

    # 準備表格資料
    if "table_rows_override" in st.session_state:
        table_rows = st.session_state["table_rows_override"]
    else:
        seen = set()
        table_rows = []
        for oid, group in grouped_all.items():
            if not oid or oid in seen:
                continue
            seen.add(oid)
            first = group[0]
            od = first.get("OrderDetails", {}) or {}
            scac = (od.get("ShipClass") or "").strip()
            sku8 = _sku8_from_order(first)
            order_date_str = _parse_order_date_str(first)  # mm/dd/yy
            table_rows.append({
                "Select": True,
                "Warehouse": "CA 91789",  # 預設
                "OriginalTxnId": oid,
                "SKU8": sku8,
                "SCAC": scac,
                "ToState": (first.get("To") or {}).get("State",""),
                "OrderDate": order_date_str,
            })

    st.caption(f"共 {len(table_rows)} 筆（依 OriginalTxnId 合併；含日期抓單 + PO 搜尋）")

    # 批次修改倉庫
    bulk_col1, bulk_col2, bulk_col3 = st.columns([1,1,6])
    with bulk_col1:
        bulk_wh = st.selectbox("批次指定倉庫", options=list(WAREHOUSES.keys()), index=0)
    with bulk_col2:
        apply_to = st.selectbox("套用對象", options=["勾選列", "全部"], index=0)
    with bulk_col3:
        if st.button("套用批次倉庫", use_container_width=True):
            new_rows = []
            if apply_to == "全部":
                for r in table_rows:
                    r2 = dict(r); r2["Warehouse"] = bulk_wh; new_rows.append(r2)
            else:
                for r in table_rows:
                    r2 = dict(r)
                    if r2.get("Select"):
                        r2["Warehouse"] = bulk_wh
                    new_rows.append(r2)
            st.session_state["table_rows_override"] = new_rows
            table_rows = new_rows
            st.success("已套用批次倉庫變更。")

    # 表格（僅允許 Select、Warehouse 可編輯）
    edited = st.data_editor(
        table_rows,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Select": st.column_config.CheckboxColumn("選取", default=True),
            "Warehouse": st.column_config.SelectboxColumn("倉庫", options=list(WAREHOUSES.keys())),
            "OriginalTxnId": st.column_config.TextColumn("PO", disabled=True),
            "SKU8": st.column_config.TextColumn("SKU", disabled=True),
            "SCAC": st.column_config.TextColumn("SCAC", disabled=True),
            "ToState": st.column_config.TextColumn("州", disabled=True),
            "OrderDate": st.column_config.TextColumn("訂單日期 (mm/dd/yy)", disabled=True),
        },
        key="orders_table",
    )

    # 產出 BOL
    if st.button("產生 BOL（勾選列）", type="primary", use_container_width=True):
        selected = [r for r in edited if r.get("Select")]
        if not selected:
            st.warning("尚未選取任何訂單。")
        else:
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            made_files = []
            for row_preview in selected:
                oid = row_preview["OriginalTxnId"]
                wh_key = row_preview["Warehouse"]
                group = grouped_all.get(oid, [])
                if not group:
                    continue

                row_dict, WH = build_row_from_group(oid, group, wh_key)

                sku8 = row_preview["SKU8"] or (_sku8_from_order(group[0]) or "NOSKU")[:8]
                wh2 = (WH["name"][:2].upper() if WH["name"] else "WH")
                scac = (row_preview["SCAC"] or "").upper() or "NOSCAC"
                filename = f"BOL_{oid}_{sku8}_{wh2}_{scac}.pdf".replace(" ", "")
                out_path = os.path.join(OUTPUT_DIR, filename)

                fill_pdf(row_dict, out_path)
                made_files.append(out_path)

            if made_files:
                st.success(f"已產生 {len(made_files)} 份 BOL。")
                mem_zip = io.BytesIO()
                with zipfile.ZipFile(mem_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                    for p in made_files:
                        zf.write(p, arcname=os.path.basename(p))
                mem_zip.seek(0)
                st.download_button(
                    "下載全部 BOL (ZIP)",
                    data=mem_zip,
                    file_name=f"BOL_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
            else:
                st.warning("沒有產生任何檔案。")
else:
    st.info("請先用左側輸入 PO（選擇性）或選天數後按『抓取訂單』。")
