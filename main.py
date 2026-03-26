import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import io
import os
import time

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors

# ================= 1. CONFIG & INITIALIZATION =================
st.set_page_config(page_title="Logistics System Pro", layout="wide")

try:
    pdfmetrics.registerFont(TTFont('ThaiFontBold', 'THSARABUN BOLD.ttf'))
except Exception:
    st.error("⚠️ ไม่พบไฟล์ฟอนต์ 'THSARABUN BOLD.ttf'")

SHEET_ID = "1fl86CxqgxlXAYU63GQOdCrL2jbPvSUdoXd1ndQvjnBM"
INV_SHEET = "Invoices"
ITEM_SHEET = "InvoiceItems"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"

@st.cache_resource
def init_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    return gspread.authorize(creds).open_by_key(SHEET_ID)

@st.cache_data(ttl=60)
def get_data_cached():
    client = init_sheet()
    try:
        inv = client.worksheet(INV_SHEET).get_all_records()
        items = client.worksheet(ITEM_SHEET).get_all_records()
        return pd.DataFrame(inv), pd.DataFrame(items)
    except:
        return pd.DataFrame(), pd.DataFrame()

# โหลดข้อมูล
inv_df, item_df = get_data_cached()
client = init_sheet()
ws_inv = client.worksheet(INV_SHEET)
ws_item = client.worksheet(ITEM_SHEET)

# ================= 2. SESSION STATE & FORM RESET =================
transport_fields = [
    "doc_status", "car_id", "driver_name", "payment_status", "date_out", "time_out",
    "date_in", "time_in", "ref_tax_id", "ref_receipt_id", "seal_no",
    "pay_term", "ship_method", "driver_license", "receiver_name",
    "issuer_name", "sender_name", "checker_name", "remark",
    "comp_name", "comp_address", "comp_tax_id", "comp_phone", "comp_doc_title"
]

def reset_form():
    st.session_state.invoice_items = []
    st.session_state.form_customer = ""
    st.session_state.form_address = ""
    st.session_state.form_date = datetime.now().strftime("%d/%m/%Y") 
    st.session_state.editing_no = None  
    for field in transport_fields:
        st.session_state[f"form_{field}"] = ""
    st.session_state.form_doc_status = "รอดำเนินการ"
    st.session_state.form_payment_status = "ค้างชำระ"

if "invoice_items" not in st.session_state:
    reset_form()

# ================= 3. CORE FUNCTIONS (PDF & LOGIC) =================

def add_single_watermark(c, w, h):
    """p2.png พื้นหลังชัด 100% และ p1.png ลายน้ำจาง"""
    try:
        if os.path.exists("p2.png"):
            c.saveState()
            c.setFillAlpha(1.0) 
            c.drawImage("p2.png", 0, 0, width=w, height=h, mask='auto', preserveAspectRatio=False)
            c.restoreState()
        if os.path.exists("p1.png"):
            c.saveState()
            c.setFillAlpha(0.18) 
            c.drawImage("p1.png", (w-12*cm)/2, (h-12*cm)/2 - (1.5*inch), width=12*cm, height=12*cm, mask='auto', preserveAspectRatio=True)
            c.restoreState()
    except: pass

def next_inv_no():
    client_fresh = init_sheet()
    data = client_fresh.worksheet(INV_SHEET).get_all_records()
    df_fresh = pd.DataFrame(data)
    now = datetime.now()
    prefix = f"INV-{now.year}-{now.month:02d}"
    if df_fresh.empty: return f"{prefix}-0001"
    curr = df_fresh[df_fresh["invoice_no"].astype(str).str.startswith(prefix)]
    if curr.empty: return f"{prefix}-0001"
    last_seq = int(str(curr["invoice_no"].iloc[-1]).split('-')[-1])
    return f"{prefix}-{last_seq + 1:04d}"

def create_pdf_content(c, inv, items):
    w, h = A4
    add_single_watermark(c, w, h)
    c.setFont("ThaiFontBold", 24) 
    c.drawString(2*cm, h-1.5*cm, str(inv.get('comp_name', '')))
    c.setFont("ThaiFontBold", 26)
    c.drawRightString(19*cm, h-1.5*cm, str(inv.get('comp_doc_title', 'ใบกำกับขนส่ง')))
    c.setFont("ThaiFontBold", 15)
    c.drawRightString(19*cm, h-2.4*cm, f"เลขที่: {inv.get('invoice_no','')}")
    c.drawRightString(19*cm, h-3.2*cm, f"วันที่: {inv.get('date','')}")
    
    # ตารางรายการสินค้า
    header = [["ลำดับ", "รายการสินค้า", "หน่วย", "จำนวน", "หมายเลขช่องถัง", "หมายเลขซีล"]]
    rows = [[i+1, it['product'], it['unit'], it['qty'], it['price'], it['amount']] for i, it in enumerate(items)]
    t_items = Table(header + rows, colWidths=[1.2*cm, 7.8*cm, 1.5*cm, 1.5*cm, 3*cm, 3*cm])
    t_items.setStyle(TableStyle([('FONT', (0,0), (-1,0), 'ThaiFontBold', 14), ('FONT', (0,1), (-1,-1), 'ThaiFontBold', 13), ('LINEBELOW', (0,0), (-1,0), 1, colors.black)]))
    tw, th = t_items.wrapOn(c, 2*cm, h-17*cm)
    t_items.drawOn(c, 2*cm, h-10.5*cm-th)

def create_pdf(inv, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    create_pdf_content(c, inv, items)
    c.showPage(); c.save(); buf.seek(0)
    return buf

# ================= 4. MAIN UI =================
st.markdown("## 🚚 ใบกำกับขนส่ง M POWER OIL")
st.link_button("📊 ฐานข้อมูล", SHEET_URL, use_container_width=True)

# --- ส่วนจัดการประวัติ (แก้ไข/ทำซ้ำ) ---
with st.expander("🔍 ค้นหาและจัดการประวัติเอกสาร"):
    if not inv_df.empty:
        options = [f"{r['invoice_no']} | {r['customer']} | วันที่: {r['date']}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกรายการประวัติ", [""] + options[::-1])
        if selected:
            sel_no = selected.split(" | ")[0]
            old_inv = inv_df[inv_df["invoice_no"] == sel_no].iloc[0].to_dict()
            raw_items = item_df[item_df["invoice_no"] == sel_no].to_dict('records')
            # Map ข้อมูลจาก Sheet กลับเข้า Form
            formatted_items = [{"product": it['product'], "unit": it['unit'], "qty": it['qty'], "price": str(it['price']), "amount": str(it['amount'])} for it in raw_items]
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                if st.button("🔄 สร้างรายการซ้ำ"):
                    reset_form()
                    st.session_state.form_customer, st.session_state.form_address = old_inv.get("customer",""), old_inv.get("address","")
                    st.session_state.invoice_items = formatted_items
                    for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                    st.rerun()
            with c2:
                if st.button("📝 แก้ไขบิล"):
                    st.session_state.editing_no = sel_no
                    st.session_state.form_customer, st.session_state.form_address = old_inv.get("customer",""), old_inv.get("address","")
                    st.session_state.invoice_items = formatted_items
                    for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                    st.rerun()
            with c3: st.download_button("📥 PDF มีราคา", create_pdf(old_inv, formatted_items), f"{sel_no}_V1.pdf")
            with c4: st.download_button("📥 PDF ไม่โชว์ราคา", create_pdf(old_inv, formatted_items), f"{sel_no}_V2.pdf")

st.divider()
if st.session_state.editing_no:
    st.warning(f"🚨 กำลังแก้ไขบิล: {st.session_state.editing_no}")
    if st.button("❌ ยกเลิกแก้ไข"): reset_form(); st.rerun()

# --- ส่วนรับข้อมูล Tab 1-4 ---
tab1, tab2, tab3, tab4 = st.tabs(["👤 ลูกค้า", "🚛 ขนส่ง", "📦 สินค้า", "🏢 บริษัท"])
with tab1:
    customer = st.text_input("ชื่อลูกค้า", value=st.session_state.form_customer)
    address = st.text_area("ที่อยู่ลูกค้า", value=st.session_state.form_address)
    invoice_date = st.text_input("วันที่", value=st.session_state.form_date)
with tab2:
    car_id = st.text_input("ทะเบียนรถ", value=st.session_state.form_car_id)
    driver_name = st.text_input("ชื่อคนขับ", value=st.session_state.form_driver_name)
with tab3:
    ci1, ci2, ci3, ci4, ci5 = st.columns([3, 1, 1, 2, 2])
    p_name = ci1.text_input("รายการ")
    p_unit = ci2.text_input("หน่วย", value="ลิตร")
    p_qty = ci3.text_input("จำนวน")
    p_tank = ci4.text_input("หมายเลขช่องถัง") # บันทึกลง Price
    p_seal = ci5.text_input("หมายเลขซีล")     # บันทึกลง Amount
    if st.button("➕ เพิ่มสินค้า"):
        st.session_state.invoice_items.append({"product": p_name, "unit": p_unit, "qty": p_qty, "price": str(p_tank), "amount": str(p_seal)})
        st.rerun()
    if st.session_state.invoice_items:
        st.table(pd.DataFrame(st.session_state.invoice_items).rename(columns={"price":"หมายเลขช่องถัง", "amount":"หมายเลขซีล"}))
with tab4:
    comp_name = st.text_input("ชื่อบริษัทผู้ขาย", value=st.session_state.form_comp_name)
    comp_doc_title = st.text_input("หัวข้อเอกสาร", value=st.session_state.form_comp_doc_title)

# --- ส่วนบันทึกและสร้าง PDF ทันที ---
if st.button("💾 บันทึกข้อมูล", type="primary", use_container_width=True):
    with st.spinner("กำลังบันทึกและสร้าง PDF..."):
        new_no = st.session_state.editing_no if st.session_state.editing_no else next_inv_no()
        if st.session_state.editing_no:
            # ลบแถวเดิมก่อนเขียนใหม่ (Logic เดิม)
            try:
                for ws in [ws_item, ws_inv]:
                    cell = ws.find(new_no)
                    while cell: ws.delete_rows(cell.row); cell = ws.find(new_no)
            except: pass
        
        # บันทึกข้อมูล
        ws_inv.append_row([new_no, invoice_date, customer, address, 0, 0, 0, 0, 0, st.session_state.form_doc_status, car_id, driver_name, st.session_state.form_payment_status])
        for it in st.session_state.invoice_items:
            ws_item.append_row([new_no, it['product'], it['unit'], it['qty'], str(it['price']), str(it['amount'])])
        
        # แสดงปุ่ม PDF เหมือนเดิม
        st.success(f"✅ บันทึกเลขที่ {new_no} สำเร็จ!")
        inv_pdf = {"invoice_no": new_no, "date": invoice_date, "customer": customer, "comp_name": comp_name, "comp_doc_title": comp_doc_title}
        st.download_button("📥 ดาวน์โหลด PDF ทันที", create_pdf(inv_pdf, st.session_state.invoice_items), f"{new_no}.pdf", use_container_width=True)
        st.cache_data.clear()
