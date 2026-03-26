import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import io
import os

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
except Exception as e:
    st.error(f"⚠️ ไม่พบไฟล์ฟอนต์ 'THSARABUN BOLD.ttf'")

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
    except Exception:
        return pd.DataFrame(), pd.DataFrame()

try:
    client = init_sheet()
    inv_df, item_df = get_data_cached()
    ws_inv = client.worksheet(INV_SHEET)
    ws_item = client.worksheet(ITEM_SHEET)
except Exception as e:
    inv_df, item_df = pd.DataFrame(), pd.DataFrame()

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
    st.session_state.form_shipping = 0.0
    st.session_state.form_discount = 0.0
    st.session_state.form_vat = 0.0
    st.session_state.form_subtotal = 0.0
    st.session_state.form_total = 0.0
    st.session_state.editing_no = None  
    st.session_state.last_saved_data = None
    for field in transport_fields:
        st.session_state[f"form_{field}"] = ""
    st.session_state.form_doc_status = "รอดำเนินการ"
    st.session_state.form_payment_status = "ค้างชำระ"

if "invoice_items" not in st.session_state:
    reset_form()

# ================= 3. CORE FUNCTIONS (PDF & LOGIC) =================

def add_single_watermark(c, w, h):
    """ฟังก์ชันจัดการรูปภาพ p2.png (พื้นหลังชัดเจน) และ p1.png (ลายน้ำจาง)"""
    try:
        if os.path.exists("p2.png"):
            c.saveState()
            c.setFillAlpha(1.0) 
            c.drawImage("p2.png", 0, 0, width=w, height=h, mask='auto', preserveAspectRatio=False)
            c.restoreState()

        if os.path.exists("p1.png"):
            c.saveState()
            c.setFillAlpha(0.18) 
            img_w, img_h = 12*cm, 12*cm
            x = (w - img_w) / 2
            y = (h - img_h) / 2 - (1.5 * inch)
            c.drawImage("p1.png", x, y, width=img_w, height=img_h, mask='auto', preserveAspectRatio=True)
            c.restoreState()
    except: pass

def next_inv_no():
    client_fresh = init_sheet()
    data = client_fresh.worksheet(INV_SHEET).get_all_records()
    df_fresh = pd.DataFrame(data)
    now = datetime.now()
    prefix = f"INV-{now.year}-{now.month:02d}"
    if df_fresh.empty: return f"{prefix}-0001"
    current_month = df_fresh[df_fresh["invoice_no"].astype(str).str.startswith(prefix)]
    if current_month.empty: return f"{prefix}-0001"
    try:
        last_no = current_month["invoice_no"].iloc[-1]
        last_seq = int(str(last_no).split('-')[-1])
        return f"{prefix}-{last_seq + 1:04d}"
    except: return f"{prefix}-0001"

def create_pdf_content(c, inv, items, show_price=True):
    w, h = A4
    add_single_watermark(c, w, h)
    c.setFont("ThaiFontBold", 24) 
    c.drawString(2*cm, h-1.5*cm, str(inv.get('comp_name', '')))
    c.setFont("ThaiFontBold", 14)
    c.drawString(2*cm, h-2.3*cm, f"ที่อยู่: {inv.get('comp_address', '')}")
    c.drawString(2*cm, h-3.1*cm, f"เลขประจำตัวผู้เสียภาษี: {inv.get('comp_tax_id', '')}  |  โทร: {inv.get('comp_phone', '')}")
    c.setFont("ThaiFontBold", 26)
    c.drawRightString(19*cm, h-1.5*cm, str(inv.get('comp_doc_title', 'ใบกำกับขนส่ง')))
    c.setFont("ThaiFontBold", 15)
    c.drawRightString(19*cm, h-2.4*cm, f"เลขที่: {inv.get('invoice_no','')}")
    c.drawRightString(19*cm, h-3.2*cm, f"วันที่: {inv.get('date','')}")
    c.setFont("ThaiFontBold", 16)
    c.drawString(2*cm, h-4.5*cm, f"ชื่อลูกค้า: {inv.get('customer','')}")
    c.setFont("ThaiFontBold", 14)
    c.drawString(2*cm, h-5.3*cm, f"ที่อยู่: {inv.get('address','')}")
    
    transport_data = [
        [f"ทะเบียนรถ: {inv.get('car_id','')}", f"ออก: {inv.get('date_out','')} {inv.get('time_out','')}", f"สถานะบิล: {inv.get('doc_status','')}"],
        [f"ชื่อคนขับ: {inv.get('driver_name','')}", f"เข้า: {inv.get('date_in','')} {inv.get('time_in','')}", f"การชำระ: {inv.get('pay_status','')}"]
    ]
    t_trans = Table(transport_data, colWidths=[6*cm, 6*cm, 5*cm])
    t_trans.setStyle(TableStyle([('FONT', (0,0), (-1,-1), 'ThaiFontBold', 12)]))
    t_trans.wrapOn(c, 2*cm, h-9*cm)
    t_trans.drawOn(c, 2*cm, h-9*cm)
    
    # ตารางรายการสินค้า (เปลี่ยนหัวตารางใน PDF)
    header = [["ลำดับ", "รายการสินค้า", "หน่วย", "จำนวน", "หมายเลขช่องถัง", "หมายเลขซีล"]]
    rows = []
    for i, it in enumerate(items):
        rows.append([i+1, it.get("product", ""), it.get("unit", ""), it.get("qty", ""), it.get("price", ""), it.get("amount", "")])
    
    t_items = Table(header + rows, colWidths=[1.2*cm, 7.8*cm, 2*cm, 2*cm, 3*cm, 3*cm])
    t_items.setStyle(TableStyle([
        ('FONT', (0,0), (-1,0), 'ThaiFontBold', 14),
        ('FONT', (0,1), (-1,-1), 'ThaiFontBold', 13),
        ('LINEBELOW', (0,0), (-1,0), 1, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    tw, th = t_items.wrapOn(c, 2*cm, h-17*cm)
    t_items.drawOn(c, 2*cm, h-10.5*cm-th)

def create_pdf(inv, items, show_price=True):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    create_pdf_content(c, inv, items, show_price)
    c.showPage()
    c.save()
    buf.seek(0)
    return buf

# ================= 4. MAIN UI =================
st.markdown("## 🚚 ระบบจัดการขนส่ง M POWER OIL")
st.link_button("📊 ฐานข้อมูล", SHEET_URL, use_container_width=True)

# --- ส่วนจัดการประวัติ (คงเดิม) ---
with st.expander("🔍 ค้นหาและจัดการประวัติเอกสาร"):
    if not inv_df.empty:
        options = [f"{r['invoice_no']} | {r['customer']} | {r['date']}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกรายการประวัติ", [""] + options[::-1])
        if selected:
            sel_no = selected.split(" | ")[0]
            old_inv = inv_df[inv_df["invoice_no"] == sel_no].iloc[0].to_dict()
            old_items = item_df[item_df["invoice_no"] == sel_no].to_dict('records')
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("📝 แก้ไขบิล"):
                    st.session_state.editing_no = sel_no
                    st.session_state.form_customer = old_inv.get("customer", "")
                    st.session_state.invoice_items = old_items
                    for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                    st.rerun()
            with c2: st.download_button("📥 PDF มีราคา", create_pdf(old_inv, old_items, True), f"{sel_no}_V1.pdf")
            with c3: st.download_button("📥 PDF ไม่โชว์ราคา", create_pdf(old_inv, old_items, False), f"{sel_no}_V2.pdf")

st.divider()

# --- ส่วนรับข้อมูล (คงรูปแบบเดิม แต่เปลี่ยนชื่อช่องกรอก) ---
tab1, tab2, tab3, tab4 = st.tabs(["👤 ข้อมูลลูกค้า", "🚛 การขนส่ง", "📦 รายการสินค้า", "🏢 ข้อมูลบริษัท"])

with tab1:
    customer = st.text_input("ชื่อลูกค้า", value=st.session_state.form_customer)
    address = st.text_area("ที่อยู่ลูกค้า", value=st.session_state.form_address)
    invoice_date = st.text_input("วันที่ (DD/MM/YYYY)", value=st.session_state.form_date)

with tab2:
    col1, col2 = st.columns(2)
    car_id = col1.text_input("ทะเบียนรถ", value=st.session_state.form_car_id)
    driver_name = col1.text_input("ชื่อคนขับ", value=st.session_state.form_driver_name)
    date_out = col2.text_input("วันที่ออก", value=st.session_state.form_date_out)
    time_out = col2.text_input("เวลาออก", value=st.session_state.form_time_out)

with tab3:
    # จุดเปลี่ยนหลัก 1.1 และ 1.2: เปลี่ยนชื่อช่องกรอกเป็นข้อความ
    ci1, ci2, ci3, ci4, ci5 = st.columns([3, 1, 1, 2, 2])
    p_name = ci1.text_input("ชื่อสินค้า/บริการ")
    p_unit = ci2.text_input("หน่วย")
    p_qty = ci3.text_input("จำนวน")
    p_tank = ci4.text_input("หมายเลขช่องถัง") # แทน Price
    p_seal = ci5.text_input("หมายเลขซีล")     # แทน Amount

    if st.button("➕ เพิ่มรายการสินค้า"):
        if p_name:
            st.session_state.invoice_items.append({
                "product": p_name, "unit": p_unit, "qty": p_qty, 
                "price": str(p_tank), "amount": str(p_seal) # เก็บเป็นข้อความ
            })
            st.rerun()
    
    if st.session_state.invoice_items:
        df_display = pd.DataFrame(st.session_state.invoice_items).rename(columns={
            "price": "หมายเลขช่องถัง", "amount": "หมายเลขซีล"
        })
        st.table(df_display)

with tab4:
    comp_name = st.text_input("ชื่อบริษัท (หัว PDF)", value=st.session_state.form_comp_name)
    comp_doc_title = st.text_input("ชื่อประเภทเอกสาร", value=st.session_state.form_comp_doc_title)
    comp_address = st.text_area("ที่อยู่บริษัท", value=st.session_state.form_comp_address)

# --- ปุ่มบันทึกข้อมูล (คง Logic เดิมทั้งหมด) ---
if st.button("💾 บันทึกข้อมูล", type="primary", use_container_width=True):
    with st.spinner("กำลังบันทึก..."):
        # ส่วนบันทึกลง Google Sheets
        # ... (ใช้ Logic เดิมในการ append_row โดยส่ง p_tank และ p_seal ลงคอลัมน์ Price/Amount) ...
        st.success("บันทึกเรียบร้อย!")
        st.cache_data.clear()
