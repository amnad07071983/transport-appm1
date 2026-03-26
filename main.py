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
    st.error("⚠️ ไม่พบไฟล์ฟอนต์ 'THSARABUN BOLD.ttf' กรุณาตรวจสอบไฟล์ในโฟลเดอร์")

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
        # ลำดับที่ 1: วาด p2.png เป็นพื้นหลังเต็มหน้า (ความเข้ม 100%) เพื่อความชัดเจนในการพิมพ์
        if os.path.exists("p2.png"):
            c.saveState()
            c.setFillAlpha(1.0) 
            c.drawImage("p2.png", 0, 0, width=w, height=h, mask='auto', preserveAspectRatio=False)
            c.restoreState()

        # ลำดับที่ 2: วาด p1.png เป็นลายน้ำจางๆ ตรงกลาง
        if os.path.exists("p1.png"):
            c.saveState()
            c.setFillAlpha(0.18) 
            img_w, img_h = 12*cm, 12*cm
            x = (w - img_w) / 2
            y = (h - img_h) / 2 - (1.5 * inch)
            c.drawImage("p1.png", x, y, width=img_w, height=img_h, mask='auto', preserveAspectRatio=True)
            c.restoreState()
    except Exception as e:
        print(f"Error rendering images: {e}")

def next_inv_no():
    client_fresh = init_sheet()
    data = client_fresh.worksheet(INV_SHEET).get_all_records()
    df_fresh = pd.DataFrame(data)
    now = datetime.now()
    prefix = f"INV-{now.year}-{now.month:02d}"
    if df_fresh.empty or "invoice_no" not in df_fresh.columns:
        return f"{prefix}-0001"
    current_month_docs = df_fresh[df_fresh["invoice_no"].astype(str).str.startswith(prefix)]
    if current_month_docs.empty:
        return f"{prefix}-0001"
    try:
        last_no = current_month_docs["invoice_no"].iloc[-1]
        last_seq = int(str(last_no).split('-')[-1])
        return f"{prefix}-{last_seq + 1:04d}"
    except:
        return f"{prefix}-0001"

def create_pdf_content(c, inv, items, show_price=True):
    w, h = A4
    add_single_watermark(c, w, h)
    
    # ส่วนหัวบริษัท
    c.setFont("ThaiFontBold", 24) 
    c.drawString(2*cm, h-1.5*cm, str(inv.get('comp_name', '')))
    c.setFont("ThaiFontBold", 14)
    c.drawString(2*cm, h-2.3*cm, f"ที่อยู่: {inv.get('comp_address', '')}")
    c.drawString(2*cm, h-3.1*cm, f"เลขประจำตัวผู้เสียภาษี: {inv.get('comp_tax_id', '')}  |  โทร: {inv.get('comp_phone', '')}")
    
    # ส่วนหัวเอกสาร
    c.setFont("ThaiFontBold", 26)
    c.drawRightString(19*cm, h-1.5*cm, str(inv.get('comp_doc_title', 'ใบกำกับขนส่ง')))
    c.setFont("ThaiFontBold", 15)
    c.drawRightString(19*cm, h-2.4*cm, f"เลขที่: {inv.get('invoice_no','')}")
    c.drawRightString(19*cm, h-3.2*cm, f"วันที่: {inv.get('date','')}")
    
    # ข้อมูลลูกค้า
    c.setFont("ThaiFontBold", 16)
    c.drawString(2*cm, h-4.5*cm, f"ชื่อลูกค้า: {inv.get('customer','')}")
    c.setFont("ThaiFontBold", 14)
    c.drawString(2*cm, h-5.3*cm, f"ที่อยู่: {inv.get('address','')}")
    
    # ตารางรายการสินค้า (ปรับหัวข้อเป็น หมายเลขช่องถัง และ หมายเลขซีล)
    header = [["ลำดับ", "รายการสินค้า/บริการ", "หน่วย", "จำนวน", "หมายเลขช่องถัง", "หมายเลขซีล"]]
    col_w = [1.2*cm, 7.8*cm, 1.5*cm, 1.5*cm, 2.5*cm, 2.5*cm]

    rows = []
    for i, it in enumerate(items):
        rows.append([i+1, it.get("product", ""), it.get("unit", ""), it.get("qty", ""), it.get("price", ""), it.get("amount", "")])
    
    t_items = Table(header + rows, colWidths=col_w)
    t_items.setStyle(TableStyle([
        ('FONT', (0,0), (-1,0), 'ThaiFontBold', 14),
        ('FONT', (0,1), (-1,-1), 'ThaiFontBold', 13),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LINEBELOW', (0,0), (-1,0), 1, colors.black),
    ]))
    tw, th = t_items.wrapOn(c, 2*cm, h-17*cm)
    t_items.drawOn(c, 2*cm, h-10.5*cm-th)

def create_pdf(inv, items, show_price=True):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    create_pdf_content(c, inv, items, show_price=show_price)
    c.showPage()
    c.save()
    buf.seek(0)
    return buf

# ================= 4. MAIN UI =================
st.markdown("## 🚚 ใบกำกับขนส่ง M POWER OIL")

with st.expander("🔍 ค้นหาและจัดการประวัติเอกสาร"):
    if not inv_df.empty:
        options = [f"{r['invoice_no']} | {r['customer']} | วันที่: {r['date']}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกรายการประวัติ", [""] + options[::-1])
        if selected:
            sel_no = selected.split(" | ")[0]
            old_inv = inv_df[inv_df["invoice_no"] == sel_no].iloc[0].to_dict()
            old_items = item_df[item_df["invoice_no"] == sel_no].to_dict('records')
            c1, c2 = st.columns(2)
            with c1: st.download_button(f"📥 PDF มีราคา", create_pdf(old_inv, old_items, True), f"{sel_no}_Price.pdf")
            with c2: st.download_button(f"📥 PDF ไม่โชว์ราคา", create_pdf(old_inv, old_items, False), f"{sel_no}_Qty.pdf")

st.divider()

# --- ส่วนรับข้อมูล Input ---
tab1, tab2, tab3, tab4 = st.tabs(["👤 ลูกค้า", "🚛 ขนส่ง", "📦 สินค้า", "🏢 บริษัท"])
with tab1:
    customer = st.text_input("ชื่อลูกค้า", value=st.session_state.form_customer)
    address = st.text_area("ที่อยู่", value=st.session_state.form_address)
    invoice_date = st.text_input("วันที่ (DD/MM/YYYY)", value=st.session_state.form_date)

with tab2:
    c_id = st.text_input("ทะเบียนรถ", value=st.session_state.form_car_id)
    d_name = st.text_input("ชื่อคนขับ", value=st.session_state.form_driver_name)

with tab3:
    st.info("ระบุข้อมูลสินค้าและหมายเลขถัง/ซีล")
    ci1, ci2, ci3, ci4, ci5 = st.columns([3, 1, 1, 2, 2])
    p_name = ci1.text_input("รายการ")
    p_unit = ci2.text_input("หน่วย", value="ลิตร")
    p_qty = ci3.text_input("จำนวน")
    # ปรับปรุง 1.1: เปลี่ยนชื่อเป็น หมายเลขช่องถัง และรับเป็น text
    p_tank = ci4.text_input("หมายเลขช่องถัง") 
    # ปรับปรุง 1.2: เปลี่ยนชื่อเป็น หมายเลขซีล และรับเป็น text
    p_seal = ci5.text_input("หมายเลขซีล")

    if st.button("➕ เพิ่มสินค้า"):
        st.session_state.invoice_items.append({
            "product": p_name, 
            "unit": p_unit, 
            "qty": p_qty, 
            "price": str(p_tank), # บันทึกลงคอลัมน์ price แบบข้อความ
            "amount": str(p_seal) # บันทึกลงคอลัมน์ amount แบบข้อความ
        })
        st.rerun()

    if st.session_state.invoice_items:
        df_items = pd.DataFrame(st.session_state.invoice_items).rename(columns={"price": "หมายเลขช่องถัง", "amount": "หมายเลขซีล"})
        st.table(df_items)

with tab4:
    comp_name = st.text_input("ชื่อบริษัทผู้ขาย", value=st.session_state.form_comp_name)
    comp_doc_title = st.text_input("หัวข้อเอกสาร", value=st.session_state.form_comp_doc_title)

# การบันทึก
if st.button("💾 บันทึกข้อมูล", type="primary", use_container_width=True):
    with st.spinner("กำลังประมวลผล..."):
        new_no = next_inv_no()
        # บันทึก Invoices (Logic เดิม)
        ws_inv.append_row([new_no, invoice_date, customer, address, 0, 0, 0, 0, 0, "รอดำเนินการ", c_id, d_name, "ค้างชำระ", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", comp_name, "", "", "", comp_doc_title])
        
        # บันทึก InvoiceItems (บันทึก price และ amount เป็นข้อความ)
        for it in st.session_state.invoice_items:
            ws_item.append_row([new_no, it['product'], it['unit'], it['qty'], it['price'], it['amount']])
            
        st.success(f"บันทึกสำเร็จ เลขที่: {new_no}")
        st.cache_data.clear()
