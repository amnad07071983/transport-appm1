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
    """p2.png พื้นหลังเต็มหน้า (ความเข้มปกติ) และ p1.png ลายน้ำกลางหน้า"""
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
            c.drawImage("p1.png", (w-img_w)/2, (h-img_h)/2 - (1.5*inch), width=img_w, height=img_h, mask='auto', preserveAspectRatio=True)
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
    last_seq = int(str(current_month["invoice_no"].iloc[-1]).split('-')[-1])
    return f"{prefix}-{last_seq + 1:04d}"

def create_pdf(inv, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    add_single_watermark(c, w, h)
    
    # Header
    c.setFont("ThaiFontBold", 24) 
    c.drawString(2*cm, h-1.5*cm, str(inv.get('comp_name', '')))
    c.setFont("ThaiFontBold", 14)
    c.drawString(2*cm, h-2.3*cm, f"ที่อยู่: {inv.get('comp_address', '')}")
    c.drawRightString(19*cm, h-1.5*cm, str(inv.get('comp_doc_title', 'ใบกำกับขนส่ง')))
    c.drawRightString(19*cm, h-2.4*cm, f"เลขที่: {inv.get('invoice_no','')}")
    
    # ข้อมูลการขนส่ง (ตารางย่อย)
    transport_data = [
        [f"ทะเบียนรถ: {inv.get('car_id','')}", f"ออก: {inv.get('date_out','')} {inv.get('time_out','')}"],
        [f"ชื่อคนขับ: {inv.get('driver_name','')}", f"เข้า: {inv.get('date_in','')} {inv.get('time_in','')}"]
    ]
    t_trans = Table(transport_data, colWidths=[8*cm, 9*cm])
    t_trans.setStyle(TableStyle([('FONT', (0,0), (-1,-1), 'ThaiFontBold', 12)]))
    t_trans.wrapOn(c, 2*cm, h-8*cm)
    t_trans.drawOn(c, 2*cm, h-8*cm)

    # ตารางรายการสินค้า (ปรับหัวตารางใหม่)
    header = [["ลำดับ", "รายการสินค้า", "หน่วย", "จำนวน", "หมายเลขช่องถัง", "หมายเลขซีล"]]
    rows = []
    for i, it in enumerate(items):
        rows.append([i+1, it.get("product", ""), it.get("unit", ""), it.get("qty", ""), it.get("price", ""), it.get("amount", "")])
    
    t_items = Table(header + rows, colWidths=[1.2*cm, 7.3*cm, 1.5*cm, 2*cm, 2.5*cm, 2.5*cm])
    t_items.setStyle(TableStyle([
        ('FONT', (0,0), (-1,0), 'ThaiFontBold', 13),
        ('FONT', (0,1), (-1,-1), 'ThaiFontBold', 12),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    tw, th = t_items.wrapOn(c, 2*cm, h-16*cm)
    t_items.drawOn(c, 2*cm, h-10*cm-th)
    
    c.showPage()
    c.save()
    buf.seek(0)
    return buf

# ================= 4. MAIN UI =================
st.markdown("## 🚚 ระบบจัดการขนส่ง M POWER OIL")

# --- ส่วนจัดการประวัติ ---
with st.expander("🔍 ค้นหาและจัดการประวัติ"):
    if not inv_df.empty:
        options = [f"{r['invoice_no']} | {r['customer']}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกบิล", [""] + options[::-1])
        if selected:
            sel_no = selected.split(" | ")[0]
            old_inv = inv_df[inv_df["invoice_no"] == sel_no].iloc[0].to_dict()
            old_items = item_df[item_df["invoice_no"] == sel_no].to_dict('records')
            st.download_button(f"📥 ดาวน์โหลด PDF {sel_no}", create_pdf(old_inv, old_items), f"{sel_no}.pdf")

st.divider()

# --- ส่วนรับข้อมูล ---
tab1, tab2, tab3 = st.tabs(["👤 ข้อมูลลูกค้า/บริษัท", "🚛 ขนส่ง", "📦 รายการสินค้า (ช่องถัง/ซีล)"])

with tab1:
    customer = st.text_input("ชื่อลูกค้า", value=st.session_state.form_customer)
    address = st.text_area("ที่อยู่", value=st.session_state.form_address)
    comp_name = st.text_input("ชื่อบริษัท (หัว PDF)", value=st.session_state.form_comp_name)
    comp_doc_title = st.text_input("ชื่อประเภทเอกสาร", value="ใบกำกับขนส่ง")

with tab2:
    col1, col2 = st.columns(2)
    car_id = col1.text_input("ทะเบียนรถ", value=st.session_state.form_car_id)
    driver_name = col1.text_input("ชื่อคนขับ", value=st.session_state.form_driver_name)
    date_out = col2.text_input("วันที่ออก", value=st.session_state.form_date_out)
    time_out = col2.text_input("เวลาออก", value=st.session_state.form_time_out)

with tab3:
    st.info("💡 เพิ่มข้อมูลรายการสินค้า พร้อมระบุหมายเลขช่องถังและซีล")
    ci1, ci2, ci3, ci4, ci5 = st.columns([3, 1, 1, 2, 2])
    p_name = ci1.text_input("รายการสินค้า")
    p_unit = ci2.text_input("หน่วย", value="ลิตร")
    p_qty = ci3.text_input("จำนวน")
    # 1.1 เปลี่ยนชื่อเป็น "หมายเลขช่องถัง" (Price ใน Sheet)
    p_tank = ci4.text_input("หมายเลขช่องถัง") 
    # 1.2 เปลี่ยนชื่อเป็น "หมายเลขซีล" (Amount ใน Sheet)
    p_seal = ci5.text_input("หมายเลขซีล")

    if st.button("➕ เพิ่มรายการ"):
        if p_name:
            st.session_state.invoice_items.append({
                "product": p_name, 
                "unit": p_unit, 
                "qty": p_qty, 
                "price": p_tank, # บันทึกลงคอลัมน์ Price
                "amount": p_seal # บันทึกลงคอลัมน์ Amount
            })
            st.rerun()

    if st.session_state.invoice_items:
        df_items = pd.DataFrame(st.session_state.invoice_items)
        # เปลี่ยนชื่อหัวตารางในหน้าแอป
        df_display = df_items.rename(columns={
            "product": "รายการ", "unit": "หน่วย", "qty": "จำนวน", 
            "price": "หมายเลขช่องถัง", "amount": "หมายเลขซีล"
        })
        st.table(df_display)
        if st.button("🧹 ล้างรายการสินค้า"):
            st.session_state.invoice_items = []
            st.rerun()

# --- บันทึกข้อมูล ---
if st.button("💾 บันทึกและออกเอกสาร", type="primary", use_container_width=True):
    if not customer or not st.session_state.invoice_items:
        st.error("กรุณากรอกข้อมูลลูกค้าและเพิ่มรายการสินค้าอย่างน้อย 1 รายการ")
    else:
        with st.spinner("กำลังบันทึกข้อมูล..."):
            new_no = next_inv_no()
            inv_date = datetime.now().strftime("%d/%m/%Y")
            
            # เตรียมข้อมูลบันทึกลง Sheet Invoices
            inv_data = [
                new_no, inv_date, customer, address, 0, 0, 0, 0, 0, 
                "รอดำเนินการ", car_id, driver_name, "ค้างชำระ", date_out, time_out, 
                "", "", "", "", "", "", "", "", "", "", "", "", "", 
                comp_name, "", "", "", comp_doc_title
            ]
            ws_inv.append_row(inv_data)
            
            # เตรียมข้อมูลบันทึกลง Sheet InvoiceItems
            for it in st.session_state.invoice_items:
                ws_item.append_row([
                    new_no, 
                    it['product'], 
                    it['unit'], 
                    it['qty'], 
                    str(it['price']),  # บันทึกเป็นข้อความ (หมายเลขช่องถัง)
                    str(it['amount'])  # บันทึกเป็นข้อความ (หมายเลขซีล)
                ])
            
            st.success(f"บันทึกสำเร็จ! เลขที่เอกสาร: {new_no}")
            st.cache_data.clear()
            # แสดงปุ่มโหลด PDF ทันทีหลังบันทึก
            st.download_button("📥 ดาวน์โหลดใบกำกับขนส่ง (PDF)", 
                             create_pdf({"invoice_no": new_no, "comp_name": comp_name, "comp_doc_title": comp_doc_title, "car_id": car_id, "driver_name": driver_name, "date_out": date_out, "time_out": time_out}, st.session_state.invoice_items), 
                             f"{new_no}.pdf")
