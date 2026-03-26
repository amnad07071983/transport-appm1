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

inv_df, item_df = get_data_cached()
client = init_sheet()
ws_inv = client.worksheet(INV_SHEET)
ws_item = client.worksheet(ITEM_SHEET)

# ฟิลด์ทั้งหมดที่ต้องมี (ห้ามหาย)
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
    for f in transport_fields:
        st.session_state[f"form_{f}"] = ""
    st.session_state.form_doc_status = "รอดำเนินการ"
    st.session_state.form_payment_status = "ค้างชำระ"

if "invoice_items" not in st.session_state:
    reset_form()

# ================= 3. CORE FUNCTIONS (PDF & LOGIC) =================

def add_single_watermark(c, w, h):
    try:
        if os.path.exists("p2.png"):
            c.saveState()
            c.setFillAlpha(1.0) 
            c.drawImage("p2.png", 0, 0, width=w, height=h, mask='auto', preserveAspectRatio=False)
            c.restoreState()
    except: pass

def next_inv_no():
    # ดึงเลขที่บิลล่าสุด
    data = client.worksheet(INV_SHEET).get_all_records()
    df = pd.DataFrame(data)
    now = datetime.now()
    prefix = f"INV-{now.year}-{now.month:02d}"
    if df.empty or "invoice_no" not in df.columns: return f"{prefix}-0001"
    curr = df[df["invoice_no"].astype(str).str.startswith(prefix)]
    if curr.empty: return f"{prefix}-0001"
    last_seq = int(str(curr["invoice_no"].iloc[-1]).split('-')[-1])
    return f"{prefix}-{last_seq + 1:04d}"

def create_pdf(inv, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    add_single_watermark(c, w, h)
    
    # วาด Header (ใช้ข้อมูลบริษัทจากฟอร์ม)
    c.setFont("ThaiFontBold", 20)
    c.drawString(2*cm, h-1.5*cm, str(inv.get('comp_name', '')))
    c.setFont("ThaiFontBold", 24)
    c.drawRightString(19*cm, h-1.5*cm, str(inv.get('comp_doc_title', 'ใบกำกับขนส่ง')))
    
    # รายละเอียดลูกค้า & เลขที่บิล
    c.setFont("ThaiFontBold", 14)
    c.drawString(2*cm, h-2.5*cm, f"ลูกค้า: {inv.get('customer','')}")
    c.drawRightString(19*cm, h-2.5*cm, f"เลขที่: {inv.get('invoice_no','')}")
    
    # ตารางสินค้า (หัวตารางตามคำสั่ง)
    header = [["ลำดับ", "รายการ", "หน่วย", "จำนวน", "หมายเลขช่องถัง", "หมายเลขซีล"]]
    rows = [[i+1, it['product'], it['unit'], it['qty'], it['price'], it['amount']] for i, it in enumerate(items)]
    t = Table(header + rows, colWidths=[1.2*cm, 8*cm, 1.5*cm, 1.5*cm, 2.9*cm, 2.9*cm])
    t.setStyle(TableStyle([('FONT', (0,0), (-1,-1), 'ThaiFontBold', 12), ('LINEBELOW', (0,0), (-1,0), 1, colors.black)]))
    t.wrapOn(c, 2*cm, h-10*cm)
    t.drawOn(c, 2*cm, h-10*cm)
    
    c.showPage(); c.save(); buf.seek(0)
    return buf

# ================= 4. MAIN UI =================
st.markdown("## 🚚 ระบบจัดการขนส่ง M POWER OIL")

# --- ค้นหา/แก้ไข/ทำซ้ำ (ดึงฟิลด์กลับมาให้ครบ) ---
with st.expander("🔍 ค้นหาและจัดการประวัติ"):
    if not inv_df.empty:
        options = [f"{r['invoice_no']} | {r['customer']}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกบิล", [""] + options[::-1])
        if selected:
            sel_no = selected.split(" | ")[0]
            old_inv = inv_df[inv_df["invoice_no"] == sel_no].iloc[0].to_dict()
            it_rows = item_df[item_df["invoice_no"] == sel_no].to_dict('records')
            
            f_items = [{"product": i['product'], "unit": i['unit'], "qty": i['qty'], "price": str(i['price']), "amount": str(i['amount'])} for i in it_rows]
            
            c1, c2 = st.columns(2)
            if c1.button("🔄 สร้างรายการซ้ำ"):
                reset_form()
                st.session_state.form_customer = old_inv.get("customer","")
                st.session_state.form_address = old_inv.get("address","")
                st.session_state.invoice_items = f_items
                for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                st.rerun()
            if c2.button("📝 แก้ไขบิล"):
                st.session_state.editing_no = sel_no
                st.session_state.form_customer = old_inv.get("customer","")
                st.session_state.form_address = old_inv.get("address","")
                st.session_state.invoice_items = f_items
                for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                st.rerun()

st.divider()

# --- FORM INPUT (ครบทุก Tab) ---
t1, t2, t3, t4 = st.tabs(["👤 ลูกค้า", "🚛 ขนส่ง", "📦 สินค้า", "🏢 บริษัท"])
with t1:
    customer = st.text_input("ชื่อลูกค้า", value=st.session_state.form_customer)
    address = st.text_area("ที่อยู่", value=st.session_state.form_address)
    invoice_date = st.text_input("วันที่", value=st.session_state.form_date)
with t2:
    col_a, col_b = st.columns(2)
    # วนลูปสร้าง Input สำหรับฟิลด์ขนส่งทั้งหมดที่มีใน transport_fields
    for i, f in enumerate(transport_fields[:14]): # ดึงมาโชว์บางส่วนในหน้าจอ
        if i % 2 == 0: col_a.text_input(f, value=st.session_state[f"form_{f}"], key=f"in_{f}")
        else: col_b.text_input(f, value=st.session_state[f"form_{f}"], key=f"in_{f}")
with t3:
    ca, cb, cc, cd, ce = st.columns([3,1,1,2,2])
    p_n = ca.text_input("รายการ")
    p_u = cb.text_input("หน่วย", value="ลิตร")
    p_q = cc.text_input("จำนวน")
    p_t = cd.text_input("หมายเลขช่องถัง")
    p_s = ce.text_input("หมายเลขซีล")
    if st.button("➕ เพิ่ม"):
        st.session_state.invoice_items.append({"product":p_n, "unit":p_u, "qty":p_q, "price":str(p_t), "amount":str(p_s)})
        st.rerun()
    st.table(pd.DataFrame(st.session_state.invoice_items))
with t4:
    # ข้อมูลบริษัท
    for f in transport_fields[19:]:
        st.text_input(f, value=st.session_state[f"form_{f}"], key=f"in_{f}")

# --- บันทึกข้อมูล (ลงครบทุกคอลัมน์) ---
if st.button("💾 บันทึกข้อมูล", type="primary"):
    with st.spinner("กำลังบันทึก..."):
        new_no = st.session_state.editing_no if st.session_state.editing_no else next_inv_no()
        
        # 1. ลบข้อมูลเก่ากรณีแก้ไข
        if st.session_state.editing_no:
            try:
                for ws in [ws_inv, ws_item]:
                    found = ws.find(new_no)
                    while found: ws.delete_rows(found.row); found = ws.find(new_no)
            except: pass

        # 2. เตรียมแถวสำหรับ Invoices (ดึงค่าจาก Input ทั้งหมด)
        # ลำดับ: No, Date, Customer, Address, Ship, Disc, Vat, Sub, Total, Status, ...ฟิลด์ที่เหลือ
        inv_row = [new_no, invoice_date, customer, address, 0, 0, 0, 0, 0]
        for f in transport_fields:
            val = st.session_state.get(f"in_{f}", st.session_state.get(f"form_{f}", ""))
            inv_row.append(val)
        ws_inv.append_row(inv_row)

        # 3. บันทึกรายการสินค้า (คอลัมน์ price=ช่องถัง, amount=ซีล เป็นข้อความ)
        for it in st.session_state.invoice_items:
            ws_item.append_row([new_no, it['product'], it['unit'], it['qty'], it['price'], it['amount']])
        
        st.success(f"บันทึกเลขที่ {new_no} เรียบร้อย!")
        st.download_button("📥 ดาวน์โหลด PDF", create_pdf({"invoice_no":new_no, "customer":customer, "comp_name":st.session_state.get("in_comp_name",""), "comp_doc_title":st.session_state.get("in_comp_doc_title","")}, st.session_state.invoice_items), f"{new_no}.pdf")
        st.cache_data.clear()
