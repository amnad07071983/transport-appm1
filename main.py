import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import io
import os

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors

# ================= 1. CONFIG & INITIALIZATION =================
st.set_page_config(page_title="Logistics System Pro", layout="wide")

# ลงทะเบียนฟอนต์ภาษาไทย (ต้องมีไฟล์ THSARABUN BOLD.ttf ในโฟลเดอร์เดียวกัน)
FONT_NAME = 'Helvetica-Bold'
try:
    if os.path.exists('THSARABUN BOLD.ttf'):
        pdfmetrics.registerFont(TTFont('ThaiFontBold', 'THSARABUN BOLD.ttf'))
        FONT_NAME = 'ThaiFontBold'
except:
    pass

SHEET_ID = "1fl86CxqgxlXAYU63GQOdCrL2jbPvSUdoXd1ndQvjnBM"
INV_SHEET = "Invoices"
ITEM_SHEET = "InvoiceItems"
INV_KEY = "invoice_no"
ITEM_KEY = "เลขที่บิล"

@st.cache_resource
def init_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    return gspread.authorize(creds).open_by_key(SHEET_ID)

@st.cache_data(ttl=5)
def get_data_cached():
    client = init_sheet()
    try:
        inv = client.worksheet(INV_SHEET).get_all_records()
        items = client.worksheet(ITEM_SHEET).get_all_records()
        return pd.DataFrame(inv), pd.DataFrame(items)
    except:
        return pd.DataFrame(), pd.DataFrame()

client = init_sheet()
inv_df, item_df = get_data_cached()
ws_inv = client.worksheet(INV_SHEET)
ws_item = client.worksheet(ITEM_SHEET)

transport_fields = [
    "ผู้รับสินค้า-ชื่อ", "ผู้รับสินค้า-ที่อยู่", "ผู้รับสินค้า-เลขผู้เสียภาษี", "ผู้รับสินค้า-เบอร์โทร",
    "คลังรับผลิตภัณฑ์-ชื่อ", "คลังรับผลิตภัณฑ์-เลขผู้เสียภาษี", "คลังรับผลิตภัณฑ์-ที่อยู่",
    "ผู้รับผลิตภัณฑ์-ชื่อ", "ผู้รับผลิตภัณฑ์-เลขผู้เสียภาษี", "ผู้รับผลิตภัณฑ์-ที่อยู่", "ผู้รับผลิตภัณฑ์-หมายเลขตั๋ว",
    "ผู้ดำเนินการขนส่ง-ชื่อ", "ผู้ดำเนินการขนส่ง-เลขผู้เสียภาษี", "ผู้ดำเนินการขนส่ง-ที่อยู่", "ผู้ดำเนินการขนส่ง-เบอร์โทร",
    "ผู้ดำเนินการขนส่ง-ประเภทผู้รับจ้าง", "ผู้ดำเนินการขนส่ง-ใบอนุญาต",
    "ข้อมูลพนักงานขับรถ-ชื่อ", "ข้อมูลพนักงานขับรถ-เลขใบขับขี่", "ข้อมูลพนักงานขับรถ-เบอร์โทร", "ข้อมูลพนักงานขับรถ-ทะเบียนรถ",
    "ข้อมูลพนักงานขับรถ-วิธีขนส่ง", "ข้อมูลพนักงานขับรถ-วันออกเดินทาง", "ข้อมูลพนักงานขับรถ-เวลาออกเดินทาง",
    "ข้อมูลพนักงานขับรถ-วันที่ถึงปลายทาง", "ข้อมูลพนักงานขับรถ-เวลาที่ถึงปลายทาง",
    "การยืนยันและรับสินค้า-ผู้ออกเอกสาร", "การยืนยันและรับสินค้า-พนักงานขับรถ", "การยืนยันและรับสินค้า-ผู้รับสินค้า",
    "ผู้จำหน่าย-ชื่อ", "ผู้จำหน่าย-ที่อยู่", "ผู้จำหน่าย-เลขผู้เสียภาษี", "ผู้จำหน่าย-เบอร์โทร",
    "ผู้จำหน่าย-ชื่อเอกสาร", "ผู้จำหน่าย-อธิบายเพิ่ม"
]

# ================= 2. SESSION STATE =================
if "invoice_items" not in st.session_state:
    st.session_state.invoice_items = []
if "editing_no" not in st.session_state:
    st.session_state.editing_no = None
if "pdf_buffer" not in st.session_state:
    st.session_state.pdf_buffer = None
if "form_date" not in st.session_state:
    st.session_state.form_date = datetime.now().strftime("%d/%m/%Y")

for f in transport_fields:
    if f"in_{f}" not in st.session_state:
        st.session_state[f"in_{f}"] = ""

def reset_form_action():
    st.session_state.invoice_items = []
    st.session_state.editing_no = None
    st.session_state.pdf_buffer = None
    st.session_state.form_date = datetime.now().strftime("%d/%m/%Y")
    for f in transport_fields:
        st.session_state[f"in_{f}"] = ""

# ================= 3. PDF GENERATOR (ALL FIELDS) =================
def generate_pdf_file(inv_no, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    if os.path.exists("p1.png"):
        c.drawImage("p1.png", 0, 0, width=w, height=h)
    
    # ส่วนหัว
    c.setFont(FONT_NAME, 16)
    c.drawCentredString(w/2, h-1.5*cm, f"{st.session_state.get('in_ผู้จำหน่าย-ชื่อเอกสาร', 'ใบกำกับขนส่ง')}")
    
    c.setFont(FONT_NAME, 11)
    c.drawString(2*cm, h-2.5*cm, f"ผู้จำหน่าย: {st.session_state.get('in_ผู้จำหน่าย-ชื่อ', '')}")
    c.drawString(2*cm, h-3.0*cm, f"ที่อยู่: {st.session_state.get('in_ผู้จำหน่าย-ที่อยู่', '')}")
    c.drawRightString(19*cm, h-2.5*cm, f"เลขที่บิล: {inv_no}")
    c.drawRightString(19*cm, h-3.0*cm, f"วันที่: {st.session_state.get('form_date', '')}")

    # ข้อมูลต้นทาง-ปลายทาง
    c.line(2*cm, h-3.8*cm, 19*cm, h-3.8*cm)
    curr_y = h-4.3*cm
    c.drawString(2*cm, curr_y, f"ต้นทาง: {st.session_state.get('in_คลังรับผลิตภัณฑ์-ชื่อ', '')} | ภาษี: {st.session_state.get('in_คลังรับผลิตภัณฑ์-เลขผู้เสียภาษี', '')}")
    c.drawString(2*cm, curr_y-0.4*cm, f"ที่อยู่คลัง: {st.session_state.get('in_คลังรับผลิตภัณฑ์-ที่อยู่', '')}")
    
    c.drawString(11*cm, curr_y, f"ผู้รับสินค้า: {st.session_state.get('in_ผู้รับสินค้า-ชื่อ', '')}")
    c.drawString(11*cm, curr_y-0.4*cm, f"ภาษีผู้รับ: {st.session_state.get('in_ผู้รับสินค้า-เลขผู้เสียภาษี', '')}")
    c.drawString(11*cm, curr_y-0.8*cm, f"ที่อยู่ผู้รับ: {st.session_state.get('in_ผู้รับสินค้า-ที่อยู่', '')}")

    # ข้อมูลขนส่ง
    curr_y -= 1.8*cm
    c.rect(2*cm, curr_y-2.2*cm, 17*cm, 2.6*cm)
    c.drawString(2.2*cm, curr_y, f"ขนส่งโดย: {st.session_state.get('in_ผู้ดำเนินการขนส่ง-ชื่อ', '')} | ใบอนุญาต: {st.session_state.get('in_ผู้ดำเนินการขนส่ง-ใบอนุญาต', '')}")
    c.drawString(2.2*cm, curr_y-0.5*cm, f"พนักงานขับ: {st.session_state.get('in_ข้อมูลพนักงานขับรถ-ชื่อ', '')} | ทะเบียน: {st.session_state.get('in_ข้อมูลพนักงานขับรถ-ทะเบียนรถ', '')}")
    c.drawString(2.2*cm, curr_y-1.0*cm, f"กำหนดเวลาเดินทาง: {st.session_state.get('in_ข้อมูลพนักงานขับรถ-วันออกเดินทาง', '')} {st.session_state.get('in_ข้อมูลพนักงานขับรถ-เวลาออกเดินทาง', '')}")
    c.drawString(2.2*cm, curr_y-1.5*cm, f"ถึงปลายทาง: {st.session_state.get('in_ข้อมูลพนักงานขับรถ-วันที่ถึงปลายทาง', '')} {st.session_state.get('in_ข้อมูลพนักงานขับรถ-เวลาที่ถึงปลายทาง', '')}")

    # ตารางสินค้า
    curr_y -= 3.2*cm
    header = [["ลำดับ", "รายการสินค้า", "หน่วย", "จำนวน", "ช่องถัง", "ซีล"]]
    data = [[i+1, it['product'], it['unit'], it['qty'], it['price'], it['amount']] for i, it in enumerate(items)]
    while len(data) < 5: data.append(["", "", "", "", "", ""])
    t = Table(header + data, colWidths=[1.2*cm, 7.3*cm, 1.5*cm, 2*cm, 2.5*cm, 2.5*cm])
    t.setStyle(TableStyle([('FONT', (0,0), (-1,-1), FONT_NAME, 10), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.whitesmoke)]))
    t.wrapOn(c, 2*cm, curr_y)
    t.drawOn(c, 2*cm, curr_y - (len(data)*0.55*cm))

    # ส่วนท้าย (ลายเซ็น)
    footer_y = 3*cm
    c.drawString(2*cm, footer_y, f"ผู้ดำเนินการ: {st.session_state.get('in_การยืนยันและรับสินค้า-ผู้ออกเอกสาร', '..........')}")
    c.drawString(8*cm, footer_y, f"พนักงานขับ: {st.session_state.get('in_การยืนยันและรับสินค้า-พนักงานขับรถ', '..........')}")
    c.drawString(14*cm, footer_y, f"ผู้รับสินค้า: {st.session_state.get('in_การยืนยันและรับสินค้า-ผู้รับสินค้า', '..........')}")
    
    c.showPage(); c.save(); buf.seek(0)
    return buf

def next_inv_no():
    prefix = f"INV-{datetime.now().year}-{datetime.now().month:02d}"
    if inv_df.empty: return f"{prefix}-0001"
    curr = inv_df[inv_df[INV_KEY].astype(str).str.startswith(prefix)]
    if curr.empty: return f"{prefix}-0001"
    last_seq = int(str(curr[INV_KEY].iloc[-1]).split('-')[-1])
    return f"{prefix}-{last_seq + 1:04d}"

# ================= 4. MAIN UI =================
st.title("🚚 M POWER OIL - ระบบออกใบกำกับ")

col_top1, col_top2 = st.columns([6, 2])
with col_top2:
    if st.button("🆕 เริ่มบิลใหม่", use_container_width=True):
        reset_form_action()
        st.rerun()

with st.expander("🔍 ค้นหา จัดการข้อมูลเก่า (แก้ไข/สร้างซ้ำ)"):
    if not inv_df.empty:
        options = [f"{r[INV_KEY]} | {r.get('ผู้รับสินค้า-ชื่อ', '')}" for _, r in inv_df.iterrows()]
        selected_bill = st.selectbox("เลือกบิล", [""] + options[::-1])
        if selected_bill:
            sel_no = selected_bill.split(" | ")[0]
            c_btn1, c_btn2 = st.columns(2)
            if c_btn1.button("📝 แก้ไขบิลนี้"):
                old_inv = inv_df[inv_df[INV_KEY] == sel_no].iloc[0].to_dict()
                st.session_state.editing_no = sel_no
                st.session_state.form_date = str(old_inv.get('date', ''))
                for f in transport_fields: st.session_state[f"in_{f}"] = str(old_inv.get(f, ""))
                it_rows = item_df[item_df[ITEM_KEY if ITEM_KEY in item_df.columns else "invoice_no"] == sel_no].to_dict('records')
                st.session_state.invoice_items = [{"product": i.get('รายการ', i.get('product','')), "unit": i.get('หน่วย', i.get('unit','')), "qty": i.get('จำนวน', i.get('qty','')), "price": str(i.get('หมายเลขช่องถัง', i.get('price',''))), "amount": str(i.get('หมายเลขซีล', i.get('amount','')))} for i in it_rows]
                st.rerun()
            if c_btn2.button("🔄 สร้างซ้ำจากบิลนี้"):
                old_inv = inv_df[inv_df[INV_KEY] == sel_no].iloc[0].to_dict()
                st.session_state.editing_no = None
                for f in transport_fields: st.session_state[f"in_{f}"] = str(old_inv.get(f, ""))
                it_rows = item_df[item_df[ITEM_KEY if ITEM_KEY in item_df.columns else "invoice_no"] == sel_no].to_dict('records')
                st.session_state.invoice_items = [{"product": i.get('รายการ', i.get('product','')), "unit": i.get('หน่วย', i.get('unit','')), "qty": i.get('จำนวน', i.get('qty','')), "price": str(i.get('หมายเลขช่องถัง', i.get('price',''))), "amount": str(i.get('หมายเลขซีล', i.get('amount','')))} for i in it_rows]
                st.rerun()

st.divider()
t1, t2, t3, t4 = st.tabs(["📦 สินค้า & ปลายทาง", "🚛 ขนส่ง", "⛽ รายการ", "🏢 ผู้จัดจำหน่าย"])
with t1:
    c1, c2 = st.columns(2)
    with c1:
        for f in transport_fields[0:4]: st.text_input(f, key=f"in_{f}")
    with c2:
        for f in transport_fields[4:11]: st.text_input(f, key=f"in_{f}")
with t2:
    c1, c2 = st.columns(2)
    with c1:
        for f in transport_fields[11:17]: st.text_input(f, key=f"in_{f}")
    with c2:
        for f in transport_fields[17:26]: st.text_input(f, key=f"in_{f}")
with t3:
    ca, cb, cc, cd, ce = st.columns([3,1,1,2,2])
    p_n = ca.text_input("รายการ", key="tmp_p")
    p_u = cb.text_input("หน่วย", value="ลิตร", key="tmp_u")
    p_q = cc.text_input("จำนวน", key="tmp_q")
    p_t = cd.text_input("ช่องถัง", key="tmp_t")
    p_s = ce.text_input("ซีล", key="tmp_s")
    if st.button("➕ เพิ่มสินค้า"):
        if p_n and p_q:
            st.session_state.invoice_items.append({"product": p_n, "unit": p_u, "qty": p_q, "price": p_t, "amount": p_s})
            st.rerun()
    for idx, item in enumerate(st.session_state.invoice_items):
        cx = st.columns([5, 1])
        cx[0].write(f"{idx+1}. {item['product']} | {item['qty']} {item['unit']} | ถัง: {item['price']} ซีล: {item['amount']}")
        if cx[1].button("🗑️", key=f"del_{idx}"):
            st.session_state.invoice_items.pop(idx); st.rerun()
with t4:
    c1, c2 = st.columns(2)
    with c1:
        for f in transport_fields[26:29]: st.text_input(f, key=f"in_{f}")
    with c2:
        st.session_state.form_date = st.text_input("วันที่", value=st.session_state.form_date)
        for f in transport_fields[29:]: st.text_input(f, key=f"in_{f}")

st.divider()
if st.session_state.pdf_buffer:
    st.download_button("📥 ดาวน์โหลดไฟล์ PDF เพื่อสั่งพิมพ์", data=st.session_state.pdf_buffer, file_name="Invoice.pdf", mime="application/pdf", use_container_width=True)

if st.button("💾 บันทึกข้อมูลและสร้างไฟล์ PDF", type="primary", use_container_width=True):
    if not st.session_state.invoice_items:
        st.error("กรุณาเพิ่มรายการสินค้า")
    else:
        new_no = st.session_state.editing_no if st.session_state.editing_no else next_inv_no()
        if st.session_state.editing_no: # ลบเก่าเพื่อทับใหม่
            try:
                for ws in [ws_inv, ws_item]:
                    found = ws.findall(new_no)
                    for f in reversed(found): ws.delete_rows(f.row)
            except: pass
        inv_row = [new_no, st.session_state.form_date] + [st.session_state[f"in_{f}"] for f in transport_fields]
        ws_inv.append_row(inv_row)
        for it in st.session_state.invoice_items:
            ws_item.append_row([new_no, it['product'], it['unit'], it['qty'], it['price'], it['amount']])
        st.session_state.pdf_buffer = generate_pdf_file(new_no, st.session_state.invoice_items)
        st.session_state.editing_no = None
        st.cache_data.clear()
        st.rerun()
