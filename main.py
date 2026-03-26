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
st.set_page_config(page_title="M POWER OIL - Logistics System", layout="wide")

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

# รายชื่อฟิลด์ทั้งหมด (ต้องตรงกับหัวตารางใน Google Sheet)
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
if "invoice_items" not in st.session_state: st.session_state.invoice_items = []
if "editing_no" not in st.session_state: st.session_state.editing_no = None
if "pdf_buffer" not in st.session_state: st.session_state.pdf_buffer = None
if "form_date" not in st.session_state: st.session_state.form_date = datetime.now().strftime("%d/%m/%Y")

for f in transport_fields:
    if f"in_{f}" not in st.session_state: st.session_state[f"in_{f}"] = ""

def reset_form_action():
    st.session_state.invoice_items = []
    st.session_state.editing_no = None
    st.session_state.pdf_buffer = None
    st.session_state.form_date = datetime.now().strftime("%d/%m/%Y")
    for f in transport_fields: st.session_state[f"in_{f}"] = ""

# ================= 3. PDF GENERATOR =================
def generate_pdf_file(inv_no, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    
    # --- Header ---
    c.setFont(FONT_NAME, 11)
    c.drawString(1.5*cm, h-1.5*cm, f"{st.session_state.get('in_ผู้จำหน่าย-ชื่อ', '')}")
    c.drawString(1.5*cm, h-2.0*cm, f"{st.session_state.get('in_ผู้จำหน่าย-ที่อยู่', '')}")
    c.drawString(1.5*cm, h-2.5*cm, f"โทร.{st.session_state.get('in_ผู้จำหน่าย-เบอร์โทร', '')}")
    c.drawString(1.5*cm, h-3.0*cm, f"เลขประจำตัวผู้เสียภาษี {st.session_state.get('in_ผู้จำหน่าย-เลขผู้เสียภาษี', '')}")

    # ฟิลด์ ผู้จำหน่าย-ชื่อเอกสาร ไว้ขวาสุดบรรทัดเดียวกับชื่อบริษัท
    c.setFont(FONT_NAME, 14)
    c.drawRightString(19.5*cm, h-1.5*cm, f"{st.session_state.get('in_ผู้จำหน่าย-ชื่อเอกสาร', 'ใบกำกับขนส่งน้ำมัน')}")
    
    c.setFont(FONT_NAME, 10)
    c.drawRightString(19.5*cm, h-2.0*cm, "(ตามประกาศกระทรวงพาณิชย์ และกรมธุรกิจพลังงาน)")
    c.drawString(13*cm, h-2.7*cm, f"เลขที่ : {inv_no}")
    c.drawString(13*cm, h-3.2*cm, f"วันที่ : {st.session_state.get('form_date', '')}")

    c.line(1*cm, h-3.5*cm, 20*cm, h-3.5*cm)

    # 1. ข้อมูลคู่ค้า
    c.setFont(FONT_NAME, 11)
    c.drawString(1.2*cm, h-4.2*cm, "1. ข้อมูลคู่ค้า")
    c.drawString(1.5*cm, h-4.8*cm, f"1.1 คลังรับผลิตภัณฑ์ : {st.session_state.get('in_คลังรับผลิตภัณฑ์-ชื่อ', '')}")
    c.drawString(1.5*cm, h-5.3*cm, f"ที่อยู่ : {st.session_state.get('in_คลังรับผลิตภัณฑ์-ที่อยู่', '')}")
    c.drawString(11*cm, h-4.8*cm, f"1.2 ผู้รับผลิตภัณฑ์ : {st.session_state.get('in_ผู้รับผลิตภัณฑ์-ชื่อ', '')}")
    c.drawString(11*cm, h-5.3*cm, f"ที่อยู่ : {st.session_state.get('in_ผู้รับผลิตภัณฑ์-ที่อยู่', '')}")
    c.drawString(1.5*cm, h-6.8*cm, f"1.3 ผู้รับสินค้า (ปลายทาง) : {st.session_state.get('in_ผู้รับสินค้า-ชื่อ', '')}")
    c.drawString(1.5*cm, h-7.3*cm, f"ที่อยู่ : {st.session_state.get('in_ผู้รับสินค้า-ที่อยู่', '')}")
    c.drawString(11*cm, h-7.3*cm, f"ตั๋วขนย้ายเลขที่ : {st.session_state.get('in_ผู้รับผลิตภัณฑ์-หมายเลขตั๋ว', '')}")

    c.line(1*cm, h-8.5*cm, 20*cm, h-8.5*cm)

    # 2. การขนส่ง
    c.drawString(1.2*cm, h-9.0*cm, "2. ข้อมูลการขนส่ง")
    c.drawString(1.5*cm, h-9.6*cm, f"2.1 ผู้ดำเนินการขนส่ง : {st.session_state.get('in_ผู้ดำเนินการขนส่ง-ชื่อ', '')}")
    c.drawString(1.5*cm, h-10.1*cm, f"ใบอนุญาต : {st.session_state.get('in_ผู้ดำเนินการขนส่ง-ใบอนุญาต', '')}")
    c.drawString(11*cm, h-9.6*cm, f"2.2 พนักงานขับรถ : {st.session_state.get('in_ข้อมูลพนักงานขับรถ-ชื่อ', '')}")
    c.drawString(11*cm, h-10.1*cm, f"ทะเบียนรถ : {st.session_state.get('in_ข้อมูลพนักงานขับรถ-ทะเบียนรถ', '')}")

    # 3. ตาราง
    header = [["ลำดับ", "ช่องถัง", "ซีล", "รายการน้ำมัน", "หน่วย", "จำนวน"]]
    data_rows = [[i+1, it['tank'], it['seal'], it['product'], it['unit'], it['qty']] for i, it in enumerate(items)]
    while len(data_rows) < 4: data_rows.append(["","","","","",""])
    t = Table(header + data_rows, colWidths=[1.2*cm, 2.5*cm, 3.5*cm, 6.8*cm, 2*cm, 3*cm])
    t.setStyle(TableStyle([('FONT', (0,0), (-1,-1), FONT_NAME, 10), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    t.wrapOn(c, 1*cm, h-16*cm)
    t.drawOn(c, 1*cm, h-16*cm)

    # 4. ลายเซ็น
    sig_y = h-23*cm
    c.drawCentredString(4.5*cm, sig_y, "..................................")
    c.drawCentredString(4.5*cm, sig_y-0.5*cm, f"( {st.session_state.get('in_การยืนยันและรับสินค้า-ผู้ออกเอกสาร', '')} )")
    c.drawCentredString(10.5*cm, sig_y, "..................................")
    c.drawCentredString(10.5*cm, sig_y-0.5*cm, f"( {st.session_state.get('in_การยืนยันและรับสินค้า-พนักงานขับรถ', '')} )")
    c.drawCentredString(16.5*cm, sig_y, "..................................")
    c.drawCentredString(16.5*cm, sig_y-0.5*cm, f"( {st.session_state.get('in_การยืนยันและรับสินค้า-ผู้รับสินค้า', '')} )")

    c.rect(1*cm, 1*cm, 19*cm, h-2*cm)
    c.showPage(); c.save(); buf.seek(0)
    return buf

# ================= 4. MAIN UI =================
st.title("🚚 LOGISTICS SYSTEM")

with st.expander("🔍 ค้นหา/แก้ไข/สร้างซ้ำ"):
    if not inv_df.empty:
        options = [f"{r[INV_KEY]} | {r.get('ผู้รับสินค้า-ชื่อ', '')}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกบิล", [""] + options[::-1])
        if selected:
            sel_no = selected.split(" | ")[0]
            col_a, col_b = st.columns(2)
            if col_a.button("📝 โหลดมาแก้ไข"):
                # ดึงข้อมูล Row ที่เลือกมาใส่ Session State
                row_data = inv_df[inv_df[INV_KEY] == sel_no].iloc[0].to_dict()
                st.session_state.editing_no = sel_no
                st.session_state.form_date = str(row_data.get('date', st.session_state.form_date))
                for f in transport_fields:
                    # แก้ไขจุดสำคัญ: ใช้ชื่อฟิลด์จากลิสต์ transport_fields ในการดึงข้อมูลจาก Sheet
                    st.session_state[f"in_{f}"] = str(row_data.get(f, ""))
                
                # ดึงรายการสินค้า
                it_rows = item_df[item_df["invoice_no"] == sel_no].to_dict('records')
                st.session_state.invoice_items = [{"product": i.get('product',''), "unit": i.get('unit',''), "qty": i.get('qty',''), "tank": str(i.get('tank','')), "seal": str(i.get('seal',''))} for i in it_rows]
                st.rerun()
                
            if col_b.button("🔄 โหลดมาสร้างซ้ำ"):
                row_data = inv_df[inv_df[INV_KEY] == sel_no].iloc[0].to_dict()
                st.session_state.editing_no = None
                for f in transport_fields:
                    st.session_state[f"in_{f}"] = str(row_data.get(f, ""))
                it_rows = item_df[item_df["invoice_no"] == sel_no].to_dict('records')
                st.session_state.invoice_items = [{"product": i.get('product',''), "unit": i.get('unit',''), "qty": i.get('qty',''), "tank": str(i.get('tank','')), "seal": str(i.get('seal',''))} for i in it_rows]
                st.rerun()

tabs = st.tabs(["📦 คู่ค้า", "🚛 ขนส่ง", "⛽ สินค้า", "🏢 บริษัท"])
with tabs[0]:
    for f in transport_fields[0:11]: st.text_input(f, key=f"in_{f}")
with tabs[1]:
    for f in transport_fields[11:26]: st.text_input(f, key=f"in_{f}")
with tabs[2]:
    ca, cb, cc, cd, ce = st.columns([3,1,1,2,2])
    p_n = ca.text_input("รายการ", key="t_n")
    p_u = cb.text_input("หน่วย", value="ลิตร", key="t_u")
    p_q = cc.text_input("จำนวน", key="t_q")
    p_p = cd.text_input("ช่องถัง", key="t_p")
    p_a = ce.text_input("ซีล", key="t_a")
    if st.button("➕ เพิ่ม"):
        st.session_state.invoice_items.append({"product":p_n, "unit":p_u, "qty":p_q, "price":p_p, "amount":p_a})
        st.rerun()
    st.write(st.session_state.invoice_items)
with tabs[3]:
    st.session_state.form_date = st.text_input("วันที่", value=st.session_state.form_date)
    for f in transport_fields[26:]: st.text_input(f, key=f"in_{f}")

if st.button("💾 บันทึกและออก PDF", type="primary", use_container_width=True):
    def get_next_no():
        prefix = f"INV-{datetime.now().year}-{datetime.now().month:02d}"
        if inv_df.empty: return f"{prefix}-0001"
        curr = inv_df[inv_df[INV_KEY].astype(str).str.startswith(prefix)]
        if curr.empty: return f"{prefix}-0001"
        last_val = str(curr[INV_KEY].iloc[-1]).split('-')[-1]
        return f"{prefix}-{int(last_val)+1:04d}"

    final_no = st.session_state.editing_no if st.session_state.editing_no else get_next_no()
    
    if st.session_state.editing_no:
        try:
            for ws in [ws_inv, ws_item]:
                found = ws.findall(final_no)
                for cell in reversed(found): ws.delete_rows(cell.row)
        except: pass

    # บันทึกข้อมูล (Column 1=ID, Column 2=Date, ต่อด้วย transport_fields)
    ws_inv.append_row([final_no, st.session_state.form_date] + [st.session_state[f"in_{f}"] for f in transport_fields])
    for it in st.session_state.invoice_items:
        ws_item.append_row([final_no, it['product'], it['unit'], it['qty'], it['price'], it['amount']])
    
    st.session_state.pdf_buffer = generate_pdf_file(final_no, st.session_state.invoice_items)
    st.session_state.editing_no = None
    st.cache_data.clear()
    st.rerun()

if st.session_state.pdf_buffer:
    st.download_button("📥 Download PDF", data=st.session_state.pdf_buffer, file_name="Invoice.pdf", mime="application/pdf")
