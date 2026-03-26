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

# ================= 3. PDF GENERATOR (MATCHING IMAGE LAYOUT) =================
def generate_pdf_file(inv_no, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    
    # --- ส่วนหัว (Header) ---
    c.setFont(FONT_NAME, 11)
    
    # ฝั่งซ้าย: ข้อมูลบริษัท (บรรทัดแรกอยู่ที่ h-1.5*cm)
    c.drawString(1.5*cm, h-1.5*cm, f"{st.session_state.get('in_ผู้จำหน่าย-ชื่อ', '')}")
    c.drawString(1.5*cm, h-2.0*cm, f"{st.session_state.get('in_ผู้จำหน่าย-ที่อยู่', '')}")
    c.drawString(1.5*cm, h-2.5*cm, f"โทร.{st.session_state.get('in_ผู้จำหน่าย-เบอร์โทร', '')}")
    c.drawString(1.5*cm, h-3.0*cm, f"เลขประจำตัวผู้เสียภาษี {st.session_state.get('in_ผู้จำหน่าย-เลขผู้เสียภาษี', '')}")

    # ฝั่งขวาบนสุด: "ชื่อเอกสาร" (ระดับบรรทัดเดียวกับชื่อบริษัท h-1.5*cm)
    c.setFont(FONT_NAME, 14)
    c.drawRightString(19.5*cm, h-1.5*cm, f"{st.session_state.get('in_ผู้จำหน่าย-ชื่อเอกสาร', 'ใบกำกับขนส่งน้ำมัน')}")
    
    # ฝั่งขวาถัดลงมา: เลขที่และวันที่
    c.setFont(FONT_NAME, 10)
    c.drawRightString(19.5*cm, h-2.0*cm, f"(ตามประกาศกระทรวงพาณิชย์ และกรมธุรกิจพลังงาน)")
    c.drawString(13*cm, h-2.7*cm, f"เลขที่ : {inv_no}")
    c.drawString(13*cm, h-3.2*cm, f"วันที่ : {st.session_state.get('form_date', '')}")

    c.line(1*cm, h-3.5*cm, 20*cm, h-3.5*cm)

    # --- 1. ข้อมูลคู่ค้า ---
    c.setFont(FONT_NAME, 11)
    c.drawString(1.2*cm, h-4.2*cm, "1. ข้อมูลคู่ค้า")
    c.drawString(1.5*cm, h-4.8*cm, "1.1 คลังรับผลิตภัณฑ์ (ต้นทาง)")
    c.drawString(1.5*cm, h-5.3*cm, f"ชื่อ : {st.session_state.get('in_คลังรับผลิตภัณฑ์-ชื่อ', '')}")
    c.drawString(1.5*cm, h-5.8*cm, f"ที่อยู่ : {st.session_state.get('in_คลังรับผลิตภัณฑ์-ที่อยู่', '')}")

    c.drawString(11*cm, h-4.8*cm, "1.2 ผู้รับผลิตภัณฑ์")
    c.drawString(11*cm, h-5.3*cm, f"ชื่อ : {st.session_state.get('in_ผู้รับผลิตภัณฑ์-ชื่อ', '')}")
    c.drawString(11*cm, h-5.8*cm, f"ที่อยู่ : {st.session_state.get('in_ผู้รับผลิตภัณฑ์-ที่อยู่', '')}")

    c.drawString(1.5*cm, h-7.3*cm, "1.3 ผู้รับสินค้า (ปลายทาง)")
    c.drawString(1.5*cm, h-7.8*cm, f"ชื่อ : {st.session_state.get('in_ผู้รับสินค้า-ชื่อ', '')}")
    c.drawString(1.5*cm, h-8.3*cm, f"ที่อยู่ : {st.session_state.get('in_ผู้รับสินค้า-ที่อยู่', '')}")
    c.drawString(11*cm, h-8.3*cm, f"ตั๋วขนย้ายเลขที่ : {st.session_state.get('in_ผู้รับผลิตภัณฑ์-หมายเลขตั๋ว', '')}")

    c.line(1*cm, h-9.5*cm, 20*cm, h-9.5*cm)

    # --- 2. ข้อมูลการขนส่ง ---
    c.drawString(1.2*cm, h-10.1*cm, "2. ข้อมูลการขนส่ง")
    c.drawString(1.5*cm, h-10.7*cm, "2.1 ผู้ดำเนินการขนส่ง")
    c.drawString(1.5*cm, h-11.2*cm, f"ชื่อ : {st.session_state.get('in_ผู้ดำเนินการขนส่ง-ชื่อ', '')}")
    c.drawString(1.5*cm, h-11.7*cm, f"ใบอนุญาตเลขที่ : {st.session_state.get('in_ผู้ดำเนินการขนส่ง-ใบอนุญาต', '')}")

    c.drawString(11*cm, h-10.7*cm, "2.2 ข้อมูลพนักงานขับรถ")
    c.drawString(11*cm, h-11.2*cm, f"ชื่อ : {st.session_state.get('in_ข้อมูลพนักงานขับรถ-ชื่อ', '')}")
    c.drawString(11*cm, h-11.7*cm, f"ทะเบียนรถ : {st.session_state.get('in_ข้อมูลพนักงานขับรถ-ทะเบียนรถ', '')}")
    c.drawString(11*cm, h-12.2*cm, f"วัน/เวลาออก : {st.session_state.get('in_ข้อมูลพนักงานขับรถ-วันออกเดินทาง', '')} {st.session_state.get('in_ข้อมูลพนักงานขับรถ-เวลาออกเดินทาง', '')}")

    c.line(1*cm, h-15.0*cm, 20*cm, h-15.0*cm)

    # --- 3. ตารางรายการสินค้า ---
    c.drawString(1.2*cm, h-15.6*cm, "3. รายละเอียดน้ำมันเชื้อเพลิง")
    header = [["ลำดับ", "ช่องถัง", "หมายเลขซีล", "รายการน้ำมัน", "หน่วย", "จำนวน"]]
    data_rows = []
    total_q = 0
    for i, it in enumerate(items):
        data_rows.append([i+1, it['price'], it['amount'], it['product'], it['unit'], it['qty']])
        try: total_q += float(it['qty'])
        except: pass
    
    while len(data_rows) < 4: data_rows.append(["", "", "", "", "", ""])
    data_rows.append(["", "ยอดรวม >>", "", "", "", f"{total_q:,.2f}"])
    
    t = Table(header + data_rows, colWidths=[1.2*cm, 2.5*cm, 3.5*cm, 6.8*cm, 2*cm, 3*cm])
    t.setStyle(TableStyle([('FONT', (0,0), (-1,-1), FONT_NAME, 10), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('SPAN', (1, -1), (4, -1))]))
    t.wrapOn(c, 1*cm, h-19.5*cm)
    t.drawOn(c, 1*cm, h-19.5*cm)

    # --- 4. การยืนยัน ---
    sig_y = h-23.5*cm
    c.drawString(1.2*cm, h-20.5*cm, "4. การยืนยันและรับสินค้า")
    c.setFont(FONT_NAME, 10)
    c.drawString(1.5*cm, h-21*cm, "ข้าพเจ้าได้รับสินค้าตามรายการข้างต้นในสภาพเรียบร้อย ถูกต้องตามจำนวนและหมายเลขซีลที่ระบุไว้")

    c.drawCentredString(4.5*cm, sig_y, "..................................")
    c.drawCentredString(4.5*cm, sig_y-0.5*cm, f"( {st.session_state.get('in_การยืนยันและรับสินค้า-ผู้ออกเอกสาร', '')} )")
    c.drawCentredString(4.5*cm, sig_y-1*cm, "ผู้ออกเอกสาร")

    c.drawCentredString(10.5*cm, sig_y, "..................................")
    c.drawCentredString(10.5*cm, sig_y-0.5*cm, f"( {st.session_state.get('in_การยืนยันและรับสินค้า-พนักงานขับรถ', '')} )")
    c.drawCentredString(10.5*cm, sig_y-1*cm, "พนักงานขับรถ")

    c.drawCentredString(16.5*cm, sig_y, "..................................")
    c.drawCentredString(16.5*cm, sig_y-0.5*cm, f"( {st.session_state.get('in_การยืนยันและรับสินค้า-ผู้รับสินค้า', '')} )")
    c.drawCentredString(16.5*cm, sig_y-1*cm, "ผู้รับสินค้า")

    # ตีกรอบนอก
    c.rect(1*cm, 1*cm, 19*cm, h-2*cm)

    c.showPage(); c.save(); buf.seek(0)
    return buf

# ================= 4. MAIN UI =================
st.title("🚚 M POWER OIL - LOGISTICS SYSTEM")

with st.expander("🔍 ค้นหา/แก้ไข/สร้างซ้ำ จากบิลเก่า"):
    if not inv_df.empty:
        options = [f"{r[INV_KEY]} | {r.get('ผู้รับสินค้า-ชื่อ', '')}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกรายการบิล", [""] + options[::-1])
        if selected:
            sel_no = selected.split(" | ")[0]
            c1, c2 = st.columns(2)
            if c1.button("📝 โหลดมาแก้ไข"):
                old = inv_df[inv_df[INV_KEY] == sel_no].iloc[0].to_dict()
                st.session_state.editing_no = sel_no
                for f in transport_fields: st.session_state[f"in_{f}"] = str(old.get(f, ""))
                it_rows = item_df[item_df[ITEM_KEY if ITEM_KEY in item_df.columns else "invoice_no"] == sel_no].to_dict('records')
                st.session_state.invoice_items = [{"product": i.get('รายการ', i.get('product','')), "unit": i.get('หน่วย', i.get('unit','')), "qty": i.get('จำนวน', i.get('qty','')), "price": str(i.get('หมายเลขช่องถัง', i.get('price',''))), "amount": str(i.get('หมายเลขซีล', i.get('amount','')))} for i in it_rows]
                st.rerun()
            if c2.button("🔄 โหลดมาสร้างซ้ำ (บิลใหม่)"):
                old = inv_df[inv_df[INV_KEY] == sel_no].iloc[0].to_dict()
                st.session_state.editing_no = None
                for f in transport_fields: st.session_state[f"in_{f}"] = str(old.get(f, ""))
                it_rows = item_df[item_df[ITEM_KEY if ITEM_KEY in item_df.columns else "invoice_no"] == sel_no].to_dict('records')
                st.session_state.invoice_items = [{"product": i.get('รายการ', i.get('product','')), "unit": i.get('หน่วย', i.get('unit','')), "qty": i.get('จำนวน', i.get('qty','')), "price": str(i.get('หมายเลขช่องถัง', i.get('price',''))), "amount": str(i.get('หมายเลขซีล', i.get('amount','')))} for i in it_rows]
                st.rerun()

tabs = st.tabs(["📦 ข้อมูลคู่ค้า/ปลายทาง", "🚛 ข้อมูลการขนส่ง", "⛽ รายการน้ำมัน", "🏢 ข้อมูลบริษัท/ผู้จัดทำ"])
with tabs[0]:
    c1, c2 = st.columns(2)
    with c1: 
        for f in transport_fields[0:4]: st.text_input(f, key=f"in_{f}")
    with c2: 
        for f in transport_fields[4:11]: st.text_input(f, key=f"in_{f}")
with tabs[1]:
    c1, c2 = st.columns(2)
    with c1: 
        for f in transport_fields[11:17]: st.text_input(f, key=f"in_{f}")
    with c2: 
        for f in transport_fields[17:26]: st.text_input(f, key=f"in_{f}")
with tabs[2]:
    ca, cb, cc, cd, ce = st.columns([3,1,1,2,2])
    p_n = ca.text_input("รายการน้ำมัน", key="tmp_n")
    p_u = cb.text_input("หน่วย", value="ลิตร", key="tmp_u")
    p_q = cc.text_input("จำนวน", key="tmp_q")
    p_t = cd.text_input("หมายเลขช่องถัง", key="tmp_t")
    p_s = ce.text_input("หมายเลขซีล", key="tmp_s")
    if st.button("➕ เพิ่มรายการสินค้า"):
        if p_n and p_q:
            st.session_state.invoice_items.append({"product": p_n, "unit": p_u, "qty": p_q, "price": p_t, "amount": p_s})
            st.rerun()
    st.divider()
    for idx, it in enumerate(st.session_state.invoice_items):
        cx = st.columns([5, 1])
        cx[0].write(f"{idx+1}. {it['product']} | {it['qty']} {it['unit']} [ถัง:{it['price']} ซีล:{it['amount']}]")
        if cx[1].button("🗑️", key=f"del_{idx}"):
            st.session_state.invoice_items.pop(idx); st.rerun()
with tabs[3]:
    c1, c2 = st.columns(2)
    with c1: 
        for f in transport_fields[26:29]: st.text_input(f, key=f"in_{f}")
    with c2: 
        st.session_state.form_date = st.text_input("วันที่ (36)", value=st.session_state.form_date)
        for f in transport_fields[29:]: st.text_input(f, key=f"in_{f}")

st.divider()
if st.session_state.pdf_buffer:
    st.success("✅ บันทึกสำเร็จ! กดดาวน์โหลด PDF ด้านล่าง")
    st.download_button("📥 ดาวน์โหลดใบกำกับขนส่ง PDF", data=st.session_state.pdf_buffer, file_name="Invoice.pdf", mime="application/pdf", use_container_width=True)

if st.button("💾 บันทึกข้อมูลและออก PDF", type="primary", use_container_width=True):
    if not st.session_state.invoice_items:
        st.error("กรุณาเพิ่มรายการสินค้าก่อนบันทึก")
    else:
        def next_no():
            prefix = f"INV-{datetime.now().year}-{datetime.now().month:02d}"
            if inv_df.empty: return f"{prefix}-0001"
            curr = inv_df[inv_df[INV_KEY].astype(str).str.startswith(prefix)]
            if curr.empty: return f"{prefix}-0001"
            return f"{prefix}-{int(str(curr[INV_KEY].iloc[-1]).split('-')[-1]) + 1:04d}"
        
        final_no = st.session_state.editing_no if st.session_state.editing_no else next_no()
        if st.session_state.editing_no: # กรณีแก้ไข ให้ลบข้อมูลเก่าก่อนเขียนทับ
            try:
                for ws in [ws_inv, ws_item]:
                    found = ws.findall(final_no)
                    for f in reversed(found): ws.delete_rows(f.row)
            except: pass
        
        # บันทึกลง Google Sheet
        ws_inv.append_row([final_no, st.session_state.form_date] + [st.session_state[f"in_{f}"] for f in transport_fields])
        for it in st.session_state.invoice_items:
            ws_item.append_row([final_no, it['product'], it['unit'], it['qty'], it['price'], it['amount']])
        
        # สร้าง PDF เก็บใน Session
        st.session_state.pdf_buffer = generate_pdf_file(final_no, st.session_state.invoice_items)
        st.session_state.editing_no = None
        st.cache_data.clear()
        st.rerun()

if st.button("🆕 ล้างฟอร์ม / เริ่มบิลใหม่"):
    reset_form_action()
    st.rerun()
