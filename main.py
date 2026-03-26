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
    st.error("⚠️ ไม่พบไฟล์ฟอนต์ 'THSARABUN BOLD.ttf' กรุณาตรวจสอบไฟล์")

SHEET_ID = "1fl86CxqgxlXAYU63GQOdCrL2jbPvSUdoXd1ndQvjnBM"
INV_SHEET = "Invoices"
ITEM_SHEET = "InvoiceItems"

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

client = init_sheet()
inv_df, item_df = get_data_cached()
ws_inv = client.worksheet(INV_SHEET)
ws_item = client.worksheet(ITEM_SHEET)

# รายการฟิลด์ภาษาไทยตามตารางที่คุณส่งมา (รวม 35 ฟิลด์หลัก)
transport_fields = [
    "ผู้รับสินค้า-ชื่อ", "ผู้รับสินค้า-ที่อยู่", "ผู้รับสินค้า-เลขผู้เสียภาษี", "ผู้รับสินค้า-เบอร์โทร",
    "คลังรับผลิตภัณฑ์-ชื่อ", "คลังรับผลิตภัณฑ์-เลขผู้เสียภาษี", "คลังรับผลิตภัณฑ์-ที่อยู่",
    "ผู้รับผลิตภัณฑ์-ชื่อ", "ผู้รับผลิตภัณฑ์-เลขผู้เสียภาษี", "ผู้รับผลิตภัณฑ์-ที่อยู่", "ผู้รับผลิตภัณฑ์-หมายเลขตั๋ว",
    "ผู้ดำเนินการขนส่ง-ชื่อ", "ผู้ดำเนินการขนส่ง-เลขผู้เสียภาษี", "ผู้ดำเนินการขนส่ง-ที่อยู่", "ผู้ดำเนินการขนส่ง-เบอร์โทร",
    "ผู้ดำเนินการขนส่ง-ประเภทผู้รับจ้าง", "ผู้ดำเนินการขนส่ง-ใบอนุญาต",
    "ข้อมูลพนักงานขับรถ-ชื่อ", "ข้อมูลพนักงานขับรถ-เลขใบขับขี่", "ข้อมูลพนักงานขับรถ-เบอร์โทร", "ข้อมูลพนักงานขับรถ-ทะเบียนรถ",
    "ข้อมูลพนักงานขับรถ-วิธีขนส่ง", "ข้อมูลพนักงานขับรถ-วันออกเดินทาง", "ข้อมูลพนักงานขับรถ-เวลาออกเดินทาง",
    "ข้อมูลพนักงานขับรถ-วันที่ถึงปลายทาง", "ข้อมูลพนักงานขับรถ-เวลาที่ถึงปลายทาง",
    "การยืนยันและรับสินค้า-ผู้ออกเอกสาร", "การยืนยันและรับสินค้า-พนักงานขับรถ", "การืนยันและรับสินค้า-ผู้รับสินค้า",
    "ผู้จำหน่าย-ชื่อ", "ผู้จำหน่าย-ที่อยู่", "ผู้จำหน่าย-เลขผู้เสียภาษี", "ผู้จำหน่าย-เบอร์โทร",
    "ผู้จำหน่าย-ชื่อเอกสาร", "ผู้จำหน่าย-อธิบายเพิ่ม"
]

# ================= 2. SESSION STATE & RESET =================
def reset_form():
    st.session_state.invoice_items = []
    st.session_state.editing_no = None  
    st.session_state.form_date = datetime.now().strftime("%d/%m/%Y")
    for f in transport_fields:
        st.session_state[f"form_{f}"] = ""

if "invoice_items" not in st.session_state:
    reset_form()

# ================= 3. PDF FUNCTIONS (With P1 Watermark) =================
def create_pdf(inv_info, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    
    # วาดลายน้ำ P1.png (ฉากหลัง)
    if os.path.exists("p1.png"):
        c.saveState()
        c.setFillAlpha(1.0) # ความคมชัด 100%
        c.drawImage("p1.png", 0, 0, width=w, height=h, mask='auto')
        c.restoreState()

    # ข้อมูลผู้จัดจำหน่าย (หัวกระดาษ)
    c.setFont("ThaiFontBold", 20)
    c.drawString(2*cm, h-2*cm, str(inv_info.get('ผู้จำหน่าย-ชื่อ', '')))
    c.setFont("ThaiFontBold", 26)
    c.drawRightString(19*cm, h-2.2*cm, str(inv_info.get('ผู้จำหน่าย-ชื่อเอกสาร', 'ใบกำกับขนส่ง')))
    
    # ข้อมูลเลขที่และลูกค้า
    c.setFont("ThaiFontBold", 14)
    c.drawString(2*cm, h-3.2*cm, f"ลูกค้า: {inv_info.get('ผู้รับสินค้า-ชื่อ','')}")
    c.drawRightString(19*cm, h-3.2*cm, f"เลขที่: {inv_info.get('invoice_no','')}")
    
    # ตารางสินค้า
    header = [["ลำดับ", "รายการ", "หน่วย", "จำนวน", "หมายเลขช่องถัง", "หมายเลขซีล"]]
    rows = [[i+1, it['product'], it['unit'], it['qty'], it['price'], it['amount']] for i, it in enumerate(items)]
    
    t = Table(header + rows, colWidths=[1.2*cm, 7.8*cm, 1.5*cm, 1.5*cm, 3*cm, 3*cm])
    t.setStyle(TableStyle([
        ('FONT', (0,0), (-1,-1), 'ThaiFontBold', 12),
        ('LINEBELOW', (0,0), (-1,0), 1, colors.black),
        ('ALIGN', (0,0), (0,-1), 'CENTER'),
        ('ALIGN', (2,0), (-1,-1), 'CENTER'),
    ]))
    
    # วาดตารางในตำแหน่งที่เหมาะสม (กึ่งกลางหน้ากระดาษ)
    t.wrapOn(c, 2*cm, h-11*cm)
    t.drawOn(c, 2*cm, h-11*cm)
    
    c.showPage()
    c.save()
    buf.seek(0)
    return buf

def next_inv_no():
    data = client.worksheet(INV_SHEET).get_all_records()
    df = pd.DataFrame(data)
    prefix = f"INV-{datetime.now().year}-{datetime.now().month:02d}"
    if df.empty: return f"{prefix}-0001"
    curr = df[df["invoice_no"].astype(str).str.startswith(prefix)]
    if curr.empty: return f"{prefix}-0001"
    last_seq = int(str(curr["invoice_no"].iloc[-1]).split('-')[-1])
    return f"{prefix}-{last_seq + 1:04d}"

# ================= 4. MAIN UI =================
st.markdown("## 🚚 ระบบออกใบกำกับขนส่ง (M POWER OIL)")

with st.expander("🔍 ค้นหาประวัติและจัดการบิล"):
    if not inv_df.empty:
        options = [f"{r['invoice_no']} | {r['ผู้รับสินค้า-ชื่อ']}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกรายการ", [""] + options[::-1])
        if selected:
            sel_no = selected.split(" | ")[0]
            old_inv = inv_df[inv_df["invoice_no"] == sel_no].iloc[0].to_dict()
            it_rows = item_df[item_df["invoice_no"] == sel_no].to_dict('records')
            
            # ดึงข้อมูลสินค้ากลับมา (Mapping ตามหัวตารางไทย)
            f_items = [{"product": i['รายการ'], "unit": i['หน่วย'], "qty": i['จำนวน'], "price": str(i['หมายเลขช่องถัง']), "amount": str(i['หมายเลขซีล'])} for i in it_rows]
            
            c1, c2 = st.columns(2)
            if c1.button("🔄 สร้างรายการซ้ำ (Copy)"):
                reset_form()
                st.session_state.invoice_items = f_items
                for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                st.rerun()
            if c2.button("📝 แก้ไขบิลนี้"):
                st.session_state.editing_no = sel_no
                st.session_state.invoice_items = f_items
                for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                st.rerun()

st.divider()

# --- จัดกลุ่ม Tabs ตามรูปภาพ image_24449c.png ---
t1, t2, t3, t4 = st.tabs(["📦 ผลิตภัณฑ์ & ปลายทาง", "🚛 ขนส่ง & พนักงานขับ", "⛽ รายการสินค้า", "🏢 ยืนยัน & ผู้จัดจำหน่าย"])

with t1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1.3 ผู้รับสินค้า (ปลายทาง)")
        for f in transport_fields[0:4]:
            st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")
    with col2:
        st.subheader("1.1 คลังรับผลิตภัณฑ์ (ต้นทาง)")
        for f in transport_fields[4:7]:
            st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")
        st.subheader("1.2 ผู้รับผลิตภัณฑ์")
        for f in transport_fields[7:11]:
            st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")

with t2:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("2.1 ผู้ดำเนินการขนส่ง")
        for f in transport_fields[11:17]:
            st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")
    with col2:
        st.subheader("2.2 ข้อมูลพนักงานขับรถ")
        for f in transport_fields[17:26]:
            st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")

with t3:
    st.subheader("3 รายการสินค้า")
    ca, cb, cc, cd, ce = st.columns([3,1,1,2,2])
    p_name = ca.text_input("ชื่อสินค้า/รายการ", key="p_n")
    p_unit = cb.text_input("หน่วย", value="ลิตร", key="p_u")
    p_qty = cc.text_input("จำนวน", key="p_q")
    p_tank = cd.text_input("หมายเลขช่องถัง", key="p_t")
    p_seal = ce.text_input("หมายเลขซีล", key="p_s")
    
    if st.button("➕ เพิ่มรายการสินค้า"):
        if p_name and p_qty:
            st.session_state.invoice_items.append({
                "product": p_name, "unit": p_unit, "qty": p_qty, 
                "price": str(p_tank), "amount": str(p_seal)
            })
            st.rerun()
            
    if st.session_state.invoice_items:
        st.write("---")
        for idx, item in enumerate(st.session_state.invoice_items):
            c1, c2, c3, c4 = st.columns([4, 2, 4, 1])
            c1.write(f"**{idx+1}. {item['product']}** ({item['qty']} {item['unit']})")
            c2.write(f"ถัง: {item['price']}")
            c3.write(f"ซีล: {item['amount']}")
            if c4.button("🗑️", key=f"del_{idx}"):
                st.session_state.invoice_items.pop(idx)
                st.rerun()

with t4:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("4 การยืนยันและรับสินค้า")
        for f in transport_fields[26:29]:
            st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")
    with col2:
        st.subheader("หัวกระดาษ - ผู้จัดจำหน่าย")
        st.session_state.form_date = st.text_input("วันที่ (date)", value=st.session_state.form_date)
        for f in transport_fields[29:]:
            st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")

# --- บันทึกและดาวน์โหลด ---
if st.button("💾 บันทึกข้อมูลและสร้าง PDF", type="primary", use_container_width=True):
    if not st.session_state.invoice_items:
        st.error("⚠️ ไม่สามารถบันทึกได้เนื่องจากไม่มีรายการสินค้า")
    else:
        with st.spinner("กำลังบันทึกข้อมูลลงระบบ..."):
            new_no = st.session_state.editing_no if st.session_state.editing_no else next_inv_no()
            
            # ลบข้อมูลเก่ากรณีแก้ไข
            if st.session_state.editing_no:
                try:
                    for ws in [ws_inv, ws_item]:
                        found = ws.find(new_no)
                        while found: ws.delete_rows(found.row); found = ws.find(new_no)
                except: pass

            # 1. บันทึก Invoices (หัวบิล + ขนส่ง)
            inv_row = [new_no, st.session_state.form_date]
            for f in transport_fields:
                inv_row.append(st.session_state.get(f"in_{f}", ""))
            ws_inv.append_row(inv_row)

            # 2. บันทึก InvoiceItems (รายการสินค้า)
            for it in st.session_state.invoice_items:
                ws_item.append_row([new_no, it['product'], it['unit'], it['qty'], it['price'], it['amount']])
            
            st.success(f"✅ บันทึกเลขที่ {new_no} เรียบร้อย!")
            
            # 3. สร้าง PDF สำหรับดาวน์โหลด
            pdf_info = {
                "invoice_no": new_no, 
                "ผู้รับสินค้า-ชื่อ": st.session_state.get("in_ผู้รับสินค้า-ชื่อ", ""), 
                "ผู้จำหน่าย-ชื่อ": st.session_state.get("in_ผู้จำหน่าย-ชื่อ", ""),
                "ผู้จำหน่าย-ชื่อเอกสาร": st.session_state.get("in_ผู้จำหน่าย-ชื่อเอกสาร", "")
            }
            st.download_button(
                "📥 ดาวน์โหลด PDF (Watermark P1)", 
                create_pdf(pdf_info, st.session_state.invoice_items), 
                f"{new_no}.pdf",
                mime="application/pdf"
            )
            st.cache_data.clear()
