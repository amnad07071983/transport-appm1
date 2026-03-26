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

# กำหนดชื่อคอลัมน์หลักให้ตรงกับใน Google Sheets (สำคัญมาก)
INV_KEY = "invoice_no"      # ชื่อหัวคอลัมน์แรกในแผ่นงาน Invoices
ITEM_KEY = "เลขที่บิล"       # ชื่อหัวคอลัมน์แรกในแผ่นงาน InvoiceItems (แก้ตามจริงใน Sheet)

@st.cache_resource
def init_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    return gspread.authorize(creds).open_by_key(SHEET_ID)

@st.cache_data(ttl=10) # ลด Cache ลงเพื่อให้เห็นข้อมูลล่าสุดเร็วขึ้น
def get_data_cached():
    client = init_sheet()
    try:
        inv = client.worksheet(INV_SHEET).get_all_records()
        items = client.worksheet(ITEM_SHEET).get_all_records()
        return pd.DataFrame(inv), pd.DataFrame(items)
    except Exception as e:
        st.error(f"Error loading sheets: {e}")
        return pd.DataFrame(), pd.DataFrame()

client = init_sheet()
inv_df, item_df = get_data_cached()
ws_inv = client.worksheet(INV_SHEET)
ws_item = client.worksheet(ITEM_SHEET)

# รายการฟิลด์ภาษาไทยตามโครงสร้างใหม่
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
def reset_form():
    st.session_state.invoice_items = []
    st.session_state.editing_no = None  
    st.session_state.form_date = datetime.now().strftime("%d/%m/%Y")
    for f in transport_fields:
        st.session_state[f"form_{f}"] = ""

if "invoice_items" not in st.session_state:
    reset_form()

# ================= 3. PDF & GENERATOR =================
def create_pdf(inv_info, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    if os.path.exists("p1.png"):
        c.saveState()
        c.drawImage("p1.png", 0, 0, width=w, height=h, mask='auto')
        c.restoreState()

    c.setFont("ThaiFontBold", 14)
    c.drawString(2*cm, h-3.2*cm, f"ลูกค้า: {inv_info.get('ผู้รับสินค้า-ชื่อ','')}")
    c.drawRightString(19*cm, h-3.2*cm, f"เลขที่: {inv_info.get('invoice_no','')}")
    
    header = [["ลำดับ", "รายการ", "หน่วย", "จำนวน", "ช่องถัง", "ซีล"]]
    rows = [[i+1, it['product'], it['unit'], it['qty'], it['price'], it['amount']] for i, it in enumerate(items)]
    t = Table(header + rows, colWidths=[1.2*cm, 7.8*cm, 1.5*cm, 1.5*cm, 3*cm, 3*cm])
    t.setStyle(TableStyle([('FONT', (0,0), (-1,-1), 'ThaiFontBold', 12), ('LINEBELOW', (0,0), (-1,0), 1, colors.black)]))
    t.wrapOn(c, 2*cm, h-11*cm)
    t.drawOn(c, 2*cm, h-11*cm)
    c.showPage(); c.save(); buf.seek(0)
    return buf

def next_inv_no():
    prefix = f"INV-{datetime.now().year}-{datetime.now().month:02d}"
    if inv_df.empty or INV_KEY not in inv_df.columns: return f"{prefix}-0001"
    curr = inv_df[inv_df[INV_KEY].astype(str).str.startswith(prefix)]
    if curr.empty: return f"{prefix}-0001"
    last_seq = int(str(curr[INV_KEY].iloc[-1]).split('-')[-1])
    return f"{prefix}-{last_seq + 1:04d}"

# ================= 4. MAIN UI =================
st.markdown("## 🚚 ระบบออกใบกำกับขนส่ง (M POWER OIL)")

with st.expander("🔍 ค้นหาประวัติและจัดการบิล"):
    if not inv_df.empty and INV_KEY in inv_df.columns:
        # ปรับการแสดงผล list รายชื่อใน dropdown
        options = [f"{r[INV_KEY]} | {r.get('ผู้รับสินค้า-ชื่อ', '')}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกรายการ", [""] + options[::-1])
        
        if selected:
            sel_no = selected.split(" | ")[0]
            # แก้ไขจุดที่เกิด Error: ใช้ INV_KEY และ ITEM_KEY ให้ตรงกับ Sheet
            old_inv = inv_df[inv_df[INV_KEY] == sel_no].iloc[0].to_dict()
            
            # ป้องกัน Error หากคอลัมน์ใน InvoiceItems เปลี่ยนเป็นภาษาไทยแล้ว
            current_item_key = ITEM_KEY if ITEM_KEY in item_df.columns else "invoice_no"
            it_rows = item_df[item_df[current_item_key] == sel_no].to_dict('records')
            
            # Mapping ข้อมูลสินค้า (ปรับ Key ตามชื่อภาษาไทยใน Sheet InvoiceItems)
            f_items = []
            for i in it_rows:
                f_items.append({
                    "product": i.get('รายการ', i.get('product', '')),
                    "unit": i.get('หน่วย', i.get('unit', '')),
                    "qty": i.get('จำนวน', i.get('qty', '')),
                    "price": str(i.get('หมายเลขช่องถัง', i.get('price', ''))),
                    "amount": str(i.get('หมายเลขซีล', i.get('amount', '')))
                })
            
            c1, c2 = st.columns(2)
            if c1.button("🔄 สร้างรายการซ้ำ (Copy)"):
                reset_form()
                st.session_state.invoice_items = f_items
                st.session_state.form_date = old_inv.get('date', datetime.now().strftime("%d/%m/%Y"))
                for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                st.rerun()
            if c2.button("📝 แก้ไขบิลนี้"):
                st.session_state.editing_no = sel_no
                st.session_state.invoice_items = f_items
                st.session_state.form_date = old_inv.get('date', datetime.now().strftime("%d/%m/%Y"))
                for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                st.rerun()
    else:
        st.info("ยังไม่มีข้อมูลประวัติในระบบ")

st.divider()

# --- INPUT TABS ---
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
    with col1: st.subheader("2.1 ผู้ดำเนินการขนส่ง")
    for f in transport_fields[11:17]: st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")
    with col2: st.subheader("2.2 ข้อมูลพนักงานขับรถ")
    for f in transport_fields[17:26]: st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")

with t3:
    st.subheader("3 รายการสินค้า")
    ca, cb, cc, cd, ce = st.columns([3,1,1,2,2])
    p_name = ca.text_input("รายการ", key="p_n")
    p_unit = cb.text_input("หน่วย", value="ลิตร", key="p_u")
    p_qty = cc.text_input("จำนวน", key="p_q")
    p_tank = cd.text_input("ช่องถัง", key="p_t")
    p_seal = ce.text_input("ซีล", key="p_s")
    if st.button("➕ เพิ่มรายการ"):
        if p_name and p_qty:
            st.session_state.invoice_items.append({"product": p_name, "unit": p_unit, "qty": p_qty, "price": p_tank, "amount": p_seal})
            st.rerun()
    for idx, item in enumerate(st.session_state.invoice_items):
        cols = st.columns([4, 2, 4, 1])
        cols[0].write(f"{item['product']} ({item['qty']} {item['unit']})")
        if cols[3].button("🗑️", key=f"del_{idx}"):
            st.session_state.invoice_items.pop(idx)
            st.rerun()

with t4:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("4 การยืนยันและรับสินค้า")
        for f in transport_fields[26:29]: st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")
    with col2:
        st.subheader("หัวกระดาษ - ผู้จัดจำหน่าย")
        st.session_state.form_date = st.text_input("วันที่", value=st.session_state.form_date)
        for f in transport_fields[29:]: st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")

# --- บันทึกข้อมูล ---
if st.button("💾 บันทึกและสร้าง PDF", type="primary", use_container_width=True):
    if not st.session_state.invoice_items:
        st.error("⚠️ กรุณาเพิ่มสินค้า")
    else:
        new_no = st.session_state.editing_no if st.session_state.editing_no else next_inv_no()
        # ลบข้อมูลเก่ากรณีแก้ไข
        if st.session_state.editing_no:
            try:
                for ws in [ws_inv, ws_item]:
                    found = ws.find(new_no)
                    while found: ws.delete_rows(found.row); found = ws.find(new_no)
            except: pass
        
        # บันทึก Invoice
        inv_row = [new_no, st.session_state.form_date]
        for f in transport_fields: inv_row.append(st.session_state.get(f"in_{f}", ""))
        ws_inv.append_row(inv_row)
        
        # บันทึก Items
        for it in st.session_state.invoice_items:
            ws_item.append_row([new_no, it['product'], it['unit'], it['qty'], it['price'], it['amount']])
            
        st.success(f"บันทึก {new_no} สำเร็จ!")
        st.download_button("📥 ดาวน์โหลด PDF", create_pdf({"invoice_no": new_no, "ผู้รับสินค้า-ชื่อ": st.session_state.get("in_ผู้รับสินค้า-ชื่อ", "")}, st.session_state.invoice_items), f"{new_no}.pdf")
        st.cache_data.clear()
