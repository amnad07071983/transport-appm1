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

# โหลดข้อมูลเริ่มต้น
client = init_sheet()
inv_df, item_df = get_data_cached()
ws_inv = client.worksheet(INV_SHEET)
ws_item = client.worksheet(ITEM_SHEET)

# รายการฟิลด์ทั้งหมดในหน้า Invoices (ลำดับต้องตรงกับ Sheet ของคุณ)
transport_fields = [
    "doc_status", "car_id", "driver_name", "payment_status", "date_out", "time_out",
    "date_in", "time_in", "ref_tax_id", "ref_receipt_id", "seal_no",
    "pay_term", "ship_method", "driver_license", "receiver_name",
    "issuer_name", "sender_name", "checker_name", "remark",
    "comp_name", "comp_address", "comp_tax_id", "comp_phone", "comp_doc_title"
]

# ================= 2. SESSION STATE & RESET =================
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

# ================= 3. PDF FUNCTIONS =================
def create_pdf(inv_info, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    
    # วาดพื้นหลัง p2.png (ชัด 100%)
    if os.path.exists("p2.png"):
        c.saveState()
        c.setFillAlpha(1.0)
        c.drawImage("p2.png", 0, 0, width=w, height=h, mask='auto')
        c.restoreState()

    # หัวเอกสาร
    c.setFont("ThaiFontBold", 22)
    c.drawString(2*cm, h-2*cm, str(inv_info.get('comp_name', '')))
    c.setFont("ThaiFontBold", 26)
    c.drawRightString(19*cm, h-2*cm, str(inv_info.get('comp_doc_title', 'ใบกำกับขนส่ง')))
    
    c.setFont("ThaiFontBold", 14)
    c.drawString(2*cm, h-3*cm, f"ลูกค้า: {inv_info.get('customer','')}")
    c.drawRightString(19*cm, h-3*cm, f"เลขที่: {inv_info.get('invoice_no','')}")
    
    # ตารางสินค้า
    header = [["ลำดับ", "รายการ", "หน่วย", "จำนวน", "หมายเลขช่องถัง", "หมายเลขซีล"]]
    rows = [[i+1, it['product'], it['unit'], it['qty'], it['price'], it['amount']] for i, it in enumerate(items)]
    
    t = Table(header + rows, colWidths=[1.2*cm, 7.8*cm, 1.5*cm, 1.5*cm, 3*cm, 3*cm])
    t.setStyle(TableStyle([('FONT', (0,0), (-1,-1), 'ThaiFontBold', 12), ('LINEBELOW', (0,0), (-1,0), 1, colors.black)]))
    t.wrapOn(c, 2*cm, h-10*cm)
    t.drawOn(c, 2*cm, h-10*cm)
    
    c.showPage(); c.save(); buf.seek(0)
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

# --- ค้นหา/แก้ไข/ทำซ้ำ ---
with st.expander("🔍 ค้นหาประวัติและจัดการบิล"):
    if not inv_df.empty:
        options = [f"{r['invoice_no']} | {r['customer']}" for _, r in inv_df.iterrows()]
        selected = st.selectbox("เลือกรายการ", [""] + options[::-1])
        if selected:
            sel_no = selected.split(" | ")[0]
            old_inv = inv_df[inv_df["invoice_no"] == sel_no].iloc[0].to_dict()
            it_rows = item_df[item_df["invoice_no"] == sel_no].to_dict('records')
            
            # แปลงรายการสินค้าให้เป็น Format ที่แก้ไขได้
            f_items = [{"product": i['product'], "unit": i['unit'], "qty": i['qty'], "price": str(i['price']), "amount": str(i['amount'])} for i in it_rows]
            
            c1, c2 = st.columns(2)
            if c1.button("🔄 สร้างรายการซ้ำ (Copy)"):
                reset_form()
                st.session_state.form_customer = old_inv.get("customer","")
                st.session_state.form_address = old_inv.get("address","")
                st.session_state.invoice_items = f_items
                for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                st.rerun()
            if c2.button("📝 แก้ไขบิลนี้"):
                st.session_state.editing_no = sel_no
                st.session_state.form_customer = old_inv.get("customer","")
                st.session_state.form_address = old_inv.get("address","")
                st.session_state.invoice_items = f_items
                for f in transport_fields: st.session_state[f"form_{f}"] = str(old_inv.get(f, ""))
                st.rerun()

st.divider()

if st.session_state.editing_no:
    st.warning(f"🚨 กำลังแก้ไขบิล: {st.session_state.editing_no}")
    if st.button("❌ ยกเลิก"): reset_form(); st.rerun()

# --- INPUT TABS ---
t1, t2, t3, t4 = st.tabs(["👤 ลูกค้า", "🚛 ขนส่ง", "📦 สินค้า", "🏢 บริษัท"])

with t1:
    customer = st.text_input("ชื่อลูกค้า", value=st.session_state.form_customer)
    address = st.text_area("ที่อยู่", value=st.session_state.form_address)
    invoice_date = st.text_input("วันที่", value=st.session_state.form_date)

with t2:
    col_l, col_r = st.columns(2)
    # แสดงฟิลด์ขนส่งสำคัญ (ตัวอย่าง 10 ฟิลด์แรก)
    for i, f in enumerate(transport_fields[:10]):
        val = st.session_state.get(f"form_{f}", "")
        if i % 2 == 0: col_l.text_input(f, value=val, key=f"in_{f}")
        else: col_r.text_input(f, value=val, key=f"in_{f}")

with t3:
    st.subheader("รายการสินค้า")
    # ส่วนเพิ่มสินค้า
    ca, cb, cc, cd, ce = st.columns([3,1,1,2,2])
    p_name = ca.text_input("สินค้า", key="p_n")
    p_unit = cb.text_input("หน่วย", value="ลิตร", key="p_u")
    p_qty = cc.text_input("จำนวน", key="p_q")
    p_tank = cd.text_input("หมายเลขช่องถัง", key="p_t")
    p_seal = ce.text_input("หมายเลขซีล", key="p_s")
    
    if st.button("➕ เพิ่มรายการ"):
        if p_name and p_qty:
            st.session_state.invoice_items.append({
                "product": p_name, "unit": p_unit, "qty": p_qty, 
                "price": str(p_tank), "amount": str(p_seal)
            })
            st.rerun()
    
    # ส่วนแสดงตารางพร้อมปุ่มลบ
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
    # ข้อมูลบริษัทและส่วนที่เหลือ
    for f in transport_fields[10:]:
        st.text_input(f, value=st.session_state.get(f"form_{f}", ""), key=f"in_{f}")

# --- บันทึกข้อมูล ---
if st.button("💾 บันทึกข้อมูล", type="primary", use_container_width=True):
    if not st.session_state.invoice_items:
        st.error("⚠️ กรุณาเพิ่มสินค้าอย่างน้อย 1 รายการ")
    else:
        with st.spinner("กำลังบันทึกข้อมูล..."):
            new_no = st.session_state.editing_no if st.session_state.editing_no else next_inv_no()
            
            # ลบแถวเก่าถ้าเป็นการแก้ไข
            if st.session_state.editing_no:
                try:
                    for ws in [ws_inv, ws_item]:
                        found = ws.find(new_no)
                        while found: ws.delete_rows(found.row); found = ws.find(new_no)
                except: pass

            # เตรียมข้อมูล Invoices (รวม 9 คอลัมน์แรก + ฟิลด์ที่เหลือ)
            inv_row = [new_no, invoice_date, customer, address, 0, 0, 0, 0, 0]
            for f in transport_fields:
                val = st.session_state.get(f"in_{f}", "")
                inv_row.append(val)
            ws_inv.append_row(inv_row)

            # บันทึก InvoiceItems
            for it in st.session_state.invoice_items:
                ws_item.append_row([new_no, it['product'], it['unit'], it['qty'], it['price'], it['amount']])
            
            st.success(f"บันทึกเลขที่ {new_no} สำเร็จ!")
            
            # สร้างข้อมูลสำหรับ PDF
            pdf_data = {
                "invoice_no": new_no, "customer": customer, "date": invoice_date,
                "comp_name": st.session_state.get("in_comp_name", ""),
                "comp_doc_title": st.session_state.get("in_comp_doc_title", "ใบกำกับขนส่ง")
            }
            st.download_button("📥 ดาวน์โหลด PDF", create_pdf(pdf_data, st.session_state.invoice_items), f"{new_no}.pdf")
            st.cache_data.clear()
