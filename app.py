import streamlit as st
import json
import base64
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io
import os
from contract_to_excel import generate_rows, write_excel, HEADERS, COL_WIDTHS

# ─── Page Config ───
st.set_page_config(
    page_title="MeowAI 🐾 Contract Extractor",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── Mascot Setup ───
try:
    from mascot_data import MASCOT_B64
except ImportError:
    MASCOT_B64 = ""

if MASCOT_B64:
    mascot_img_html = f'<div class="mascot-container"><img src="data:image/png;base64,{MASCOT_B64}" class="mascot-avatar" /></div>'
else:
    mascot_img_html = '<div class="mascot-container" style="display:flex; align-items:center; justify-content:center; font-size:80px;">🐱</div>'

# ─── Premium Grey Cat AI Theme CSS ───
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
    /* Safe font override that doesn't break Streamlit Material Icons */
    html, body, .stApp, p, label, h1, h2, h3, h4, h5, h6, input, button, textarea, .stMarkdown {
        font-family: 'Prompt', sans-serif;
    }
    
    /* Hide top bar and force sidebar to stay open */
    #MainMenu, footer {visibility: hidden !important;}
    header {background-color: transparent !important;}
    [data-testid="collapsedControl"] { display: none !important; }
    
    .block-container { padding: 2rem 3rem !important; max-width: 1200px; }

    /* Grey Tabby Cat Palette: Silvery Grey, Slate, Charcoal */
    .stApp {
        background-color: #F8FAFC !important; /* Silvery white/grey */
        background-image: radial-gradient(#E2E8F0 1px, transparent 1px);
        background-size: 20px 20px;
    }

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 2px solid #E2E8F0 !important;
    }
    
    .sidebar-profile {
        text-align: center;
        padding: 1.5rem 0;
        margin-bottom: 1rem;
    }
    
    .mascot-container {
        width: 180px;
        height: 180px;
        margin: 0 auto 15px auto;
        animation: float 5s ease-in-out infinite;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    
    .mascot-avatar {
        width: 180px;
        height: 180px;
        object-fit: contain;
        mix-blend-mode: multiply;
        display: block;
    }
    
    @keyframes float {
        0% { transform: translateY(0px) rotate(0deg); }
        50% { transform: translateY(-8px) rotate(1deg); }
        100% { transform: translateY(0px) rotate(0deg); }
    }

    .sidebar-title {
        color: #334155;
        font-weight: 800;
        font-size: 1.5rem;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .sidebar-subtitle {
        color: #64748B;
        font-size: 0.95rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1.5px;
    }

    /* Chat Bubble / Hero Section */
    .chat-bubble {
        background: #FFFFFF;
        border-radius: 24px;
        border-bottom-left-radius: 4px;
        padding: 2rem;
        box-shadow: 0 10px 30px rgba(100, 116, 139, 0.1);
        position: relative;
        margin-bottom: 2.5rem;
        border: 2px solid #E2E8F0;
        display: flex;
        align-items: center;
        gap: 2rem;
    }
    .chat-bubble::before {
        content: '';
        position: absolute;
        bottom: -2px;
        left: -20px;
        border-width: 20px 20px 0 0;
        border-style: solid;
        border-color: #E2E8F0 transparent transparent transparent;
    }
    .chat-bubble::after {
        content: '';
        position: absolute;
        bottom: 0px;
        left: -16px;
        border-width: 17px 17px 0 0;
        border-style: solid;
        border-color: #FFFFFF transparent transparent transparent;
    }
    .chat-bubble h3 {
        color: #1E293B !important;
        margin: 0 0 0.5rem 0 !important;
        font-size: 1.6rem !important;
        font-weight: 700 !important;
    }
    .chat-bubble p {
        color: #475569 !important;
        margin: 0 !important;
        font-size: 1.1rem !important;
        line-height: 1.6 !important;
    }
    
    /* Inputs */
    .stTextInput > div > div > input {
        border-radius: 12px !important;
        border: 2px solid #E2E8F0 !important;
        padding: 0.7rem 1.2rem !important;
        background-color: #F8FAFC !important;
        color: #334155 !important;
        font-weight: 500;
        transition: all 0.3s;
    }
    .stTextInput > div > div > input:focus {
        border-color: #94A3B8 !important;
        box-shadow: 0 0 0 4px rgba(148, 163, 184, 0.15) !important;
        background-color: #FFFFFF !important;
    }

    /* Primary Button */
    div[data-testid="stButton"] > button {
        background: linear-gradient(135deg, #64748B 0%, #475569 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 14px !important;
        padding: 0.8rem 2rem !important;
        font-weight: 700 !important;
        font-size: 1.1rem !important;
        width: 100%;
        box-shadow: 0 6px 20px rgba(71, 85, 105, 0.3) !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }
    div[data-testid="stButton"] > button:hover {
        transform: translateY(-3px) scale(1.02) !important;
        box-shadow: 0 10px 25px rgba(71, 85, 105, 0.45) !important;
    }
    
    /* Checkbox */
    [data-testid="stCheckbox"] label span {
        color: #475569 !important;
        font-weight: 500 !important;
    }

</style>
""", unsafe_allow_html=True)


# ─── Sidebar configuration ───
with st.sidebar:
    st.markdown(
f"""<div class="sidebar-profile">
{mascot_img_html}
<h2 class="sidebar-title">MeowAI 🐾</h2>
<p class="sidebar-subtitle">Smart Extractor</p>
</div>""", 
        unsafe_allow_html=True
    )
    
    st.markdown("### 🔑 API Access")
    api_key = st.text_input("Gemini API Key", type="password", placeholder="Paste your API key...")
    
    st.markdown("---")
    st.markdown("### 🏨 Property Details")
    prop_type = st.radio("Type", ["Hotel / Resort", "Cruise Ship"])
    hotel_id = st.text_input("Property ID", placeholder="e.g. 12711588")
    supplier = st.text_input("Supplier", placeholder="e.g. Agoda")


# ─── Main Content ───
st.markdown("""
<div class="chat-bubble">
    <div style="font-size: 3.5rem; animation: float 3s ease-in-out infinite;">😸</div>
    <div>
        <h3>สวัสดี DATA TEAM! 🐾</h3>
        <p>หนู MeowAI พร้อมทำงานแล้วค่ะ! แค่โยนไฟล์ PDF สัญญาโรงแรมมาให้หนูเดี๋ยวหนูจะใช้พลังเหมียวสกัดข้อมูล Rate, Date, Terms ทั้งหมดให้ออกมาเป็น Excel สวยๆ พร้อมนำไปใช้งานต่อได้เลย!</p>
    </div>
</div>
""", unsafe_allow_html=True)

col_left, col_right = st.columns([1, 1.2], gap="large")

with col_left:
    st.markdown("#### 📄 1. โยนไฟล์สัญญามาเลย! (PDF)")
    uploaded_file = st.file_uploader("Select Contract PDF", type=["pdf"])
    
    st.markdown("#### 🐟 2. ประเภทของสัญญา (Contract Types)")
    ct_main = st.checkbox("Main Contract (Base Rate)", value=True)
    
    col_eb1, col_eb2 = st.columns([1, 1])
    with col_eb1:
        ct_eb = st.checkbox("Early Bird", value=False)
    with col_eb2:
        eb_code = st.text_input("EB Prefix", value="E.B DAYS", disabled=not ct_eb, label_visibility="collapsed")
        
    col_pr1, col_pr2, col_pr3 = st.columns([1.2, 1, 1])
    with col_pr1:
        ct_promo = st.checkbox("Promotion", value=False)
    with col_pr2:
        promo_code = st.text_input("Code", placeholder="Code", disabled=not ct_promo, label_visibility="collapsed")
    with col_pr3:
        promo_till = st.text_input("Till", placeholder="YYYY-MM-DD", disabled=not ct_promo, label_visibility="collapsed")
        
    ct_por = st.checkbox("POR (Price on Request)", value=True)

with col_right:
    st.markdown("#### 🚪 3. ตั้งค่าห้องพัก (Rooms Config)")
    
    if 'rooms' not in st.session_state:
        st.session_state.rooms = [{"room_name": "", "room_id": ""}]

    def add_room():
        st.session_state.rooms.append({"room_name": "", "room_id": ""})

    def remove_room(i):
        if len(st.session_state.rooms) > 1:
            st.session_state.rooms.pop(i)

    for idx, room in enumerate(list(st.session_state.rooms)):
        rc1, rc2, rc3 = st.columns([5, 4, 1.5])
        with rc1:
            room["room_name"] = st.text_input(f"Name {idx}", value=room["room_name"], key=f"rn_{idx}", placeholder="Deluxe Room", label_visibility="collapsed")
        with rc2:
            room["room_id"] = st.text_input(f"ID {idx}", value=room["room_id"], key=f"ri_{idx}", placeholder="101", label_visibility="collapsed")
        with rc3:
            if st.button("🗑️", key=f"del_{idx}", help="ลบห้องนี้"):
                remove_room(idx)
                st.rerun()
                
    st.button("🐾 เพิ่มห้องพัก (Add Room)", on_click=add_room, use_container_width=True)


# ─── Data Extraction Logic ───
st.markdown("<br>", unsafe_allow_html=True)

valid_rooms = [r for r in st.session_state.rooms if r["room_name"].strip() and r["room_id"].strip()]

cts = []
if ct_main: cts.append({"type": "main", "label": "Main Contract"})
if ct_eb: cts.append({"type": "eb", "label": "Early Bird", "code": eb_code})
if ct_promo: cts.append({"type": "promo", "label": "Promotion", "promo_code": promo_code, "promo_till": promo_till})
if ct_por: cts.append({"type": "por", "label": "POR"})

if st.button("🚀 สั่งเหมียวทำงาน! (Process Contract)", use_container_width=True):
    if not api_key:
        st.error("⚠️ เมี้ยว! อย่าลืมใส่ Gemini API Key ที่แถบด้านซ้ายนะเหมียว")
    elif not uploaded_file:
        st.error("⚠️ เมี้ยว! ขาดไฟล์ PDF นะเหมียว อัพโหลดให้หน่อย")
    elif not hotel_id:
        st.error("⚠️ เมี้ยว! ใส่ Property ID ให้ครบด้วยนะ")
    elif not valid_rooms:
        st.error("⚠️ เมี้ยว! ต้องมีห้องพักอย่างน้อย 1 ห้องนะ")
    elif not cts:
        st.error("⚠️ เมี้ยว! เลือกประเภทสัญญาให้หน่อยสิ")
    else:
        with st.spinner("🐱 น้องแมวกำลังอ่าน contract อยู่นะ… รอสักครู่ (ไม่เกิน 60 วินาที)"):
            try:
                file_bytes = uploaded_file.read()
                b64_data = base64.b64encode(file_bytes).decode("utf-8")

                ct_desc_parts = []
                for c in cts:
                    if c["type"] == "main": ct_desc_parts.append("Main Contract: extract all standard net rates per period.")
                    elif c["type"] == "eb": ct_desc_parts.append(f'Early Bird: detect ALL tiers (days + discount%). promo_code format "{c["code"]} X DAYS". Respect blackout dates (eb_blackout=true).')
                    elif c["type"] == "promo": ct_desc_parts.append(f'Promotion: promo_code="{c["promo_code"]}", promo_book_till="{c["promo_till"]} 23:59:59".')
                    elif c["type"] == "por": ct_desc_parts.append('POR: net_price=0, promo_code="POR Rate".')
                ct_desc = " | ".join(ct_desc_parts)

                prompt_text = f"""You are an expert hotel contract data extractor. Extract ALL data from this PDF.

ROOMS (use exactly these IDs and names, do not invent): {json.dumps(valid_rooms)}
CONTRACT TYPES TO EXTRACT: {ct_desc}
PROPERTY TYPE: {prop_type}

CRITICAL RULES:
1. Extract ONE period per distinct date range. Split if surcharge/blackout/min-nights rules differ.
2. WEEKDAY/WEEKEND RATES:
   - Set has_weekday_weekend=true
   - rates dict = BASE (weekday) rate per room
   - has_surcharge=true, surcharge_rates = per-room weekend supplement dict e.g. {{"room_id_1": 1500, "room_id_2": 3000}}
   - If same surcharge for all rooms, use surcharge_amount (single value) instead
3. ALL HTML fields must use real HTML tags: <p>, <strong>, <span style="color:#ff0000;">
4. meals_and_info must be full HTML summary: contract name, validity, surcharges, meal rates, blackout dates, min nights, all key notes.
5. cancellation_policy and child_policy must be full HTML extracted from PDF.
6. promo_book_till format: "YYYY-MM-DD 23:59:59" (only fill if PDF explicitly states a booking deadline)
7. Early Bird: promo_book_till = null UNLESS PDF clearly states a "book by" deadline for EB.
8. net_price = integer only (no decimals, no commas).
9. rates dict key = room_id string exactly as given in ROOMS above.
10. For CRUISE property: detect night package from PDF (e.g. "1 Night", "2 Nights") and use as promo_code in promotions/tiers.
11. hotel_id = "{hotel_id}", hotel_supplier = "{supplier}"

Return ONLY valid JSON (no markdown, no explanation):
{{
  "hotel_id": "{hotel_id}",
  "hotel_supplier": "{supplier}",
  "abf": "Included",
  "cancellation_policy": "<p><strong>CANCELLATION:</strong></p><p>...</p>",
  "child_policy": "<p><span style=\"color:#008000;\"><strong>Maximum Occupancy: 2A+1C</strong></span></p><p>Child 0-6.99 years: FOC</p>",
  "meals_and_info": "<p><strong>CONTRACT NAME : VALIDITY DATES</strong></p><p>Notes...</p>",
  "child_share_bed_abf": null,
  "child_extra_bed_abf": null,
  "extra_bed_abf": null,
  "extra_bed_no_abf": null,
  "full_board": null,
  "half_board": null,
  "rooms": {json.dumps(valid_rooms)},
  "periods": [
    {{
      "start_date": "YYYY-MM-DD",
      "end_date": "YYYY-MM-DD",
      "season": "Season name",
      "is_peak": false,
      "eb_blackout": false,
      "has_surcharge": false,
      "surcharge_amount": 0,
      "surcharge_rates": {{}},
      "surcharge_currency": "THB",
      "min_nights_stay": null,
      "cutoff_date": null,
      "room_allotment": null,
      "has_weekday_weekend": false,
      "rates": {{"room_id_here": 0}},
      "early_bird_tiers": [
        {{"days": 60, "discount_pct": 15, "promo_code": "E.B 60 DAYS", "promo_book_till": null}}
      ],
      "promotions": [
        {{"discount_pct": 20, "promo_code": "EBO", "promo_book_till": "YYYY-MM-DD 23:59:59"}}
      ]
    }}
  ]
}}"""

                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
                payload = {
                    "contents": [{
                        "role": "user",
                        "parts": [
                            {"inlineData": {"mimeType": "application/pdf", "data": b64_data}},
                            {"text": prompt_text}
                        ]
                    }],
                    "generationConfig": {
                        "temperature": 0.1,
                        "responseMimeType": "application/json"
                    }
                }

                resp = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=120)
                resp_data = resp.json()

                if resp.status_code != 200:
                    raise Exception(resp_data.get("error", {}).get("message", "Gemini API Error"))

                text_output = resp_data["candidates"][0]["content"]["parts"][0]["text"]
                parsed_data = json.loads(text_output)

                # ─── Generate rows using contract_to_excel logic ───
                rows = generate_rows(parsed_data, cts)

                # ─── Write 43-column Excel (dashboard format) ───
                wb = Workbook()
                ws = wb.active
                ws.title = "Sheet1"

                header_fill = PatternFill("solid", start_color="1F4E79")
                header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
                data_font   = Font(name="Arial", size=10)

                # Write header row — ALL 43 columns always present
                for c_idx, h in enumerate(HEADERS, 1):
                    cell = ws.cell(1, c_idx, h)
                    cell.fill      = header_fill
                    cell.font      = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    ws.column_dimensions[get_column_letter(c_idx)].width = COL_WIDTHS.get(h, 12)
                ws.row_dimensions[1].height = 20

                date_cols = {"start_date", "end_date", "created_date", "edited_date", "cutoff_date"}

                # Write data rows — ALL 43 columns every row, empty if no value
                for r_idx, row in enumerate(rows, 2):
                    for c_idx, h in enumerate(HEADERS, 1):
                        val  = row.get(h)          # None if missing → empty cell
                        cell = ws.cell(r_idx, c_idx, val)
                        cell.font = data_font
                        if h in date_cols and val:
                            cell.number_format = "DD/MM/YYYY"
                        if h == "net_price" and val is not None:
                            cell.value = str(val)  # Keep as string per dashboard requirement

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)

                st.success(f"✅ สำเร็จ! เหมียวดึงข้อมูลได้ **{len(rows)}** แถว ({len(parsed_data.get('periods',[]))} periods × {len(valid_rooms)} rooms)")
                st.balloons()
                st.download_button(
                    label="📥  Download Excel (.xlsx)",
                    data=buf,
                    file_name=f"Contract_{hotel_id}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

