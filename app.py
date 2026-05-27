import streamlit as st
import json
import base64
import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import io
import os

# ─── Page Config ───
st.set_page_config(
    page_title="AI Cat Assistant 🐾 Contract PDF → Excel",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── Base64 Mascot ───
def get_base64_image(file_path):
    if os.path.exists(file_path):
        with open(file_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode('utf-8')
    return ""

mascot_b64 = get_base64_image("mascot.png")
mascot_img_html = f'<img src="data:image/png;base64,{mascot_b64}" class="mascot-avatar" />' if mascot_b64 else '<div class="mascot-avatar">🐾</div>'

# ─── Advanced Premium CSS ───
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
    /* Global Streamlit overrides */
    #MainMenu, footer, header {visibility: hidden !important;}
    .block-container { padding: 2rem 3rem !important; }
    
    html, body, [class*="css"], .stMarkdown, p, span, label, h1, h2, h3, h4, h5, h6, div, input, button, textarea {
        font-family: 'Prompt', sans-serif !important;
    }

    /* Main background */
    .stApp {
        background-color: #F8F9FA !important;
        background-image: 
            radial-gradient(circle at 10% 20%, rgba(74, 88, 153, 0.05) 0%, transparent 20%),
            radial-gradient(circle at 90% 80%, rgba(74, 88, 153, 0.05) 0%, transparent 20%);
    }

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E9ECEF !important;
        box-shadow: 4px 0 24px rgba(0,0,0,0.02) !important;
    }
    
    .sidebar-profile {
        text-align: center;
        padding: 1rem 0 2rem 0;
        border-bottom: 1px dashed #E9ECEF;
        margin-bottom: 2rem;
    }
    .mascot-avatar {
        width: 120px;
        height: 120px;
        border-radius: 50%;
        object-fit: cover;
        background: #F8F9FA;
        border: 4px solid #FFFFFF;
        box-shadow: 0 8px 24px rgba(74, 88, 153, 0.15);
        margin: 0 auto 1rem auto;
        display: block;
        animation: float 6s ease-in-out infinite;
    }
    @keyframes float {
        0% { transform: translateY(0px); }
        50% { transform: translateY(-8px); }
        100% { transform: translateY(0px); }
    }
    .sidebar-title {
        color: #2B3A67;
        font-weight: 700;
        font-size: 1.2rem;
        margin: 0;
    }
    .sidebar-subtitle {
        color: #6C757D;
        font-size: 0.9rem;
        font-weight: 400;
    }

    /* Chat Bubble */
    .chat-bubble {
        background: #FFFFFF;
        border-radius: 20px;
        padding: 1.5rem 2rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.04);
        position: relative;
        margin-bottom: 2rem;
        border-left: 6px solid #4A5899;
        display: flex;
        align-items: center;
        gap: 1.5rem;
    }
    .chat-bubble h3 {
        color: #2B3A67 !important;
        margin: 0 0 0.5rem 0 !important;
        font-size: 1.4rem !important;
    }
    .chat-bubble p {
        color: #495057 !important;
        margin: 0 !important;
        font-size: 1rem !important;
    }

    /* Cards */
    .dashboard-card {
        background: #FFFFFF;
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        border: 1px solid #F1F3F5;
        height: 100%;
    }
    
    /* Inputs */
    .stTextInput > div > div > input {
        border-radius: 10px !important;
        border: 1px solid #DEE2E6 !important;
        padding: 0.6rem 1rem !important;
        background-color: #F8F9FA !important;
        color: #495057 !important;
        transition: all 0.2s;
    }
    .stTextInput > div > div > input:focus {
        border-color: #4A5899 !important;
        background-color: #FFFFFF !important;
        box-shadow: 0 0 0 3px rgba(74,88,153,0.1) !important;
    }

    /* Uploader */
    [data-testid="stFileUploaderDropzone"] {
        background-color: #F8F9FA !important;
        border: 2px dashed #CED4DA !important;
        border-radius: 16px !important;
        padding: 2rem !important;
        transition: all 0.3s;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #4A5899 !important;
        background-color: #F4F6FB !important;
    }

    /* Primary Button */
    div[data-testid="stButton"] > button {
        background: linear-gradient(135deg, #4A5899 0%, #3B477A 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.8rem 2rem !important;
        font-weight: 600 !important;
        width: 100%;
        box-shadow: 0 4px 12px rgba(74,88,153,0.2) !important;
        transition: all 0.3s !important;
    }
    div[data-testid="stButton"] > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(74,88,153,0.3) !important;
    }
</style>
""", unsafe_allow_html=True)


# ─── Sidebar configuration ───
with st.sidebar:
    st.markdown(f"""
    <div class="sidebar-profile">
        {mascot_img_html}
        <h2 class="sidebar-title">AI Cat Assistant</h2>
        <p class="sidebar-subtitle">Contract PDF Extractor</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### ⚙️ System Config")
    api_key = st.text_input("Gemini API Key", type="password", placeholder="Paste your API key...")
    
    st.markdown("---")
    st.markdown("### 🏨 Property Details")
    prop_type = st.radio("Type", ["Hotel / Resort", "Cruise Ship"])
    hotel_id = st.text_input("Property ID", placeholder="e.g. 12711588")
    supplier = st.text_input("Supplier", placeholder="e.g. Agoda")


# ─── Main Content ───
st.markdown("""
<div class="chat-bubble">
    <div style="font-size: 3rem;">👋</div>
    <div>
        <h3>สวัสดีชาว Hotelier! เมี๊ยว~</h3>
        <p>อัพโหลดไฟล์ PDF Contract สัญญาโรงแรมด้านล่างนี้เลย เดี๋ยวหนูจะช่วยสกัดข้อมูล Rate, Date, Terms ทั้งหมดให้ออกมาเป็น Excel สวยๆ พร้อมเอาไปใช้งานต่อทันทีค่ะ!</p>
    </div>
</div>
""", unsafe_allow_html=True)

col_left, col_right = st.columns([1, 1.2], gap="large")

with col_left:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.markdown("#### 📄 1. Upload Contract")
    uploaded_file = st.file_uploader("Select PDF file", type=["pdf"], label_visibility="collapsed")
    
    st.markdown("#### 📝 2. Contract Types")
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
    st.markdown('</div>', unsafe_allow_html=True)

with col_right:
    st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
    st.markdown("#### 🚪 3. Rooms Configuration")
    
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
            if st.button("🗑️", key=f"del_{idx}", help="Remove Room"):
                remove_room(idx)
                st.rerun()
                
    st.button("＋ Add Room", on_click=add_room, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


# ─── Data Extraction Logic ───
st.markdown("<br>", unsafe_allow_html=True)

valid_rooms = [r for r in st.session_state.rooms if r["room_name"].strip() and r["room_id"].strip()]

cts = []
if ct_main: cts.append({"type": "main", "label": "Main Contract"})
if ct_eb: cts.append({"type": "eb", "label": "Early Bird", "code": eb_code})
if ct_promo: cts.append({"type": "promo", "label": "Promotion", "promo_code": promo_code, "promo_till": promo_till})
if ct_por: cts.append({"type": "por", "label": "POR"})

if st.button("🚀 Process Contract & Generate Excel", use_container_width=True):
    if not api_key:
        st.error("⚠️ Please enter your Gemini API Key in the sidebar.")
    elif not uploaded_file:
        st.error("⚠️ Please upload a PDF file.")
    elif not hotel_id:
        st.error("⚠️ Please enter a Property ID in the sidebar.")
    elif not valid_rooms:
        st.error("⚠️ Please add at least one Room Name & Room ID pair.")
    elif not cts:
        st.error("⚠️ Please select at least one contract type.")
    else:
        with st.spinner("🐱 น้องแมวกำลังอ่าน contract อยู่นะ… รอสักครู่ (ไม่เกิน 60 วินาที)"):
            try:
                file_bytes = uploaded_file.read()
                b64_data = base64.b64encode(file_bytes).decode("utf-8")

                room_hint = f"Use exactly these rooms (match strictly by room_name to get room_id): {json.dumps(valid_rooms)}"
                ct_desc_parts = []
                for c in cts:
                    if c["type"] == "main": ct_desc_parts.append("Main Contract: standard net rates.")
                    elif c["type"] == "eb": ct_desc_parts.append(f'Early Bird: detect ALL tiers. promo_code format "{c["code"]} X DAYS". Handle blackout dates carefully.')
                    elif c["type"] == "promo": ct_desc_parts.append(f'Promotion: promo_code="{c["promo_code"]}", promo_book_till="{c["promo_till"]}".')
                    elif c["type"] == "por": ct_desc_parts.append('POR: net_price=0, promo_code="POR Rate".')
                ct_desc = " | ".join(ct_desc_parts)

                prompt_text = f"""You are a data extraction expert. Analyze the hotel contract PDF and extract rates into an Excel-ready flat array.
Rules:
1. Surcharges/Min Nights/Blackout dates: auto-detect and apply to the appropriate date ranges. Split periods if necessary.
2. {room_hint}
3. Contract types to generate: {ct_desc}
4. For every single rate combination (room + period + contract_type), generate a row.

Return ONLY a JSON object with this EXACT structure (no markdown fences, just the JSON string):
{{
  "hotel_name": "string",
  "hotel_id": "{hotel_id}",
  "hotel_supplier": "{supplier}",
  "rooms_count": {len(valid_rooms)},
  "periods_count": 0,
  "all_rows": [
    {{
      "hotel_name": "string",
      "hotel_id": "{hotel_id}",
      "hotel_supplier": "{supplier}",
      "room_id": "string",
      "room_name": "string",
      "start_date": "YYYY-MM-DD",
      "end_date": "YYYY-MM-DD",
      "contract_type": "string (e.g. Main Contract, Early Bird)",
      "net_price": 0,
      "promo_code": "string",
      "promo_book_till": "YYYY-MM-DD or empty",
      "min_advance_days": 0,
      "min_nights_stay": 0,
      "promo_note": "string (surcharges, inclusions, weekday/weekend)"
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
                headers = {"Content-Type": "application/json"}
                resp = requests.post(url, json=payload, headers=headers, timeout=120)
                data = resp.json()

                if resp.status_code != 200:
                    raise Exception(data.get("error", {}).get("message", "Gemini API Error"))

                text_output = data["candidates"][0]["content"]["parts"][0]["text"]
                result_json = json.loads(text_output)
                rows = result_json.get("all_rows", [])

                wb = Workbook()
                ws = wb.active
                ws.title = "Extracted Rates"
                headers_excel = [
                    "Hotel Name", "Hotel ID", "Supplier", "Room ID", "Room Name",
                    "Start Date", "End Date", "Contract Type", "Net Price",
                    "Promo Code", "Book Till", "Min Advance Days", "Min Nights", "Notes"
                ]
                ws.append(headers_excel)
                for r in rows:
                    ws.append([
                        r.get("hotel_name", ""), r.get("hotel_id", ""), r.get("hotel_supplier", ""),
                        r.get("room_id", ""), r.get("room_name", ""), r.get("start_date", ""),
                        r.get("end_date", ""), r.get("contract_type", ""), r.get("net_price", 0),
                        r.get("promo_code", ""), r.get("promo_book_till", ""), r.get("min_advance_days", 0),
                        r.get("min_nights_stay", 0), r.get("promo_note", "")
                    ])

                for col in ws.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 10)

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)

                st.success(f"✅ สำเร็จ! ดึงข้อมูลได้ **{len(rows)}** แถวจาก contract")
                st.balloons()
                st.download_button(
                    label="📥  Download Excel (.xlsx)",
                    data=buf,
                    file_name=f"Contract_Rates_{hotel_id}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
