import streamlit as st
import json
import base64
import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import io

# Page Config
st.set_page_config(
    page_title="AI Cat Assistant | Contract PDF → Excel",
    page_icon="🐾",
    layout="centered"
)

# Custom Styling for Premium UI
st.markdown("""
<style>
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Premium Dark Mesh Gradient Background */
    .stApp {
        background: radial-gradient(circle at 15% 50%, rgba(99, 102, 241, 0.15) 0%, transparent 50%),
                    radial-gradient(circle at 85% 30%, rgba(16, 185, 129, 0.1) 0%, transparent 50%),
                    #0f172a !important;
        color: #f8fafc;
    }
    
    /* Typography */
    h1, h2, h3, h4, h5, h6, p, span, div, label {
        font-family: 'Outfit', sans-serif !important;
    }
    
    h1 {
        background: linear-gradient(135deg, #818cf8 0%, #e0e7ff 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        font-weight: 800 !important;
        letter-spacing: -1px;
        margin-bottom: 2rem !important;
        text-shadow: 0 4px 20px rgba(99, 102, 241, 0.2);
    }
    
    h2, h3 {
        color: #e2e8f0 !important;
        font-weight: 700 !important;
        border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        padding-bottom: 10px;
        margin-top: 2rem !important;
    }

    /* Glassmorphism Containers - Target main blocks */
    [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {
        background: rgba(30, 41, 59, 0.6) !important;
        backdrop-filter: blur(16px) !important;
        -webkit-backdrop-filter: blur(16px) !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        border-radius: 24px !important;
        padding: 32px !important;
        box-shadow: 0 20px 40px -10px rgba(0, 0, 0, 0.3) !important;
    }

    /* Inputs - Glassmorphism & Neumorphism blend */
    .stTextInput > div > div > input {
        border-radius: 16px;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        background-color: rgba(15, 23, 42, 0.6) !important;
        color: #f8fafc !important;
        padding: 12px 20px;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: inset 0 2px 4px rgba(0,0,0,0.2);
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #818cf8 !important;
        box-shadow: 0 0 0 3px rgba(129, 140, 248, 0.2) !important;
        background-color: rgba(15, 23, 42, 0.9) !important;
    }
    
    /* Premium File Uploader */
    [data-testid="stFileUploader"] {
        background-color: rgba(30, 41, 59, 0.4) !important;
        border-radius: 20px !important;
        padding: 24px !important;
        border: 2px dashed rgba(129, 140, 248, 0.3) !important;
        text-align: center;
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #818cf8 !important;
        background-color: rgba(30, 41, 59, 0.6) !important;
        box-shadow: 0 8px 24px rgba(0,0,0,0.2);
    }

    /* Buttons - Glowing Gradient */
    .stButton > button {
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
        color: #ffffff !important;
        border-radius: 100px !important;
        border: none !important;
        padding: 12px 32px !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        box-shadow: 0 10px 20px -5px rgba(99, 102, 241, 0.5) !important;
        width: 100%;
    }
    
    .stButton > button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 15px 25px -5px rgba(99, 102, 241, 0.6) !important;
        background: linear-gradient(135deg, #818cf8 0%, #6366f1 100%) !important;
    }
    
    .stButton > button:active {
        transform: translateY(0) !important;
        box-shadow: 0 5px 10px -5px rgba(99, 102, 241, 0.5) !important;
    }

    /* Custom Checkboxes and Radios */
    .stCheckbox p, .stRadio p {
        color: #cbd5e1 !important;
        font-weight: 500 !important;
    }
    
    /* Remove weird backgrounds from Streamlit's default radio/checkbox wrappers */
    div[data-baseweb="radio"] > div, div[data-baseweb="checkbox"] > div {
        background-color: transparent !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="card">', unsafe_allow_html=True)
st.title("Contract PDF → Excel  🐾")
st.write("Upload a contract PDF and let the AI Cat extract rates, dates, and terms into a dashboard-ready Excel file.")
st.markdown('</div>', unsafe_allow_html=True)

# 1. API Key
st.markdown('<div class="card">', unsafe_allow_html=True)
api_key = st.text_input("🔑 Gemini API Key", type="password", help="Your API key is used client-side and not saved on our servers.")
st.markdown('</div>', unsafe_allow_html=True)

# 2. PDF Upload
st.markdown('<div class="card">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("📄 Upload Contract PDF", type=["pdf"]) 
st.markdown('</div>', unsafe_allow_html=True)

# 3. Property Type
st.markdown('<div class="card">', unsafe_allow_html=True)
prop_type = st.radio("🏢 Property Type", ["Hotel / Resort", "Cruise Ship"]) 
st.markdown('</div>', unsafe_allow_html=True)

# 4. Info Grid
st.markdown('<div class="card">', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    hotel_id = st.text_input("Hotel/Cruise ID (e.g. 12711588)")
with col2:
    supplier = st.text_input("Supplier Name")
st.markdown('</div>', unsafe_allow_html=True)

# 5. Rooms Config
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("🚪 Rooms Configuration")
if 'rooms' not in st.session_state:
    st.session_state.rooms = [{"room_name": "", "room_id": ""}]

def add_room_callback():
    st.session_state.rooms.append({"room_name": "", "room_id": ""})

def remove_room_callback(index):
    if len(st.session_state.rooms) > 1:
        st.session_state.rooms.pop(index)
for idx in range(len(st.session_state.rooms)):
    if idx >= len(st.session_state.rooms):
        break
    room = st.session_state.rooms[idx]
    r_col1, r_col2, r_col3 = st.columns([5, 4, 1])
    with r_col1:
        room["room_name"] = st.text_input(f"Room Name #{idx+1}", value=room["room_name"], key=f"rn_{idx}")
    with r_col2:
        room["room_id"] = st.text_input(f"Room ID #{idx+1}", value=room["room_id"], key=f"ri_{idx}")
    with r_col3:
        st.write("##")
        if st.button("🗑️", key=f"del_{idx}"):
            remove_room_callback(idx)
            st.rerun()
st.button("+ Add Room Row", on_click=add_room_callback)
st.markdown('</div>', unsafe_allow_html=True)

# 6. Contract Types
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("📝 Contract Types to Generate")
ct_main = st.checkbox("Main Contract (Base Rate)", value=True)

# Early bird options
ct_eb = st.checkbox("Early Bird (Advance Booking)", value=False)
eb_code = "E.B DAYS"
if ct_eb:
    eb_code = st.text_input("Early Bird Promo Prefix (e.g. E.B DAYS)", value="E.B DAYS")

# Promo options
ct_promo = st.checkbox("Promotion (Special Offer)", value=False)
promo_code = ""
promo_till = ""
if ct_promo:
    p_col1, p_col2 = st.columns(2)
    with p_col1:
        promo_code = st.text_input("Promo Code (e.g. FLASH2026)")
    with p_col2:
        promo_till = st.text_input("Book Till Date (e.g. YYYY-MM-DD)")

ct_por = st.checkbox("POR (Price on Request)", value=True)

# Compile selected contract list
cts = []
if ct_main:
    cts.append({"type": "main", "label": "Main Contract"})
if ct_eb:
    cts.append({"type": "eb", "label": "Early Bird", "code": eb_code})
if ct_promo:
    cts.append({"type": "promo", "label": "Promotion", "promo_code": promo_code, "promo_till": promo_till})
if ct_por:
    cts.append({"type": "por", "label": "POR"})
st.markdown('</div>', unsafe_allow_html=True)

valid_rooms = [r for r in st.session_state.rooms if r["room_name"].strip() and r["room_id"].strip()]

# 7. Start Analysis Button
st.markdown('<div class="card">', unsafe_allow_html=True)
if st.button("Let the AI Cat Extract Data! 🐾", use_container_width=True):
    if not api_key:
        st.error("Please enter your Gemini API Key.")
    elif not uploaded_file:
        st.error("Please upload a PDF file.")
    elif not hotel_id:
        st.error("Please enter a Hotel ID.")
    elif not valid_rooms:
        st.error("Please add at least one Room Name & Room ID pair.")
    elif not cts:
        st.error("Please select at least one contract type.")
    else:
        with st.spinner("AI is reading the contract and extracting rate entries... (may take up to 60s)"):
            try:
                # Read file bytes and encode to base64
                file_bytes = uploaded_file.read()
                b64_data = base64.b64encode(file_bytes).decode("utf-8")
                
                room_hint = f"Use exactly these rooms (match strictly by room_name to get room_id): {json.dumps(valid_rooms)}"
                
                ct_desc_parts = []
                for c in cts:
                    if c["type"] == "main":
                        ct_desc_parts.append("Main Contract: standard net rates.")
                    elif c["type"] == "eb":
                        ct_desc_parts.append(f"Early Bird: detect ALL tiers. promo_code format \"{c['code']} X DAYS\". Handle blackout dates carefully.")
                    elif c["type"] == "promo":
                        ct_desc_parts.append(f"Promotion: promo_code=\"{c['promo_code']}\", promo_book_till=\"{c['promo_till']}\".")
                    elif c["type"] == "por":
                        ct_desc_parts.append("POR: net_price=0, promo_code=\"POR Rate\".")
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
                
                # Structuring the Excel workbook in memory
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
                        r.get("hotel_name", ""),
                        r.get("hotel_id", ""),
                        r.get("hotel_supplier", ""),
                        r.get("room_id", ""),
                        r.get("room_name", ""),
                        r.get("start_date", ""),
                        r.get("end_date", ""),
                        r.get("contract_type", ""),
                        r.get("net_price", 0),
                        r.get("promo_code", ""),
                        r.get("promo_book_till", ""),
                        r.get("min_advance_days", 0),
                        r.get("min_nights_stay", 0),
                        r.get("promo_note", "")
                    ])
                    
                for col in ws.columns:
                    max_len = 0
                    for cell in col:
                        val_str = str(cell.value or "")
                        if len(val_str) > max_len:
                            max_len = len(val_str)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
                
                excel_data = io.BytesIO()
                wb.save(excel_data)
                excel_data.seek(0)
                
                st.success(f"Success! Extracted {len(rows)} rows.")
                
                st.download_button(
                    label="📥 Download Excel (.xlsx)",
                    data=excel_data,
                    file_name=f"Contract_Rates_Extraction_{hotel_id}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"Analysis failed: {str(e)}")
st.markdown('</div>', unsafe_allow_html=True)
