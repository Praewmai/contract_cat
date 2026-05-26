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

# Custom Styling for Dark UI
st.markdown("""
<style>
    .main {
        background-color: #0f172a;
    }
    div[data-testid="stMarkdownContainer"] p {
        font-family: 'Outfit', sans-serif;
    }
</style>
""", unsafe_allow_html=True)

st.title("Contract PDF → Excel  🐾")
st.write("Upload a contract PDF and let the AI Cat extract rates, dates, and terms into a dashboard-ready Excel file.")

# 1. API Key
api_key = st.text_input("🔑 Gemini API Key", type="password", help="Your API key is used client-side and not saved on our servers.")

# 2. PDF Upload
uploaded_file = st.file_uploader("📄 Upload Contract PDF", type=["pdf"])

# 3. Property Type
prop_type = st.radio("🏢 Property Type", ["Hotel / Resort", "Cruise Ship"])

# 4. Info Grid
col1, col2 = st.columns(2)
with col1:
    hotel_id = st.text_input("Hotel/Cruise ID (e.g. 12711588)")
with col2:
    supplier = st.text_input("Supplier Name")

# 5. Rooms Config
st.subheader("🚪 Rooms Configuration")

# Initialize rooms list in session state
if 'rooms' not in st.session_state:
    st.session_state.rooms = [{"room_name": "", "room_id": ""}]

# Callback to add a room row
def add_room_callback():
    st.session_state.rooms.append({"room_name": "", "room_id": ""})

# Callback to remove a room row
def remove_room_callback(index):
    if len(st.session_state.rooms) > 1:
        st.session_state.rooms.pop(index)

# Render each room row
for idx in range(len(st.session_state.rooms)):
    # Safely avoid index out of range if deleted during iteration
    if idx >= len(st.session_state.rooms):
        break
    room = st.session_state.rooms[idx]
    
    r_col1, r_col2, r_col3 = st.columns([5, 4, 1])
    with r_col1:
        room["room_name"] = st.text_input(f"Room Name #{idx+1}", value=room["room_name"], key=f"rn_{idx}")
    with r_col2:
        room["room_id"] = st.text_input(f"Room ID #{idx+1}", value=room["room_id"], key=f"ri_{idx}")
    with r_col3:
        st.write("##")  # spacer to align button with text inputs
        if st.button("🗑️", key=f"del_{idx}"):
            remove_room_callback(idx)
            st.rerun()

st.button("+ Add Room Row", on_click=add_room_callback)

# 6. Contract Types
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

valid_rooms = [r for r in st.session_state.rooms if r["room_name"].strip() and r["room_id"].strip()]

# 7. Start Analysis Button
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
