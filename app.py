import streamlit as st
import json
import base64
import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import io

# ─── Page Config ───
st.set_page_config(
    page_title="AI Cat Assistant | Contract PDF → Excel",
    page_icon="🐾",
    layout="centered"
)

# ─── Premium CSS (targets actual Streamlit DOM) ───
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
    /* ── Hide Streamlit chrome ── */
    #MainMenu, footer, header {visibility: hidden;}

    /* ── Root app background ── */
    .stApp {
        background:
            radial-gradient(ellipse at 15% 50%, rgba(99,102,241,0.18) 0%, transparent 55%),
            radial-gradient(ellipse at 85% 25%, rgba(16,185,129,0.12) 0%, transparent 55%),
            radial-gradient(ellipse at 50% 95%, rgba(244,114,182,0.08) 0%, transparent 50%),
            #0f172a !important;
    }

    /* ── Global typography ── */
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif !important;
    }

    /* ── Animated gradient title ── */
    h1 {
        background: linear-gradient(135deg, #a5b4fc, #818cf8, #c4b5fd, #e0e7ff);
        background-size: 300% 300%;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: gradient-shift 6s ease infinite;
        text-align: center !important;
        font-weight: 800 !important;
        font-size: 2.4rem !important;
        letter-spacing: -0.5px;
        padding-bottom: 0.5rem;
    }
    @keyframes gradient-shift {
        0%   { background-position: 0% 50%; }
        50%  { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* ── Subheaders ── */
    h2, h3 {
        color: #c7d2fe !important;
        font-weight: 700 !important;
    }

    /* ── All text ── */
    p, span, label, .stMarkdown {
        color: #e2e8f0 !important;
    }

    /* ── Container / expander cards ── */
    [data-testid="stExpander"] {
        background: rgba(30, 41, 59, 0.65) !important;
        backdrop-filter: blur(20px) !important;
        -webkit-backdrop-filter: blur(20px) !important;
        border: 1px solid rgba(255,255,255,0.07) !important;
        border-radius: 20px !important;
        margin-bottom: 1rem !important;
        box-shadow: 0 8px 32px rgba(0,0,0,0.25) !important;
        overflow: hidden;
    }
    [data-testid="stExpander"] summary {
        font-weight: 600 !important;
        font-size: 1.05rem !important;
        color: #e2e8f0 !important;
        padding: 1rem 1.25rem !important;
    }
    [data-testid="stExpander"] summary:hover {
        color: #a5b4fc !important;
    }
    [data-testid="stExpander"] [data-testid="stExpanderDetails"] {
        padding: 0 1.25rem 1.25rem !important;
    }

    /* ── Text inputs ── */
    .stTextInput > div > div > input {
        border-radius: 14px !important;
        border: 1px solid rgba(129,140,248,0.25) !important;
        background: rgba(15, 23, 42, 0.7) !important;
        color: #f8fafc !important;
        padding: 0.7rem 1rem !important;
        font-family: 'Outfit', sans-serif !important;
        transition: all 0.3s ease !important;
    }
    .stTextInput > div > div > input:focus {
        border-color: #818cf8 !important;
        box-shadow: 0 0 0 3px rgba(129,140,248,0.15), 0 0 20px rgba(129,140,248,0.1) !important;
    }
    .stTextInput label {
        color: #a5b4fc !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
    }

    /* ── File uploader ── */
    [data-testid="stFileUploader"] {
        background: rgba(30, 41, 59, 0.4) !important;
        border: 2px dashed rgba(129,140,248,0.3) !important;
        border-radius: 18px !important;
        padding: 1.5rem !important;
        transition: all 0.3s ease !important;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #818cf8 !important;
        background: rgba(30, 41, 59, 0.6) !important;
        box-shadow: 0 0 25px rgba(129,140,248,0.1) !important;
    }
    [data-testid="stFileUploader"] label {
        color: #a5b4fc !important;
        font-weight: 600 !important;
    }

    /* ── Radio buttons ── */
    .stRadio label {
        color: #a5b4fc !important;
        font-weight: 600 !important;
    }
    .stRadio [data-baseweb="radio"] label {
        color: #cbd5e1 !important;
        font-weight: 400 !important;
    }

    /* ── Checkboxes ── */
    .stCheckbox label {
        color: #cbd5e1 !important;
        font-weight: 500 !important;
    }

    /* ── Primary action button ── */
    .stButton > button[kind="primary"],
    .stButton > button {
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #6366f1 100%) !important;
        background-size: 200% 200% !important;
        color: #fff !important;
        border: none !important;
        border-radius: 100px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        font-family: 'Outfit', sans-serif !important;
        letter-spacing: 0.3px;
        box-shadow: 0 8px 25px -5px rgba(99,102,241,0.45) !important;
        transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }
    .stButton > button:hover {
        background-position: 100% 50% !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 12px 30px -5px rgba(99,102,241,0.55) !important;
    }
    .stButton > button:active {
        transform: translateY(0px) !important;
    }

    /* ── Download button ── */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #10b981, #059669) !important;
        color: #fff !important;
        border: none !important;
        border-radius: 100px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 700 !important;
        box-shadow: 0 8px 25px -5px rgba(16,185,129,0.4) !important;
        transition: all 0.35s ease !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 12px 30px -5px rgba(16,185,129,0.5) !important;
    }

    /* ── Success / Error messages ── */
    .stSuccess {
        background: rgba(16,185,129,0.15) !important;
        border: 1px solid rgba(16,185,129,0.3) !important;
        border-radius: 14px !important;
    }
    .stError, [data-testid="stAlert"] {
        border-radius: 14px !important;
    }

    /* ── Spinner ── */
    .stSpinner > div {
        border-top-color: #818cf8 !important;
    }

    /* ── Scrollbar ── */
    ::-webkit-scrollbar { width: 8px; }
    ::-webkit-scrollbar-track { background: #0f172a; }
    ::-webkit-scrollbar-thumb { background: #334155; border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: #475569; }

    /* ── Sidebar (if ever used) ── */
    [data-testid="stSidebar"] {
        background: rgba(15,23,42,0.95) !important;
        backdrop-filter: blur(20px) !important;
    }

    /* ── Divider ── */
    hr {
        border-color: rgba(255,255,255,0.06) !important;
        margin: 1.5rem 0 !important;
    }

    /* ── Columns gap ── */
    [data-testid="column"] {
        padding: 0 0.5rem !important;
    }
</style>
""", unsafe_allow_html=True)

# ─── Header ───
st.title("Contract PDF → Excel  🐾")
st.markdown(
    "<p style='text-align:center; color:#94a3b8 !important; font-size:1.05rem; margin-top:-1rem; margin-bottom:2rem;'>"
    "Upload a contract PDF and let the AI Cat extract rates, dates, and terms into a dashboard-ready Excel file."
    "</p>",
    unsafe_allow_html=True
)

# ─── 1. API Key ───
with st.expander("🔑 API Key", expanded=True):
    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        help="Your API key is used client-side only and never saved.",
        label_visibility="collapsed",
        placeholder="Paste your Gemini API key here…"
    )

# ─── 2. PDF Upload ───
with st.expander("📄 Upload Contract PDF", expanded=True):
    uploaded_file = st.file_uploader(
        "Upload PDF",
        type=["pdf"],
        label_visibility="collapsed"
    )

# ─── 3. Property Info ───
with st.expander("🏢 Property Information", expanded=True):
    prop_type = st.radio(
        "Property Type",
        ["Hotel / Resort", "Cruise Ship"],
        horizontal=True
    )
    st.markdown("")  # spacer
    col1, col2 = st.columns(2)
    with col1:
        hotel_id = st.text_input("Hotel / Cruise ID", placeholder="e.g. 12711588")
    with col2:
        supplier = st.text_input("Supplier Name", placeholder="e.g. Agoda, Expedia")

# ─── 4. Rooms ───
with st.expander("🚪 Rooms Configuration", expanded=True):
    if 'rooms' not in st.session_state:
        st.session_state.rooms = [{"room_name": "", "room_id": ""}]

    def add_room():
        st.session_state.rooms.append({"room_name": "", "room_id": ""})

    def remove_room(i):
        if len(st.session_state.rooms) > 1:
            st.session_state.rooms.pop(i)

    for idx in range(len(st.session_state.rooms)):
        if idx >= len(st.session_state.rooms):
            break
        room = st.session_state.rooms[idx]
        c1, c2, c3 = st.columns([5, 4, 1])
        with c1:
            room["room_name"] = st.text_input(
                f"Room Name #{idx+1}", value=room["room_name"], key=f"rn_{idx}",
                placeholder="e.g. Deluxe Twin"
            )
        with c2:
            room["room_id"] = st.text_input(
                f"Room ID #{idx+1}", value=room["room_id"], key=f"ri_{idx}",
                placeholder="e.g. 101"
            )
        with c3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️", key=f"del_{idx}"):
                remove_room(idx)
                st.rerun()

    st.button("＋ Add Room", on_click=add_room, use_container_width=True)

# ─── 5. Contract Types ───
with st.expander("📝 Contract Types", expanded=True):
    ct_main = st.checkbox("Main Contract (Base Rate)", value=True)
    ct_eb = st.checkbox("Early Bird (Advance Booking)", value=False)
    eb_code = "E.B DAYS"
    if ct_eb:
        eb_code = st.text_input("Early Bird Promo Prefix", value="E.B DAYS", placeholder="e.g. E.B DAYS")

    ct_promo = st.checkbox("Promotion (Special Offer)", value=False)
    promo_code = ""
    promo_till = ""
    if ct_promo:
        pc1, pc2 = st.columns(2)
        with pc1:
            promo_code = st.text_input("Promo Code", placeholder="e.g. FLASH2026")
        with pc2:
            promo_till = st.text_input("Book Till Date", placeholder="YYYY-MM-DD")

    ct_por = st.checkbox("POR (Price on Request)", value=True)

    cts = []
    if ct_main:
        cts.append({"type": "main", "label": "Main Contract"})
    if ct_eb:
        cts.append({"type": "eb", "label": "Early Bird", "code": eb_code})
    if ct_promo:
        cts.append({"type": "promo", "label": "Promotion", "promo_code": promo_code, "promo_till": promo_till})
    if ct_por:
        cts.append({"type": "por", "label": "POR"})

# ─── Collect valid rooms ───
valid_rooms = [r for r in st.session_state.rooms if r["room_name"].strip() and r["room_id"].strip()]

# ─── 6. Extract Button ───
st.markdown("")  # spacer
if st.button("🐾  Let the AI Cat Extract Data!", use_container_width=True):
    if not api_key:
        st.error("⚠️ Please enter your Gemini API Key.")
    elif not uploaded_file:
        st.error("⚠️ Please upload a PDF file.")
    elif not hotel_id:
        st.error("⚠️ Please enter a Hotel / Cruise ID.")
    elif not valid_rooms:
        st.error("⚠️ Please add at least one Room Name & Room ID pair.")
    elif not cts:
        st.error("⚠️ Please select at least one contract type.")
    else:
        with st.spinner("✨ AI is reading the contract and extracting rate entries… (up to 60 s)"):
            try:
                file_bytes = uploaded_file.read()
                b64_data = base64.b64encode(file_bytes).decode("utf-8")

                room_hint = f"Use exactly these rooms (match strictly by room_name to get room_id): {json.dumps(valid_rooms)}"

                ct_desc_parts = []
                for c in cts:
                    if c["type"] == "main":
                        ct_desc_parts.append("Main Contract: standard net rates.")
                    elif c["type"] == "eb":
                        ct_desc_parts.append(f'Early Bird: detect ALL tiers. promo_code format "{c["code"]} X DAYS". Handle blackout dates carefully.')
                    elif c["type"] == "promo":
                        ct_desc_parts.append(f'Promotion: promo_code="{c["promo_code"]}", promo_book_till="{c["promo_till"]}".')
                    elif c["type"] == "por":
                        ct_desc_parts.append('POR: net_price=0, promo_code="POR Rate".')
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

                # ── Build Excel ──
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
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 10)

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)

                st.success(f"✅ Success! Extracted **{len(rows)}** rows from the contract.")
                st.download_button(
                    label="📥  Download Excel (.xlsx)",
                    data=buf,
                    file_name=f"Contract_Rates_{hotel_id}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ Analysis failed: {str(e)}")

# ─── Footer ───
st.markdown("---")
st.markdown(
    "<p style='text-align:center; color:#475569 !important; font-size:0.85rem;'>"
    "Made with 🐾 by AI Cat Assistant &nbsp;•&nbsp; Powered by Gemini 2.5 Flash"
    "</p>",
    unsafe_allow_html=True
)
