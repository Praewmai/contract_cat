import streamlit as st
import json
import base64
import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import io

# ─── Page Config ───
st.set_page_config(
    page_title="AI Cat Assistant 🐾 Contract PDF → Excel",
    page_icon="🐾",
    layout="centered"
)

# ─── Cat-Themed Premium CSS ───
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
    /* ── Hide Streamlit UI chrome ── */
    #MainMenu, footer, header {visibility: hidden !important;}
    .block-container { padding-top: 2rem !important; }

    /* ── Warm cream background with subtle pattern ── */
    .stApp {
        background-color: #FDF8F3 !important;
        background-image:
            radial-gradient(circle at 20% 80%, rgba(74,88,153,0.04) 0%, transparent 50%),
            radial-gradient(circle at 80% 20%, rgba(245,168,176,0.06) 0%, transparent 50%),
            radial-gradient(circle at 50% 50%, rgba(245,200,66,0.03) 0%, transparent 60%);
    }

    /* ── Global font ── */
    html, body, [class*="css"], .stMarkdown, p, span, label, h1, h2, h3, h4, h5, h6, div, input, button, textarea {
        font-family: 'Nunito', sans-serif !important;
    }

    /* ── Custom header banner ── */
    .cat-banner {
        background: linear-gradient(135deg, #4A5899 0%, #6B7DB8 60%, #8B9DD6 100%);
        border-radius: 24px;
        padding: 2.5rem 2rem 2rem;
        text-align: center;
        margin-bottom: 1.5rem;
        position: relative;
        overflow: hidden;
        box-shadow: 0 8px 30px rgba(74,88,153,0.2);
    }
    .cat-banner::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 60%);
        animation: shimmer 8s ease-in-out infinite;
    }
    @keyframes shimmer {
        0%, 100% { transform: translate(0, 0); }
        50% { transform: translate(10%, 10%); }
    }
    .cat-banner h2 {
        color: #FFFFFF !important;
        font-size: 2rem !important;
        font-weight: 800 !important;
        margin: 0 !important;
        padding: 0 !important;
        border: none !important;
        position: relative;
        z-index: 1;
        text-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    .cat-banner p {
        color: rgba(255,255,255,0.85) !important;
        font-size: 1rem !important;
        margin-top: 0.5rem !important;
        position: relative;
        z-index: 1;
        font-weight: 500;
    }

    /* ── Paw divider ── */
    .paw-divider {
        text-align: center;
        font-size: 1.2rem;
        letter-spacing: 12px;
        color: #C8B8A8;
        margin: 0.8rem 0;
        user-select: none;
    }

    /* ── Section cards ── */
    .section-card {
        background: #FFFFFF;
        border: 1.5px solid #EDE8E1;
        border-radius: 20px;
        padding: 1.5rem 1.8rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.04);
        transition: all 0.3s ease;
    }
    .section-card:hover {
        box-shadow: 0 6px 24px rgba(74,88,153,0.1);
        border-color: #C8C0D8;
    }
    .section-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #4A5899;
        margin-bottom: 0.8rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .section-title .icon {
        font-size: 1.3rem;
    }

    /* ── Streamlit text inputs ── */
    .stTextInput label {
        color: #5C5C5C !important;
        font-weight: 600 !important;
        font-size: 0.88rem !important;
    }
    .stTextInput > div > div > input {
        border-radius: 14px !important;
        border: 1.5px solid #E0DAD2 !important;
        background: #FAFAF8 !important;
        color: #3D3D3D !important;
        padding: 0.65rem 1rem !important;
        font-size: 0.95rem !important;
        transition: all 0.25s ease !important;
    }
    .stTextInput > div > div > input:focus {
        border-color: #4A5899 !important;
        box-shadow: 0 0 0 3px rgba(74,88,153,0.1) !important;
        background: #FFFFFF !important;
    }
    .stTextInput > div > div > input::placeholder {
        color: #B8B0A8 !important;
    }

    /* ── File uploader ── */
    [data-testid="stFileUploader"] {
        background: linear-gradient(135deg, #FAFAF8 0%, #F5F0EB 100%) !important;
        border: 2px dashed #D0C8BE !important;
        border-radius: 18px !important;
        padding: 1.5rem !important;
        transition: all 0.3s ease !important;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #4A5899 !important;
        background: linear-gradient(135deg, #FFFFFF 0%, #F8F4EF 100%) !important;
        box-shadow: 0 4px 16px rgba(74,88,153,0.08) !important;
    }
    [data-testid="stFileUploader"] label {
        color: #5C5C5C !important;
        font-weight: 600 !important;
    }
    [data-testid="stFileUploader"] small {
        color: #9B9490 !important;
    }

    /* ── Radio buttons ── */
    .stRadio > label {
        color: #5C5C5C !important;
        font-weight: 600 !important;
    }
    .stRadio [role="radiogroup"] label {
        color: #5C5C5C !important;
        font-weight: 500 !important;
    }
    .stRadio [role="radiogroup"] label:hover {
        color: #4A5899 !important;
    }

    /* ── Checkboxes ── */
    .stCheckbox label span {
        color: #5C5C5C !important;
        font-weight: 500 !important;
    }

    /* ── Primary CTA button (extract) ── */
    div[data-testid="stButton"] > button {
        background: linear-gradient(135deg, #4A5899 0%, #6B7DB8 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 100px !important;
        padding: 0.85rem 2rem !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        letter-spacing: 0.3px;
        box-shadow: 0 6px 20px rgba(74,88,153,0.3) !important;
        transition: all 0.3s cubic-bezier(0.4,0,0.2,1) !important;
    }
    div[data-testid="stButton"] > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 10px 28px rgba(74,88,153,0.35) !important;
        background: linear-gradient(135deg, #5A68A9 0%, #7B8DC8 100%) !important;
    }
    div[data-testid="stButton"] > button:active {
        transform: translateY(0) !important;
    }

    /* ── Download button ── */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #5DA87E 0%, #7EC89E 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 100px !important;
        padding: 0.85rem 2rem !important;
        font-weight: 700 !important;
        box-shadow: 0 6px 20px rgba(93,168,126,0.3) !important;
        transition: all 0.3s ease !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 10px 28px rgba(93,168,126,0.35) !important;
    }

    /* ── Success message ── */
    [data-testid="stAlert"][data-baseweb*="notification"] {
        border-radius: 16px !important;
    }

    /* ── Subheader in sections ── */
    h3 {
        color: #4A5899 !important;
        font-weight: 700 !important;
        font-size: 1.1rem !important;
        border: none !important;
        padding: 0 !important;
        margin-top: 0.5rem !important;
    }

    /* ── Custom scrollbar ── */
    ::-webkit-scrollbar { width: 8px; }
    ::-webkit-scrollbar-track { background: #FDF8F3; }
    ::-webkit-scrollbar-thumb { background: #D8D0C8; border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: #C0B8B0; }

    /* ── Cat badge ── */
    .cat-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(74,88,153,0.08);
        color: #4A5899;
        font-weight: 600;
        font-size: 0.82rem;
        padding: 4px 14px;
        border-radius: 100px;
        margin-bottom: 0.8rem;
    }

    /* ── Footer ── */
    .cat-footer {
        text-align: center;
        padding: 1.5rem 0 1rem;
        color: #B8B0A8;
        font-size: 0.85rem;
        font-weight: 500;
    }
    .cat-footer a {
        color: #4A5899;
        text-decoration: none;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ─── Banner ───
st.markdown("""
<div class="cat-banner">
    <h2>🐾 AI Cat Assistant</h2>
    <p>Upload a contract PDF and let me extract rates, dates & terms into Excel for you~ ✨</p>
</div>
""", unsafe_allow_html=True)

# ─── Paw divider helper ───
def paw_divider():
    st.markdown('<div class="paw-divider">🐾🐾🐾</div>', unsafe_allow_html=True)

# ─── Section header helper ───
def section_start(icon, title):
    st.markdown(f"""
    <div style="
        display:flex; align-items:center; gap:10px;
        margin-bottom:0.6rem; margin-top:0.3rem;
    ">
        <span style="
            background:rgba(74,88,153,0.1);
            border-radius:12px;
            padding:6px 10px;
            font-size:1.2rem;
        ">{icon}</span>
        <span style="
            font-weight:700;
            color:#4A5899;
            font-size:1.05rem;
        ">{title}</span>
    </div>
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════
#  1 · API KEY
# ═══════════════════════════════════════
section_start("🔑", "API Key")
api_key = st.text_input(
    "Gemini API Key",
    type="password",
    placeholder="Paste your Gemini API key here…",
    help="Used client-side only. Never stored.",
    label_visibility="collapsed"
)

paw_divider()

# ═══════════════════════════════════════
#  2 · PDF UPLOAD
# ═══════════════════════════════════════
section_start("📄", "Upload Contract PDF")
uploaded_file = st.file_uploader(
    "Upload PDF",
    type=["pdf"],
    label_visibility="collapsed"
)

paw_divider()

# ═══════════════════════════════════════
#  3 · PROPERTY INFO
# ═══════════════════════════════════════
section_start("🏨", "Property Information")
prop_type = st.radio("Property Type", ["Hotel / Resort", "Cruise Ship"], horizontal=True)
col1, col2 = st.columns(2)
with col1:
    hotel_id = st.text_input("Hotel / Cruise ID", placeholder="e.g. 12711588")
with col2:
    supplier = st.text_input("Supplier Name", placeholder="e.g. Agoda, Expedia")

paw_divider()

# ═══════════════════════════════════════
#  4 · ROOMS
# ═══════════════════════════════════════
section_start("🚪", "Rooms Configuration")

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
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.button("🗑️", key=f"del_{idx}"):
            remove_room(idx)
            st.rerun()

st.button("＋ Add Room", on_click=add_room, use_container_width=True)

paw_divider()

# ═══════════════════════════════════════
#  5 · CONTRACT TYPES
# ═══════════════════════════════════════
section_start("📝", "Contract Types")

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

# ═══════════════════════════════════════
#  6 · EXTRACT BUTTON
# ═══════════════════════════════════════
valid_rooms = [r for r in st.session_state.rooms if r["room_name"].strip() and r["room_id"].strip()]

st.markdown("<br>", unsafe_allow_html=True)
if st.button("🐾  Let the AI Cat Extract Data!", use_container_width=True):
    if not api_key:
        st.error("⚠️ กรุณาใส่ Gemini API Key")
    elif not uploaded_file:
        st.error("⚠️ กรุณาอัพโหลดไฟล์ PDF")
    elif not hotel_id:
        st.error("⚠️ กรุณาใส่ Hotel / Cruise ID")
    elif not valid_rooms:
        st.error("⚠️ กรุณาเพิ่มอย่างน้อย 1 ห้อง (Room Name & Room ID)")
    elif not cts:
        st.error("⚠️ กรุณาเลือกอย่างน้อย 1 contract type")
    else:
        with st.spinner("🐱 น้องแมวกำลังอ่าน contract อยู่นะ… รอสักครู่ (ไม่เกิน 60 วินาที)"):
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

# ─── Footer ───
st.markdown("""
<div class="cat-footer">
    <div style="font-size:1.5rem; margin-bottom:0.3rem;">🐾</div>
    Made with ♡ by <strong>AI Cat Assistant</strong><br>
    <span style="font-size:0.8rem;">Powered by Gemini 2.5 Flash</span>
</div>
""", unsafe_allow_html=True)
