import streamlit as st
import json
import base64
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io
import os
from contract_to_excel import generate_rows, HEADERS, COL_WIDTHS

# ─── Page Config ───
st.set_page_config(
    page_title="MeowAI 🐾 Contract Extractor",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── Mascot Setup (Updated) ───
try:
    from mascot_data import MASCOT_B64
except ImportError:
    MASCOT_B64 = ""

if MASCOT_B64:
    mascot_img_html = f'<div class="mascot-container"><img src="data:image/png;base64,{MASCOT_B64}" class="mascot-avatar" /></div>'
else:
    mascot_img_html = '<div class="mascot-container" style="display:flex;align-items:center;justify-content:center;font-size:80px;">🐱</div>'

# ─── CSS ───
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif !important;
    }

    /* ── Background ── */
    .stApp {
        background: linear-gradient(135deg, #F8FAFC 0%, #F1F5F9 50%, #E2E8F0 100%) !important;
        min-height: 100vh;
    }

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1E293B 0%, #0F172A 100%) !important;
        border-right: 1px solid rgba(255,255,255,0.05) !important;
    }
    [data-testid="stSidebar"] * { color: #CBD5E1 !important; }
    [data-testid="stSidebar"] h4 { color: #94A3B8 !important; font-size: 0.75rem !important; text-transform: uppercase; letter-spacing: 1.5px; margin-top: 1rem !important; }
    [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.1) !important; }
    [data-testid="stSidebar"] .stTextInput > div > div > input {
        background: rgba(255,255,255,0.07) !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        color: #E2E8F0 !important;
        border-radius: 8px !important;
    }
    [data-testid="stSidebar"] .stRadio label span { color: #CBD5E1 !important; }

    /* ── Mascot ── */
    .mascot-container {
        display: flex;
        justify-content: center;
        align-items: center;
        padding: 1.2rem 0 0.5rem 0;
    }
    .mascot-avatar {
        width: 120px;
        height: 120px;
        object-fit: contain;
        mix-blend-mode: multiply;
        filter: drop-shadow(0 4px 15px rgba(0,0,0,0.15));
    }
    .sidebar-title {
        text-align: center;
        font-size: 1.4rem !important;
        font-weight: 800 !important;
        color: #F1F5F9 !important;
        margin: 0.3rem 0 0.1rem 0 !important;
    }
    .sidebar-subtitle {
        text-align: center;
        font-size: 0.78rem !important;
        color: #64748B !important;
        margin: 0 !important;
        letter-spacing: 1px;
        text-transform: uppercase;
    }
    .sidebar-pill {
        background: linear-gradient(135deg, #334155, #475569);
        border-radius: 10px;
        padding: 0.5rem 1rem;
        text-align: center;
        font-size: 0.78rem;
        color: #E2E8F0 !important;
        margin: 0.8rem 0;
        border: 1px solid rgba(255,255,255,0.1);
    }

    /* ── Chat Bubble ── */
    .chat-bubble {
        background: #FFFFFF;
        border-radius: 24px;
        border-bottom-left-radius: 4px;
        padding: 1.8rem 2rem;
        box-shadow: 0 10px 30px rgba(100,116,139,0.1);
        position: relative;
        margin-bottom: 1.5rem;
        border: 2px solid #E2E8F0;
        display: flex;
        align-items: center;
        gap: 1.5rem;
    }
    .chat-bubble::before {
        content: '';
        position: absolute;
        bottom: -2px; left: -20px;
        border-width: 20px 20px 0 0;
        border-style: solid;
        border-color: #E2E8F0 transparent transparent transparent;
    }
    .chat-bubble::after {
        content: '';
        position: absolute;
        bottom: 0; left: -16px;
        border-width: 17px 17px 0 0;
        border-style: solid;
        border-color: #FFFFFF transparent transparent transparent;
    }
    .chat-bubble h3 { color: #1E293B !important; margin: 0 0 0.4rem 0 !important; font-size: 1.5rem !important; font-weight: 800 !important; }
    .chat-bubble p  { color: #475569 !important; margin: 0 !important; font-size: 1rem !important; line-height: 1.6 !important; }

    /* ── Output Badge ── */
    .output-badge {
        background: linear-gradient(135deg, #1E293B 0%, #334155 100%);
        border-radius: 16px;
        padding: 1rem 1.5rem;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        gap: 1.2rem;
        color: #fff;
        box-shadow: 0 8px 25px rgba(30,41,59,0.25);
    }
    .output-badge h4 { margin: 0; font-size: 1rem; font-weight: 700; color: #fff; }
    .output-badge p  { margin: 0; font-size: 0.8rem; opacity: 0.7; color: #fff; }
    .output-stat {
        background: rgba(255,255,255,0.12);
        border-radius: 10px;
        padding: 8px 16px;
        font-size: 1.5rem;
        font-weight: 800;
        text-align: center;
        min-width: 70px;
        flex-shrink: 0;
    }
    .output-stat span { display: block; font-size: 0.62rem; font-weight: 500; opacity: 0.75; margin-top: 2px; }

    /* ── Step Cards ── */
    .step-card {
        background: #FFFFFF;
        border-radius: 16px;
        border: 2px solid #E2E8F0;
        padding: 1.2rem 1.4rem 0.8rem 1.4rem;
        margin-bottom: 0.8rem;
        box-shadow: 0 4px 12px rgba(100,116,139,0.07);
        transition: box-shadow 0.25s;
    }
    .step-card:hover { box-shadow: 0 8px 24px rgba(100,116,139,0.13); }
    .step-title {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1.8px;
        color: #64748B;
        margin-bottom: 0.9rem;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .step-num {
        background: linear-gradient(135deg, #64748B, #475569);
        color: #fff;
        width: 22px; height: 22px;
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 0.72rem;
        font-weight: 800;
        flex-shrink: 0;
    }

    /* ── Room Table Header ── */
    .room-header {
        display: grid;
        grid-template-columns: 1fr 100px 32px;
        gap: 8px;
        padding: 4px 4px 6px 4px;
        font-size: 0.68rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #94A3B8;
        border-bottom: 2px solid #E2E8F0;
        margin-bottom: 2px;
    }

    /* ── Summary Badges ── */
    .ct-row {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-bottom: 0.8rem;
    }
    .ct-badge {
        background: #F1F5F9;
        border: 1.5px solid #CBD5E1;
        border-radius: 20px;
        padding: 4px 14px;
        font-size: 0.78rem;
        font-weight: 600;
        color: #475569;
    }
    .ct-badge-active {
        background: #1E293B;
        border-color: #1E293B;
        color: #fff;
    }

    /* ── Inputs ── */
    .stTextInput > div > div > input {
        border-radius: 10px !important;
        border: 2px solid #E2E8F0 !important;
        padding: 0.6rem 1rem !important;
        background-color: #F8FAFC !important;
        color: #334155 !important;
        font-weight: 500;
        transition: all 0.25s;
    }
    .stTextInput > div > div > input:focus {
        border-color: #94A3B8 !important;
        box-shadow: 0 0 0 3px rgba(148,163,184,0.15) !important;
        background-color: #FFFFFF !important;
    }

    /* ── Checkbox ── */
    [data-testid="stCheckbox"] label span { color: #475569 !important; font-weight: 500 !important; }
    [data-testid="stCheckbox"] label { white-space: nowrap !important; }

    /* ── Main Button ── */
    div[data-testid="stButton"] > button {
        background: linear-gradient(135deg, #1E293B 0%, #334155 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 14px !important;
        padding: 0.85rem 2rem !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        width: 100%;
        box-shadow: 0 6px 20px rgba(30,41,59,0.3) !important;
        transition: all 0.3s cubic-bezier(0.4,0,0.2,1) !important;
        letter-spacing: 0.3px;
    }
    div[data-testid="stButton"] > button:hover {
        transform: translateY(-2px) scale(1.01) !important;
        box-shadow: 0 12px 28px rgba(30,41,59,0.4) !important;
    }

    /* ── Download Button ── */
    div[data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #16A34A 0%, #15803D 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 14px !important;
        padding: 0.85rem 2rem !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        width: 100%;
        box-shadow: 0 6px 20px rgba(22,163,74,0.3) !important;
        transition: all 0.3s !important;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 12px 28px rgba(22,163,74,0.4) !important;
    }

    /* ── Float animation ── */
    @keyframes float {
        0%   { transform: translateY(0px); }
        50%  { transform: translateY(-8px); }
        100% { transform: translateY(0px); }
    }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════
with st.sidebar:
    st.markdown(
        f"""<div class="mascot-container">{mascot_img_html}</div>
<h2 class="sidebar-title">MeowAI 🐾</h2>
<p class="sidebar-subtitle">Contract Extractor</p>""",
        unsafe_allow_html=True
    )

    st.markdown("""
<div class="sidebar-pill">
    📊 Output: <strong>43 Columns</strong> · Dashboard Ready
</div>
""", unsafe_allow_html=True)

    st.markdown("#### 🔑 Gemini API Key")
    api_key = st.text_input("API Key", type="password", placeholder="AIza...", label_visibility="collapsed")

    st.markdown("---")
    st.markdown("#### 🏨 Property Info")
    prop_type = st.radio("Property Type", ["Hotel / Resort", "Cruise Ship"])
    hotel_id  = st.text_input("Property ID", placeholder="e.g. 12711588")
    supplier  = st.text_input("Hotel Name", placeholder="e.g. InterContinental Khao Yai")

    st.markdown("---")
    st.markdown("""
<div style="text-align:center;color:#475569;font-size:0.72rem;padding:0.3rem 0;">
    🐾 Powered by Gemini 2.5 Flash<br>
    เหมียวดึงข้อมูลให้ครบ ไม่มีตกหล่น
</div>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════
#  HERO
# ═══════════════════════════════════════════
st.markdown("""
<div class="chat-bubble">
    <div style="flex-shrink:0;font-size:3.2rem;animation:float 3s ease-in-out infinite;">😸</div>
    <div>
        <h3>สวัสดี DATA TEAM! 🐾</h3>
        <p>โยน PDF สัญญาโรงแรมหรือ Cruise มาเลย! หนู MeowAI จะสกัดข้อมูล Rate, Date, Surcharge, EB, Promo ทุก field ให้ครบ <strong>43 คอลัมน์</strong> ตามฟอร์แมต Dashboard พร้อม Upload ได้ทันที ไม่มีตกหล่น!</p>
    </div>
</div>
""", unsafe_allow_html=True)



# ═══════════════════════════════════════════
#  MAIN COLUMNS
# ═══════════════════════════════════════════
col_left, col_right = st.columns([1.1, 1], gap="large")

with col_left:
    # ── Step 1: Upload ──────────────────────
    st.markdown("""
<div class="step-card">
    <div class="step-title"><span class="step-num">1</span> อัพโหลดไฟล์สัญญา (PDF)</div>
</div>
""", unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "ลากหรือเลือก PDF Contract",
        type=["pdf"],
        help="รองรับ PDF ภาษาอังกฤษ — Hotel, Resort และ Cruise"
    )

    # ── Step 2: Contract Types ──────────────
    st.markdown("""
<div class="step-card" style="margin-top:0.6rem;">
    <div class="step-title"><span class="step-num">2</span> ประเภทสัญญาที่จะสร้าง</div>
</div>
""", unsafe_allow_html=True)

    ct_main = st.checkbox("📋 Main Contract", value=True,
                          help="Base net rate — สร้างทุก period ใน PDF")

    ct_eb = st.checkbox("🦅 Early Bird", value=False,
                        help="AI จะ detect tier, days, discount% จาก PDF อัตโนมัติ — promo_code = E.B X DAYS, min_advance_days = X")

    c3, c4, c5 = st.columns([1.8, 1.1, 1.1])
    with c3:
        ct_promo = st.checkbox("🎯 Promotion", value=False,
                               help="Early Bird Offer หรือโปรโมชั่นพิเศษ")
    with c4:
        promo_code = st.text_input(
            "Promo Code", placeholder="EBO",
            disabled=not ct_promo, label_visibility="collapsed"
        )
    with c5:
        promo_till = st.text_input(
            "Book Till", placeholder="YYYY-MM-DD",
            disabled=not ct_promo, label_visibility="collapsed"
        )

    ct_por = st.checkbox("💎 POR (Price on Request)", value=False,
                         help="net_price=0, contract_type=POR")

    if ct_promo and promo_till:
        st.caption(f"📅 promo_book_till จะเป็น: `{promo_till} 23:59:59`")

with col_right:
    # ── Step 3: Rooms ───────────────────────
    st.markdown("""
<div class="step-card">
    <div class="step-title"><span class="step-num">3</span> ตั้งค่าห้องพัก / Cabin</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("""
<div class="room-header">
    <span>Room Name</span>
    <span>Room ID</span>
    <span></span>
</div>
""", unsafe_allow_html=True)

    if "rooms" not in st.session_state:
        st.session_state.rooms = [{"room_name": "", "room_id": ""}]

    def add_room():
        st.session_state.rooms.append({"room_name": "", "room_id": ""})

    def remove_room(i):
        if len(st.session_state.rooms) > 1:
            st.session_state.rooms.pop(i)

    for idx, room in enumerate(list(st.session_state.rooms)):
        rc1, rc2, rc3 = st.columns([5, 3.5, 1])
        with rc1:
            room["room_name"] = st.text_input(
                f"room_name_{idx}", value=room["room_name"],
                key=f"rn_{idx}", placeholder="Deluxe Pool Villa",
                label_visibility="collapsed"
            )
        with rc2:
            room["room_id"] = st.text_input(
                f"room_id_{idx}", value=room["room_id"],
                key=f"ri_{idx}", placeholder="12711588001",
                label_visibility="collapsed"
            )
        with rc3:
            if st.button("✕", key=f"del_{idx}", help="ลบห้องนี้"):
                remove_room(idx)
                st.rerun()

    st.button("＋ เพิ่มห้องพัก", on_click=add_room, use_container_width=True)

    valid_count = sum(1 for r in st.session_state.rooms if r["room_name"].strip() and r["room_id"].strip())
    total_count = len(st.session_state.rooms)
    if valid_count:
        st.caption(f"✅ {valid_count}/{total_count} ห้องพร้อมใช้งาน")
    else:
        st.caption(f"⚠️ กรุณากรอก Room Name และ Room ID ให้ครบ")


# ═══════════════════════════════════════════
#  PROCESS LOGIC
# ═══════════════════════════════════════════
st.markdown("<br>", unsafe_allow_html=True)

valid_rooms = [r for r in st.session_state.rooms if r["room_name"].strip() and r["room_id"].strip()]

cts = []
if ct_main:  cts.append({"type": "main",  "label": "📋 Main Contract"})
if ct_eb:    cts.append({"type": "eb",    "label": "🦅 Early Bird"})
if ct_promo: cts.append({"type": "promo", "label": "🎯 Promotion",  "promo_code": promo_code, "promo_till": promo_till})
if ct_por:   cts.append({"type": "por",   "label": "💎 POR"})

# Summary badges
badge_html = ""
for c in cts:
    badge_html += f'<span class="ct-badge ct-badge-active">{c["label"]}</span>'
if valid_rooms:
    badge_html += f'<span class="ct-badge">🚪 {len(valid_rooms)} room(s)</span>'
if badge_html:
    st.markdown(f'<div class="ct-row">{badge_html}</div>', unsafe_allow_html=True)

ready = bool(api_key and uploaded_file and hotel_id and valid_rooms and cts)

if not ready:
    missing = []
    if not api_key:       missing.append("🔑 Gemini API Key")
    if not uploaded_file: missing.append("📄 ไฟล์ PDF")
    if not hotel_id:      missing.append("🏨 Property ID")
    if not valid_rooms:   missing.append("🚪 ห้องพัก (≥ 1 ห้อง)")
    if not cts:           missing.append("📋 ประเภทสัญญา")
    st.info("เมี้ยว~ ยังขาดข้อมูล: " + " · ".join(missing))

if st.button("🐾 สั่งเหมียวดึงข้อมูล — Generate Excel (43 columns)", use_container_width=True, key="process_btn"):
    if not ready:
        st.warning("⚠️ กรุณากรอกข้อมูลให้ครบก่อนนะเหมียว!")
    else:
        with st.spinner("🐱 น้องแมวกำลังอ่าน contract อยู่นะ… รอสักครู่ (ไม่เกิน 60 วินาที)"):
            try:
                file_bytes = uploaded_file.read()
                b64_data   = base64.b64encode(file_bytes).decode("utf-8")

                ct_desc_parts = []
                for c in cts:
                    if c["type"] == "main":
                        ct_desc_parts.append("Main Contract: extract all standard net rates per period.")
                    elif c["type"] == "eb":
                        ct_desc_parts.append(
                            'Early Bird: detect ALL tiers from PDF automatically. '
                            'For each tier, extract days_in_advance and discount_pct. '
                            'promo_code = "E.B {days} DAYS" (e.g. "E.B 30 DAYS"), '
                            'min_advance_days = days. Respect blackout dates (eb_blackout=true).'
                        )
                    elif c["type"] == "promo":
                        ct_desc_parts.append(
                            f'Promotion: promo_code="{c["promo_code"]}", '
                            f'promo_book_till="{c["promo_till"]} 23:59:59".'
                        )
                    elif c["type"] == "por":
                        ct_desc_parts.append('POR: net_price=0, promo_code="POR Rate".')
                ct_desc = " | ".join(ct_desc_parts)

                prompt_text = f"""You are an expert hotel contract data extractor. Extract ALL data from this PDF.

ROOMS (use exactly these IDs and names, do not invent new ones): {json.dumps(valid_rooms)}
CONTRACT TYPES TO EXTRACT: {ct_desc}
PROPERTY TYPE: {prop_type}

CRITICAL RULES:
1. Extract ONE period per distinct date range. Split periods if surcharge/blackout/min-nights rules differ.
2. WEEKDAY/WEEKEND RATES:
   - Set has_weekday_weekend=true
   - rates dict = BASE (weekday) rate per room
   - has_surcharge=true
   - surcharge_rates = per-room weekend supplement dict e.g. {{"room_id_1": 1500, "room_id_2": 3000}}
   - If same surcharge for all rooms use surcharge_amount (single value) instead
3. ALL HTML fields must use real HTML: <p>, <strong>, <span style="color:#ff0000;">
4. meals_and_info = full HTML: contract name, validity, surcharges, meal rates, blackout dates, min nights, all key notes.
5. cancellation_policy and child_policy = full HTML extracted from PDF.
6. promo_book_till format: "YYYY-MM-DD 23:59:59" (ONLY if PDF explicitly states a booking deadline)
7. Early Bird: promo_book_till = null UNLESS PDF clearly states a "book by" deadline.
8. net_price = integer only (no decimals, no commas).
9. rates dict key = room_id string EXACTLY as given in ROOMS above.
10. CRUISE: detect night package from PDF (e.g. "1 Night", "2 Nights") and use as promo_code.
11. hotel_id = "{hotel_id}", hotel_supplier = "{supplier}"

Return ONLY valid JSON (no markdown fences, no explanation):
{{
  "hotel_id": "{hotel_id}",
  "hotel_supplier": "{supplier}",
  "abf": "Included",
  "cancellation_policy": "<p><strong>CANCELLATION:</strong></p><p>...</p>",
  "child_policy": "<p><span style=\\"color:#008000;\\"><strong>Maximum Occupancy: 2A+1C</strong></span></p><p>Child 0-6.99 years: FOC</p>",
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

                url = (
                    "https://generativelanguage.googleapis.com/v1beta/"
                    f"models/gemini-2.5-flash:generateContent?key={api_key}"
                )
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

                resp      = requests.post(url, json=payload, headers={"Content-Type": "application/json"}, timeout=120)
                resp_data = resp.json()

                if resp.status_code != 200:
                    raise Exception(resp_data.get("error", {}).get("message", "Gemini API Error"))

                text_output = resp_data["candidates"][0]["content"]["parts"][0]["text"]
                parsed_data = json.loads(text_output)

                # ── Generate rows (43-column dashboard format) ──
                rows = generate_rows(parsed_data, cts)

                # ── Build Excel workbook ──
                wb = Workbook()
                ws = wb.active
                ws.title = "Sheet1"

                header_fill = PatternFill("solid", start_color="1F4E79")
                header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
                data_font   = Font(name="Arial", size=10)

                # Header row — ALL 43 columns, always present
                for c_idx, h in enumerate(HEADERS, 1):
                    cell = ws.cell(1, c_idx, h)
                    cell.fill      = header_fill
                    cell.font      = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    ws.column_dimensions[get_column_letter(c_idx)].width = COL_WIDTHS.get(h, 12)
                ws.row_dimensions[1].height = 20

                date_only_cols = {"start_date", "end_date", "edited_date", "cutoff_date"}
                datetime_cols  = {"created_date"}

                # Data rows — ALL 43 columns, None → empty cell
                for r_idx, row in enumerate(rows, 2):
                    for c_idx, h in enumerate(HEADERS, 1):
                        val  = row.get(h)
                        cell = ws.cell(r_idx, c_idx, val)
                        cell.font = data_font
                        if h in date_only_cols and val:
                            cell.number_format = "DD/MM/YYYY"
                        elif h in datetime_cols and val:
                            cell.number_format = "DD/MM/YYYY HH:MM:SS"
                        if h == "net_price" and val is not None:
                            cell.value = str(val)

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)

                n_periods = len(parsed_data.get("periods", []))
                n_rooms   = len(valid_rooms)
                st.success(
                    f"✅ สำเร็จ! เหมียวดึงข้อมูลได้ **{len(rows)} แถว** "
                    f"({n_periods} periods × {n_rooms} rooms · 43 columns ครบทุก field)"
                )
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
