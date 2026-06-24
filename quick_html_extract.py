import streamlit as st
import base64
import requests
import time
import json

def render_ui(api_key):
    st.markdown("""
<div class="step-card">
    <div class="step-title"><span class="step-num">⚡</span> ดึงข้อมูลเฉพาะ HTML (Quick Extract)</div>
    <p style="color: #475569; font-size: 0.9rem;">
        อัปโหลดไฟล์ PDF สัญญาหรือโปรโมชั่น แล้ว AI จะดึงเฉพาะ <strong>Child Policy, Cancellation, และ Meal & Info</strong> ออกมาเป็นโค้ด HTML ให้คุณก๊อปปี้ไปใช้ได้ทันที โดยไม่ต้องตั้งค่าใดๆ
    </p>
</div>
""", unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "ลากหรือเลือก PDF Contract ที่นี่",
        type=["pdf"],
        key="quick_html_uploader"
    )

    if not api_key:
        st.warning("⚠️ กรุณาใส่ API Key ที่แถบด้านซ้ายก่อนใช้งาน")
        return

    if st.button("🚀 สกัดโค้ด HTML", type="primary", use_container_width=True, disabled=not uploaded_file):
        
        loading_placeholder = st.empty()
        loading_placeholder.markdown("""
            <div id="full-screen-loader" style="
                position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
                background-color: rgba(255, 255, 255, 0.85); backdrop-filter: blur(5px);
                z-index: 9999; display: flex; flex-direction: column; justify-content: center; align-items: center;
            ">
                <div style="font-size: 4rem; animation: float 2s ease-in-out infinite;">😸</div>
                <h2 style="color: #334155; margin-top: 1rem;">กำลังอ่านข้อมูล...</h2>
                <p style="color: #64748B;">MeowAI กำลังวิเคราะห์และจัดหน้า HTML ให้ค่ะ โปรดรอสักครู่ 🐾</p>
            </div>
        """, unsafe_allow_html=True)

        try:
            b64_data = base64.b64encode(uploaded_file.read()).decode("utf-8")
            
            prompt_text = """Extract the following information from the provided hotel contract or promotion PDF.
Format the output strictly as a valid JSON object matching the requested schema.
Do NOT use markdown code blocks (e.g. ```json). Output raw JSON.

REQUIREMENTS:
1. Extract the child policy, cancellation policy, and meals/info.
2. Format them EXACTLY according to the HTML templates below.
3. Replace bracketed placeholders like [Age], [Price], [DATES] with actual data from the PDF.

HTML FORMATTING PATTERNS (STRICT):
- child_policy: (Do NOT include any food/meal-related information here)
<p><span style="color: #008000;"><strong>Maximum Occupancy: [Occ]</strong></span></p>
<p>Child [Age] years old Sharing bed + ABF = [Price/FOC] [Currency]</p>
<p>Child/Adult Extra bed + ABF = [Price] [Currency]</p>
<p><span style="color: #ff0000;"><strong>*Cannot add an extra bed</strong></span></p>

- cancellation_policy: 
<p><strong>Cancellation: [Season/Condition]</strong></p>
<p>• Cancellation up to [X] days prior to arrival date, No charge.</p>
<p><strong>No Show & Early Check-Out:</strong></p>
<p>• The equivalent of the full originally booked length of stay will be charged.</p>

- meals_and_info:
If the document is a Main Contract:
<p><strong>MAIN CONTRACT [YEAR] : [DATE] - [DATE]</strong></p>
<p><strong>※ MEAL PLAN</strong></p>
<p>• [Details...]</p>
<p><strong>※ MINIMUM NIGHTS & BLACKOUT DATES</strong></p>
<p>• Minimum [X] Nights stay required on [DATES]</p>
<p><span style="color: #008000;"><strong>COMPULSORY</strong></span> GALA DINNER [Details]</p>
<p><strong>※ SUPPLEMENT CHARGE</strong> [Details]</p>
<p><strong>※ EARLY BIRD & SPECIAL OFFERS</strong></p>
<p>• [Details of Early bird...]</p>
<p><span style="color: #ff0000;"><strong>Remark:</strong></span> [Food space/location info]</p>

If the document is a Promotion or Early Bird offer:
<p><strong>PROMOTION : [PROMOTION NAME]</strong></p>
<p><strong>Promo code : <span style="color: #0000ff;">[CODE]</span></strong></p>
<p><strong>Stay</strong> : [DATE] - [DATE]</p>
<p><strong>Book by</strong> : [DATE]</p>
<p><span style="color: #ff0000;"><strong>*Black Out : [DATES]</strong></span></p>
<p> </p>
<p><strong>⁜ Supplement charge (update [DATE])</strong></p>
<p>• Supplement charge of weekend <span style="color: #800080;"><strong>([DAYS])</strong></span> long weekend = [PRICE] THB per night at <strong>[ROOM TYPE]</strong></p>
<p> </p>
<p><strong>Terms and Conditions</strong></p>
<p>• Commission: [INFO]</p>
<p>• Breakfast: [INFO]</p>
<hr />
<p> <span style="color: #000000;"><strong>※ FAMILY BENEFIT</strong></span></p>
<p><span style="color: #000000;">Child [AGE] years old Sharing bed + ABF = [PRICE]</span><br /></p>
<p><span style="color: #ff0000;">*maximum one sofa bed per room</span></p>

OUTPUT SCHEMA (JSON):
{
  "child_policy": "<p>...</p>",
  "cancellation_policy": "<p>...</p>",
  "meals_and_info": "<p>...</p>"
}"""
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

            models_to_try = ["gemini-2.5-flash"]
            max_retries = 4
            resp_data = None
            last_error = None

            for model_name in models_to_try:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
                for attempt in range(max_retries):
                    try:
                        resp = requests.post(url, json=payload, timeout=300)
                        if resp.status_code == 200:
                            resp_data = resp.json()
                            break
                        error_msg = resp.json().get("error", {}).get("message", f"HTTP {resp.status_code}")
                        last_error = error_msg
                        if "high demand" in error_msg.lower() or resp.status_code in (429, 500, 503):
                            if attempt < max_retries - 1:
                                time.sleep(15 * (2 ** attempt))
                                continue
                        break
                    except requests.exceptions.Timeout:
                        last_error = "Request timed out"
                        if attempt < max_retries - 1:
                            time.sleep(15 * (2 ** attempt))
                            continue
                    except Exception as e:
                        last_error = str(e)
                        break
                
                if resp_data and resp.status_code == 200:
                    break

            loading_placeholder.markdown("""
                <style>#full-screen-loader { display: none !important; }</style>
            """, unsafe_allow_html=True)

            if not resp_data or resp.status_code != 200:
                st.error(f"เกิดข้อผิดพลาด: {last_error}")
                return
            
            text_output = resp_data["candidates"][0]["content"]["parts"][0]["text"]
            parsed_data = json.loads(text_output)

            st.success("✅ ดึงข้อมูลสำเร็จ! คุณสามารถกดปุ่ม Copy มุมขวาบนของแต่ละกล่องเพื่อคัดลอกโค้ดได้เลยค่ะ")
            
            st.markdown("### 👶 Child Policy")
            st.code(parsed_data.get("child_policy", ""), language="html")

            st.markdown("### ❌ Cancellation Policy")
            st.code(parsed_data.get("cancellation_policy", ""), language="html")

            st.markdown("### 🍽️ Meals & Info")
            st.code(parsed_data.get("meals_and_info", ""), language="html")

        except Exception as e:
            loading_placeholder.markdown("""
                <style>#full-screen-loader { display: none !important; }</style>
            """, unsafe_allow_html=True)
            st.error(f"เกิดข้อผิดพลาดในระบบ: {str(e)}")
