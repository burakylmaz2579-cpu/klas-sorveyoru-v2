import streamlit as st
from google import genai
from google.genai import types
import json
import time
import os
import tempfile
import re
import pandas as pd
from io import BytesIO

# --- SAYFA AYARLARI ---
st.set_page_config(
    page_title="Klas Sörveyörü V2.2",
    page_icon="🚢",
    layout="wide"
)

# --- CSS (Kategori Kartları İçin) ---
st.markdown("""
<style>
    .cat-header { font-size: 1.4rem; font-weight: 700; color: #1e293b; margin-top: 1.5rem; border-bottom: 2px solid #3b82f6; padding-bottom: 0.5rem; }
    .finding-card { border-radius: 12px; padding: 1.2rem; margin-bottom: 1rem; border-left: 6px solid #e2e8f0; border: 1px solid #e2e8f0; background: #ffffff; }
</style>
""", unsafe_allow_html=True)

# --- ROBUST JSON PARSER ---
def robust_json_parser(clean_text):
    try:
        json_match = re.search(r"\{.*\}", clean_text, re.DOTALL)
        if json_match: return json.loads(json_match.group(0))
        return json.loads(clean_text)
    except: return None

# --- API AYARI ---
try:
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
except:
    st.error("API Anahtarı bulunamadı!")
    st.stop()

# --- ARAYÜZ ---
st.title("🚢 Klas Sörveyörü V2.2")
col1, col2 = st.columns([1.2, 1])

# Değişkenleri en başta None tanımlıyoruz (NameError engeller)
parsed_data = None

with col1:
    vessel_name = st.text_input("Gemi Adı")
    imo_number = st.text_input("IMO Numarası")
    grt_dwt = st.text_input("GT / DWT")
    vessel_type = st.selectbox("Gemi Türü", ["General Cargo", "Bulk Carrier", "Oil Tanker", "Chemical Tanker", "Container", "Diğer"])
    # 1. GÜNCELLEME: Sörvey Tipi Seçimi
    survey_type = st.radio("Denetim Tipi (Belgenin başındaki işaret ile eşleşmeli)", ["Annual", "Intermediate", "Renewal", "Initial"])

with col2:
    uploaded_files = st.file_uploader("Belgeleri Yükle", type=["pdf"], accept_multiple_files=True)
    selected_model = st.selectbox("Model Seçimi", ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"])

analyze_btn = st.button("🚀 Belgeleri Oku ve Analiz Et", type="primary", use_container_width=True)

# --- PROMPT (Sörvey Tipi ve Kategori Zorunluluğu) ---
system_instruction = f"""
Sen bir Baş Klas Sörveyörüsün. Şu an '{survey_type}' denetimi yapıyorsun.
1. Belgenin başında işaretlenmiş survey tipi ile '{survey_type}' uyuşmuyorsa "Kırmızı Alarm" ver.
2. Bulguları mutlaka şu kategorilerden birine ata: 'Dokümantasyon', 'Yapısal', 'Makine', 'Emniyet', 'Çevre'.
3. Çıktıda 'category' alanı ZORUNLUDUR.
4. Her madde için ilgili kural referansı (SOLAS/MARPOL vb.) ver.
"""

if analyze_btn and uploaded_files:
    uploaded_gemini_files = []
    tmp_file_paths = []
    
    try:
        with st.spinner("Dosyalar yükleniyor..."):
            for f in uploaded_files:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(f.getvalue())
                    tmp_file_paths.append(tmp.name)
                file_obj = client.files.upload(file=tmp.name)
                uploaded_gemini_files.append(file_obj)
            time.sleep(2)

        prompt_text = f"Analiz edilecek bilgiler: Gemi: {vessel_name}, IMO: {imo_number}, GT/DWT: {grt_dwt}, Tip: {vessel_type}. Survey Tipi: {survey_type}."
        
        with st.spinner("🧠 Analiz yapılıyor..."):
            response = client.models.generate_content(
                model=selected_model,
                contents=uploaded_gemini_files + [prompt_text],
                config=types.GenerateContentConfig(response_mime_type="application/json", system_instruction=system_instruction)
            )
            parsed_data = robust_json_parser(response.text)
    except Exception as e:
        st.error(f"Hata: {str(e)}")

# --- SONUÇ GÖSTERİMİ ---
if parsed_data and "findings" in parsed_data:
    findings = parsed_data["findings"]
    df = pd.DataFrame(findings)
    
    # 2. GÜNCELLEME: Kategori Bazlı Görünüm
    for cat in df['category'].unique():
        st.markdown(f'<div class="cat-header">{cat}</div>', unsafe_allow_html=True)
        for _, row in df[df['category'] == cat].iterrows():
            severity_class = f"card-{row.get('severity', 'info')}"
            st.markdown(f'''
            <div class="finding-card {severity_class}">
                <b>{row["title"]}</b> ({row.get("status", "N/A")})<br>
                <small>Kural: {row.get("rule", "Belirtilmemiş")}</small><br>
                {row["description"]}
            </div>
            ''', unsafe_allow_html=True)
    
    # 3. GÜNCELLEME: Excel'in Sıfırlanmaması İçin
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sörvey Raporu')
    
    st.download_button(
        label="📥 Excel Olarak İndir",
        data=output.getvalue(),
        file_name=f"{vessel_name}_Rapor.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Temizlik
for p in tmp_file_paths:
    if os.path.exists(p): os.remove(p)
