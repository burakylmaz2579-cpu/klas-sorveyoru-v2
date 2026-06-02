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
st.set_page_config(page_title="Klas Sörveyörü V2.1", layout="wide")

# --- CSS VE TASARIM ---
st.markdown("""
<style>
    .cat-header { font-size: 1.4rem; font-weight: 700; color: #1e293b; margin-top: 1.5rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 0.5rem; }
    .finding-card { border-radius: 12px; padding: 1.2rem; margin-bottom: 1rem; border-left: 6px solid; border: 1px solid #e2e8f0; background: #ffffff; }
</style>
""", unsafe_allow_html=True)

# --- SİSTEM AYARLARI ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    st.error("Secrets (API KEY) ayarı eksik. Lütfen Manage app > Settings > Secrets kısmından ekleyin.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# --- ARAYÜZ VE INPUT ---
st.title("🚢 Klas Sörveyörü V2.1")
col1, col2 = st.columns(2)

with col1:
    vessel_name = st.text_input("Gemi Adı")
    imo_number = st.text_input("IMO Numarası")
    grt_dwt = st.text_input("GT / DWT")
    vessel_type = st.selectbox("Gemi Türü", ["General Cargo", "Bulk Carrier", "Oil Tanker", "Chemical Tanker", "Container", "Diğer"])
    survey_type = st.radio("Denetim Tipi", ["Annual", "Intermediate", "Renewal", "Initial"])

with col2:
    uploaded_files = st.file_uploader("Belgeleri Yükle", type=["pdf"], accept_multiple_files=True)
    selected_model = st.selectbox("Model", ["gemini-1.5-flash", "gemini-2.0-flash"])

# --- ANALİZ MANTIĞI ---
parsed_data = None # Değişkeni burada tanımlıyoruz
analyze_btn = st.button("🚀 Analizi Başlat")

if analyze_btn and uploaded_files:
    # 1. Yükleme ve İşleme
    uploaded_gemini_files = []
    tmp_file_paths = []
    
    with st.spinner("Dosyalar işleniyor..."):
        for uploaded_file in uploaded_files:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_file_paths.append(tmp.name)
            file_obj = client.files.upload(file=tmp.name)
            uploaded_gemini_files.append(file_obj)
        time.sleep(2)

    # 2. Prompt ve API Çağrısı
    system_instruction = f"""Sen bir Baş Klas Sörveyörüsün. '{survey_type}' denetimi yapıyorsun. 
    Belgeleri çapraz kontrol et. JSON dön. Her madde için mutlaka SOLAS/MARPOL/IMO kural referansı ver."""
    
    contents = uploaded_gemini_files + [f"Gemi: {vessel_name}, IMO: {imo_number}, Tür: {vessel_type}"]
    
    response = client.models.generate_content(
        model=selected_model,
        contents=contents,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            system_instruction=system_instruction
        )
    )
    
    # 3. JSON Çıktısı
    clean_res = response.text.replace('```json', '').replace('```', '').strip()
    try:
        parsed_data = json.loads(clean_res)
    except:
        st.error("JSON okuma hatası.")

# --- SONUÇLARI GÖSTERME (parsed_data varsa) ---
if parsed_data and "findings" in parsed_data:
    findings = parsed_data["findings"]
    df = pd.DataFrame(findings)
    
    for cat in df['category'].unique():
        st.markdown(f'<div class="cat-header">{cat}</div>', unsafe_allow_html=True)
        for _, row in df[df['category'] == cat].iterrows():
            st.markdown(f'<div class="finding-card"><b>{row["title"]}</b><br>{row["description"]}</div>', unsafe_allow_html=True)
    
    # Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    st.download_button("📥 Excel Olarak İndir", output.getvalue(), "rapor.xlsx", "application/vnd.ms-excel")
