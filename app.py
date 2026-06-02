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

st.set_page_config(page_title="Klas Sörveyörü V2.1", layout="wide")

# --- CSS (Kategori Kartları İçin) ---
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
    st.error("Secrets ayarı eksik.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# --- ARAYÜZ ---
st.title("🚢 Klas Sörveyörü V2.1 | Gelişmiş Denetim")
col1, col2 = st.columns([1, 1])

with col1:
    vessel_name = st.text_input("Gemi Adı")
    imo_number = st.text_input("IMO Numarası")
    grt_dwt = st.text_input("GT / DWT")
    vessel_type = st.selectbox("Gemi Türü", ["General Cargo", "Bulk Carrier", "Oil Tanker", "Chemical Tanker", "Container", "Diğer"])
    # Survey Tipi Seçimi
    survey_type = st.radio("Denetim Tipi (Belgedeki kutucukları doğrula)", ["Annual", "Intermediate", "Renewal", "Initial"])

with col2:
    uploaded_files = st.file_uploader("Belgeleri Yükle", type=["pdf"], accept_multiple_files=True)

analyze_btn = st.button("🚀 Analizi Başlat", type="primary")

# --- MASTER PROMPT (Survey Tipi Eklendi) ---
system_instruction = f"""
Sen bir Baş Klas Sörveyörüsün. Kullanıcı şu an '{survey_type}' denetimi yapıyor.
SANA GELEN PDF'LERİ İNCELERKEN:
1. Belgenin başında işaretlenmiş olan (Tik veya çarpı atılmış) 'Survey Type' kısmını oku. Eğer kullanıcının seçtiği '{survey_type}' tipi ile belgedeki işaretli tip UYUŞMUYORSA 'Kırmızı Alarm' ver.
2. Bulguları şu kategorilere ayır: [Dokümantasyon, Yapısal, Makine, Emniyet, Çevre/Kirlilik].
3. JSON çıktında 'category' anahtarını zorunlu tut.
4. Excel çıktısı için her bulguyu ayrı satır yap.
"""

if analyze_btn and uploaded_files:
    # (Kodun dosya yükleme ve API çağırma kısmı aynı kalacak)
    # ... [Yükleme ve API çağrısı buraya gelecek] ...
    
    # JSON ayrıştırma ve Excel
    if parsed_data and "findings" in parsed_data:
        findings = parsed_data["findings"]
        
        # Kategori Bazlı Gösterim
        df = pd.DataFrame(findings)
        categories = df['category'].unique()
        
        for cat in categories:
            st.markdown(f'<div class="cat-header">{cat}</div>', unsafe_allow_html=True)
            for f in findings:
                if f['category'] == cat:
                    # [Finding card çizimi buraya]
                    pass
        
        # Excel Butonu (Excel'de her şeyin görünmesi için df.to_excel kullanıldı)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 Excel İndir", output.getvalue(), "rapor.xlsx", "application/vnd.ms-excel")
