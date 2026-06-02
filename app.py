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
    page_title="Klas Sörveyörü V3.0 | Canlı Analiz Motoru",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- PREMIUM TASARIM (CSS) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"], .stApp { font-family: 'Plus Jakarta Sans', sans-serif; }
    .header-container { background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 2.5rem; border-radius: 16px; color: white; margin-bottom: 2rem; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1); }
    .header-title { font-size: 2.2rem; font-weight: 700; margin: 0; background: linear-gradient(to right, #ffffff, #94a3b8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .header-subtitle { color: #94a3b8; font-size: 1.1rem; margin-top: 0.5rem; }
    .cat-header { font-size: 1.4rem; font-weight: 700; color: #0f172a; margin-top: 1.8rem; border-bottom: 3px solid #3b82f6; padding-bottom: 0.4rem; margin-bottom: 1rem; }
    .finding-card { border-radius: 12px; padding: 1.5rem; margin-bottom: 1.2rem; border-left: 6px solid; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); border: 1px solid rgba(0,0,0,0.03); background: #ffffff; }
    .card-critical { border-left-color: #ef4444; background: #fef2f2; }
    .card-error { border-left-color: #f43f5e; background: #fff1f2; }
    .card-warning { border-left-color: #f59e0b; background: #fffbeb; }
    .card-info { border-left-color: #3b82f6; background: #eff6ff; }
    .card-success { border-left-color: #10b981; background: #f0fdf4; }
    .finding-title { font-size: 1.15rem; font-weight: 700; color: #0f172a; margin-bottom: 0.4rem; }
    .finding-rule { display: inline-block; padding: 0.2rem 0.75rem; border-radius: 9999px; font-size: 0.75rem; font-weight: 700; margin-bottom: 0.8rem; background: #1e293b; color: #f8fafc; }
    .finding-desc { color: #334155; font-size: 0.95rem; margin-bottom: 0.8rem; }
</style>
""", unsafe_allow_html=True)

# --- ROBUST JSON PARSER ---
def robust_json_parser(clean_text):
    try:
        json_match = re.search(r"\{.*\}", clean_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group(0))
        return json.loads(clean_text)
    except Exception:
        return None

# --- EXCEL EXPORT ---
def generate_excel(findings):
    df = pd.DataFrame(findings)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Survey Report')
    return output.getvalue()

# --- STATE MANAGEMENT ---
if "analysis_results" not in st.session_state: st.session_state["analysis_results"] = None
if "analyzed_vessel_name" not in st.session_state: st.session_state["analyzed_vessel_name"] = ""

# --- SOL SIDEBAR (KOTA VE MODEL YÖNETİMİ) ---
with st.sidebar:
    st.title("⚙️ Gelişmiş Ayarlar")
    
    # SUNUM KURTARICI: Canlı API anahtarı kutusu
    live_key = st.text_input("🔑 Canlı/Yedek API Anahtarı", type="password", 
                            help="Eğer kota hatası alırsanız, yeni aldığınız anahtarı buraya yapıştırıp analizi tekrar başlatın.")
    
    selected_model = st.selectbox("Yapay Zeka Modeli", ["gemini-2.0-flash", "gemini-1.5-flash"])
    st.info("💡 İpucu: Sunumda 429 Kota hatası alırsanız, 10 saniye bekleyin veya sol kutuya yeni bir anahtar yapıştırın.")

# API Key Belirleme Mantığı
API_KEY = None
if live_key.strip():
    API_KEY = live_key.strip()
elif "GEMINI_API_KEY" in st.secrets:
    API_KEY = st.secrets["GEMINI_API_KEY"]

if not API_KEY:
    st.error("🚨 API Anahtarı bulunamadı! Lütfen sol menüden veya Secrets alanından bir anahtar sağlayın.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# --- ARAYÜZ ANA GÖVDE ---
st.markdown("""
<div class="header-container">
    <h1 class="header-title">🚢 Klas Kuruluşu Sörveyörü V3.0</h1>
    <p class="header-subtitle">Multi-Document Çapraz Kontrol & Canlı Analiz Platformu</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1.2, 1])

with col1:
    st.subheader("📋 Gemi ve Denetim Bilgileri")
    vessel_name = st.text_input("Gemi Adı (Referans)")
    
    col_a, col_b = st.columns(2)
    with col_a:
        imo_number = st.text_input("IMO Numarası")
    with col_b:
        grt_dwt = st.text_input("GT / DWT")
        
    vessel_type = st.selectbox("Gemi Türü", ["Seçiniz", "General Cargo", "Bulk Carrier", "Oil Tanker", "Chemical Tanker", "Container", "Diğer"])
    survey_type = st.radio("Denetim Tipi (Belgedeki kutucukla eşleşmeli)", ["Annual", "Intermediate", "Renewal", "Initial"])
    surveyor_notes = st.text_area("Sörveyör Notları / Özel Talimatlar", placeholder="Örn: 1.3 pls unmarked, 12.4 check and verify...")

with col2:
    st.subheader("📁 Çapraz Kontrol İçin Belgeleri Yükle")
    uploaded_files = st.file_uploader("Çoklu PDF Rapor/Sertifika Yükleme", type=["pdf"], accept_multiple_files=True)

analyze_btn = st.button("🚀 Belgeleri Canlı Oku ve Çapraz Kontrol Yap", type="primary", use_container_width=True)

# --- MASTER PROMPT (HALÜSİNASYON ENGELLEMELİ) ---
system_instruction = f"""
Sen uluslararası IACS standartlarında çalışan, son derece titiz bir 'Baş Klas Sörveyörü'sün.
Kullanıcı şu an '{survey_type}' denetimi yapıyor. Sörveyörün özel notları: '{surveyor_notes}'

⚠️ KESİN VE KATI KURAL (110 MADDE UYDURMA HATASI ENGELLEME):
SADECE VE SADECE kullanıcının yüklediği PDF belgelerinin içinde fiilen, açıkça geçen maddeleri ve satırları analiz et!
Arka plandaki genel bilgi dağarcığından veya standart klas listelerinden ASLA yeni maddeler türetme, kafandan uydurma!
Eğer yüklenen PDF'te sadece 20 madde varsa, senin çıktı listende de tam olarak sadece o 20 madde olmalıdır. Boşta kalan sörvey şablonlarını buraya dökme!

UYULACAK KURALLAR:
1. BOŞ KUTULAR (☐): 'Uygunsuz' sayma. 'Bilgi (info)' seviyesinde, "Doldurulmamış Olabilir" şeklinde belirt.
2. SEE ATTACHMENT: Sörveyör eke atıf yapmışsa, bunu 'Uygun (success)' kabul et ve açıklamasında "[EK BELGE KONTROLÜ GEREKLİ]" yaz.
3. TİK İŞARETİ (☑): Tablolardaki Tik işaretini daima "Uygundur/Sorunsuz" (success) kabul et.
4. TARİH VE VİZE KONTROLÜ: Sertifikaların geçerlilik tarihi geçmişse veya yıllık vizeleri (Endorsement) eksikse "Kırmızı Alarm" ver.
5. IMO/GEMİ UYUŞMAZLIĞI: Belgelerdeki gemi isimleri veya IMO numaraları uyuşmuyorsa "Kırmızı Alarm" ver.
6. DENETİM TİPİ DOĞRULAMA: Belgenin başında işaretlenmiş denetim türü ile kullanıcının seçtiği '{survey_type}' uyuşmuyorsa "Kırmızı Alarm" üret.
7. SÖRVEYÖR NOTLARI: Sörveyör notlarında belirtilen maddeleri (örn. 1.3 veya 12.4) öncelikle incele ve açıklamasında sörveyörün uyarısına atıf yap.

KATEGORİ VE REFERANS: Her bulguyu şu 5 kategoriden birine ata: [Dokümantasyon, Yapısal, Makine, Emniyet, Çevre]. 'category' ve 'rule' (SOLAS/MARPOL kuralı) alanları zorunludur.

Çıktıyı SADECE aşağıdaki JSON formatında ver:
{{
  "vessel_evaluation": "Genel değerlendirme metni...",
  "findings": [
    {{
      "item_no": "Madde Sıra No",
      "title": "Madde Başlığı",
      "rule": "İlgili Kural Referansı",
      "category": "Dokümantasyon / Yapısal / Makine / Emniyet / Çevre",
      "status": "Durum Açıklaması",
      "severity": "success | info | warning | error | critical",
      "description": "Detaylı açıklama..."
    }}
  ]
}}
"""

if analyze_btn:
    if not uploaded_files:
        st.error("Lütfen en az bir adet PDF belgesi yükleyin.")
    elif vessel_type == "Seçiniz":
        st.error("Lütfen çapraz kontrol için Gemi Türünü seçin.")
    else:
        st.session_state["analysis_results"] = None
        uploaded_gemini_files = []
        tmp_file_paths = []
        
        try:
            with st.spinner("📤 Belgeler işleniyor..."):
                for uploaded_file in uploaded_files:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(uploaded_file.getvalue())
                        tmp_file_paths.append(tmp.name)
                    file_obj = client.files.upload(file=tmp.name)
                    uploaded_gemini_files.append(file_obj)
                time.sleep(2)

            prompt_text = f"Gemi: {vessel_name}, IMO: {imo_number}, GT/DWT: {grt_dwt}, Tür: {vessel_type}, Seçilen Denetim: {survey_type}, Notlar: {surveyor_notes}"
            
            # --- AKILLI YOĞUNLUK VE KOTA KALKANI ---
            max_retries = 4
            response = None
            for attempt in range(max_retries):
                try:
                    with st.spinner("🧠 Yapay zeka belgeleri çapraz kontrol ediyor..."):
                        response = client.models.generate_content(
                            model=selected_model,
                            contents=uploaded_gemini_files + [prompt_text],
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json",
                                system_instruction=system_instruction,
                                temperature=0.1
                            )
                        )
                    break
                except Exception as api_err:
                    if ("429" in str(api_err) or "503" in str(api_err)) and attempt < max_retries - 1:
                        wait = (attempt + 1) * 4
                        st.warning(f"⏳ Sunucu Limiti (429/503). {wait} saniye içinde otomatik tekrar deneniyor...")
                        time.sleep(wait)
                        continue
                    else:
                        raise api_err

            clean_res = response.text.replace('```json', '').replace('```', '').strip()
            parsed_data = robust_json_parser(clean_res)
            
            if parsed_data and "findings" in parsed_data:
                st.session_state["analysis_results"] = parsed_data
                st.session_state["analyzed_vessel_name"] = vessel_name if vessel_name else "Survey"
                st.success("Analiz başarıyla tamamlandı!")
            else:
                st.error("Rapor çözümlenemedi. Lütfen tekrar deneyin veya yedek anahtar kullanın.")
                
        except Exception as e:
            st.error(f"Sistem Hatası: {str(e)}")
            if "429" in str(e):
                st.info("💡 Tavsiye: Sol menüdeki '🔑 Canlı/Yedek API Anahtarı' kutusuna yeni oluşturduğunuz anahtarı yapıştırarak anında kotayı sıfırlayabilirsiniz!")
        finally:
            for f in uploaded_gemini_files:
                try: client.files.delete(name=f.name)
                except: pass
            for p in tmp_file_paths:
                try: 
                    if os.path.exists(p): os.remove(p)
                except: pass

# --- BULGULARI KATEGORİ BAZLI GÖSTERME ---
if st.session_state["analysis_results"] is not None:
    data = st.session_state["analysis_results"]
    findings = data.get("findings", [])
    vessel_name_ref = st.session_state["analyzed_vessel_name"]
    
    st.markdown("## 📊 Çapraz Kontrol ve Sörvey Sonuçları")
    st.info(f"**Yapay Zeka Genel Değerlendirmesi:** {data.get('vessel_evaluation', '')}")
    
    # Metrik Skorları
    c_crit = sum(1 for f in findings if f.get("severity") == "critical")
    c_err = sum(1 for f in findings if f.get("severity") == "error")
    c_warn = sum(1 for f in findings if f.get("severity") == "warning")
    c_succ = sum(1 for f in findings if f.get("severity") == "success")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🚨 Kritik / Alarmler", c_crit)
    m2.metric("❌ Uygunsuzluklar", c_err)
    m3.metric("⚠️ Uyarılar", c_warn)
    m4.metric("✅ Uygun Maddeler", c_succ)
    
    st.write("---")
    
    # Excel Butonu
    excel_data = generate_excel(findings)
    st.download_button(
        label="📥 Raporu Excel Olarak İndir",
        data=excel_data,
        file_name=f"{vessel_name_ref}_Survey_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    st.write("---")
    
    # Kart Tasarım Yardımcısı
    def render_card(f):
        sev = f.get("severity", "info").lower()
        icon = "🚨" if sev == "critical" else "❌" if sev == "error" else "⚠️" if sev == "warning" else "✅" if sev == "success" else "ℹ️"
        st.markdown(f"""
        <div class="finding-card card-{sev}">
            <span class="finding-rule">{f.get('rule', 'Kural Belirtilmemiş')}</span>
            <div class="finding-title">{f.get('item_no', '-')}. {icon} {f.get('title')} ({f.get('status')})</div>
            <div class="finding-desc">{f.get('description')}</div>
        </div>
        """, unsafe_allow_html=True)

    # GRUPLANMIŞ KATEGORİ GÖRÜNÜMÜ
    if findings:
        df_findings = pd.DataFrame(findings)
        
        tab_cat, tab_raw = st.tabs(["📂 Kategori Bazlı Görünüm", "🔍 Tüm Liste (Tablo)"])
        
        with tab_cat:
            if 'category' in df_findings.columns:
                for cat in df_findings['category'].unique():
                    st.markdown(f'<div class="cat-header">{cat}</div>', unsafe_allow_html=True)
                    for f in [x for x in findings if x.get("category") == cat]:
                        render_card(f)
            else:
                for f in findings: render_card(f)
                
        with tab_raw:
            st.dataframe(df_findings)
