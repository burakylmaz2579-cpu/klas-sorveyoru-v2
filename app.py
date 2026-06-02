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
    page_title="Klas Sörveyörü V2 | Double-Check Engine",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- V2 PREMIUM TASARIM (CSS) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"], .stApp { font-family: 'Plus Jakarta Sans', sans-serif; }
    .header-container { background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 2.5rem; border-radius: 16px; color: white; margin-bottom: 2rem; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1); }
    .header-title { font-size: 2.2rem; font-weight: 700; margin: 0; background: linear-gradient(to right, #ffffff, #94a3b8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .header-subtitle { color: #94a3b8; font-size: 1.1rem; margin-top: 0.5rem; }
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

# --- EXCEL EXPORT FONKSİYONU ---
def generate_excel(findings, vessel_info):
    df = pd.DataFrame(findings)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Survey Report')
    processed_data = output.getvalue()
    return processed_data

# --- GÜVENLİ API BAĞLANTISI (SECRETS) ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    st.error("🚨 API Anahtarı bulunamadı! Lütfen Streamlit ayarlarından 'Secrets' kısmına GEMINI_API_KEY ekleyin.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# --- STATE MANAGEMENT FOR ANALYSIS RESULTS ---
if "analysis_results" not in st.session_state:
    st.session_state["analysis_results"] = None
if "analyzed_vessel_name" not in st.session_state:
    st.session_state["analyzed_vessel_name"] = ""

# --- SOL MENÜ (MODEL SEÇİCİ) ---
with st.sidebar:
    st.title("⚙️ Sistem Ayarları")
    st.markdown("Google API bölgenize göre çalışan modeli seçin.")
    selected_model = st.selectbox(
        "Yapay Zeka Modeli", 
        ["gemini-2.5-flash", "gemini-flash-latest", "gemini-1.5-flash", "gemini-1.5-pro", "gemini-pro-latest"]
    )
    st.info("💡 Not: Analiz sırasında '404 NOT FOUND' hatası alırsanız, kodu değiştirmeden bu menüden farklı bir model seçip 'Analiz Et' butonuna tekrar basmanız yeterlidir.")

# --- ARAYÜZ ---
st.markdown("""
<div class="header-container">
    <h1 class="header-title">🚢 Klas Kuruluşu Sörveyörü V2</h1>
    <p class="header-subtitle">Multi-Document Çapraz Kontrol (Double-Check) Motoru</p>
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
    
    surveyor_notes = st.text_area("Sörveyör Özel Notları / Talimatları (Opsiyonel)", 
                                  placeholder="Örn: 1.3 maddesini check and verify yapın veya 12.4 unmarked olarak değerlendirilsin.",
                                  help="Sörveyörün veya PHRS yetkililerinin eklediği notları/talimatları buraya yazabilirsiniz. Yapay zeka tüm belgeyle birlikte bu notları da analiz edecektir.")

with col2:
    st.subheader("📁 Çapraz Kontrol İçin Belgeleri Yükle")
    st.info("Sörvey raporu, Narrative Report ve Sertifikaları (PDF) aynı anda yükleyebilirsiniz.")
    uploaded_files = st.file_uploader("Çoklu PDF Yükleme", type=["pdf"], accept_multiple_files=True)

analyze_btn = st.button(f"🚀 Belgeleri Oku ve Çapraz Kontrol Yap", type="primary", use_container_width=True)

# --- MASTER PROMPT (V2 KURALLARI) ---
system_instruction = """
Sen uluslararası IACS standartlarında çalışan, son derece titiz bir 'Baş Klas Sörveyörü'sün.
Sana bir veya birden fazla gemi evrakı (Sörvey Kontrol Listesi, Statutory Sertifikalar, Narrative Raporlar) verilecek.
Amacın bu belgeleri teker teker incelemek ve belgeler arası ÇAPRAZ KONTROL (Double-Check) yapmaktır.

AŞAĞIDAKİ ALTIN KURALLARA KESİNLİKLE UYACAKSIN:
1. BOŞ KUTULAR (☐): Eğer bir formda kutu boş bırakılmışsa bunu 'Uygunsuz' sayma. Bunu 'Bilgi (info)' seviyesinde, "Gözden Kaçmış/Doldurulmamış Olabilir" şeklinde belirt.
2. SEE ATTACHMENT: Sörveyör değer girmek yerine eke atıf yapmışsa (SEE ATTACHMENT), bunu 'Uygun (success)' kabul et ve açıklamasında "[EK BELGE KONTROLÜ GEREKLİ]" yaz.
3. YÜK LİSTESİ (IMSBC/DANG): Sayfalar dolusu yük isimlerini tek tek listeleme. SADECE geminin türü (örn: General Cargo) ile onaylı yüklerin doğası çelişiyorsa "Kırmızı Alarm" ver.
4. TİK İŞARETİ (☑): Şablonun açıklaması ne olursa olsun, tablolardaki Tik (☑) işaretini daima "Uygundur/Sorunsuz" (success) olarak kabul et.
5. TARİH VE VİZE KONTROLÜ (KRİTİK!): Sertifikaların bitiş tarihlerini (Valid until) ve Annual Endorsement (Yıllık Vize) sayfalarını KONTROL ET. Süresi geçmişse veya vize atılmamışsa "Kırmızı Alarm" ver.
6. IMO/GEMİ UYUŞMAZLIĞI (KRİTİK!): Yüklenen belgelerdeki gemi isimleri veya IMO'lar birbiriyle uyuşmuyorsa, "Kırmızı Alarm: Farklı gemi evrakları yüklendi" uyarısı ver.
7. EKİPMAN ÇELİŞKİSİ: Sertifikadaki cihaz markası ile saha raporundaki marka uyuşmuyorsa bunu "Uyarı (warning)" olarak raporla.
8. TONAJ KONTROLÜ: Kullanıcının girdiği GT ve DWT değerlerini dikkate alarak SOLAS/MARPOL kurallarının uygulanabilirliğini (örn. 400 GT altı/üstü kural farkları) mutlaka kontrol et.

ÖNEMLİ ZORUNLULUK: Analiz ettiğin HER MADDENİN yanına KESİNLİKLE ilgili kuralı (SOLAS Bölüm..., MARPOL Ek..., IMO MSC.Circ...) referans olarak ekleyeceksin. Referanssız madde kalmayacak.

Çıktıyı SADECE aşağıdaki JSON formatında ver, başka hiçbir metin ekleme:
{
  "cross_check_status": "Başarılı / Başarısız",
  "vessel_evaluation": "Genel değerlendirme...",
  "compliance_score": 85,
  "findings": [
    {
      "item_no": "Madde Sıra No (örn: 1, 2, 3...)",
      "title": "Madde Başlığı",
      "rule": "İlgili Kural (Örn: SOLAS Ch. II-2 Reg. 10)",
      "status": "Uygun | Bilgi | Uyarı | Uygunsuz | Kırmızı Alarm",
      "severity": "success | info | warning | error | critical",
      "description": "Detaylı açıklama..."
    }
  ]
}
"""

if analyze_btn:
    if not uploaded_files:
        st.error("Lütfen en az bir adet PDF belgesi yükleyin.")
    elif vessel_type == "Seçiniz":
        st.error("Lütfen çapraz kontrol için Gemi Türünü seçin.")
    else:
        st.session_state["analysis_results"] = None  # Reset state before new run
        uploaded_gemini_files = []
        tmp_file_paths = []
        
        try:
            with st.spinner(f"📤 Belgeler güvenli sunucuya yükleniyor..."):
                for uploaded_file in uploaded_files:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(uploaded_file.getvalue())
                        tmp_file_paths.append(tmp.name)
                        
                    file_obj = client.files.upload(file=tmp.name)
                    uploaded_gemini_files.append(file_obj)
                
                time.sleep(3) 

            prompt_text = f"""Kullanıcının Girdiği Gemi Referans Bilgileri:
Adı: {vessel_name}
IMO: {imo_number}
GT/DWT: {grt_dwt}
Tür: {vessel_type}

Sörveyör Özel Notları / Ek Talimatları:
{surveyor_notes if surveyor_notes else 'Yok'}

Lütfen tüm belgeleri analiz edip yukarıdaki özel sörveyör notlarını/talimatlarını da dikkate alarak kurallara uygun JSON dön."""
            
            contents = uploaded_gemini_files.copy()
            contents.append(prompt_text)

            with st.spinner(f"🧠 {selected_model} belgeleri çapraz kontrol ediyor. Bu işlem kapsamlı olduğu için 1-2 dakika sürebilir..."):
                response = client.models.generate_content(
                    model=selected_model, 
                    contents=contents,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        system_instruction=system_instruction,
                        temperature=0.1
                    )
                )
            
            clean_response = response.text.replace('```json', '').replace('```', '').strip()
            parsed_data = robust_json_parser(clean_response)
            
            if parsed_data and "findings" in parsed_data:
                st.session_state["analysis_results"] = parsed_data
                st.session_state["analyzed_vessel_name"] = vessel_name if vessel_name else "Survey"
                st.success("Analiz başarıyla tamamlandı!")
            else:
                st.error("JSON ayrıştırma hatası oluştu.")
                st.code(clean_response)

        except Exception as e:
            st.error(f"Sistem Hatası: {str(e)}")
            
        finally:
            for f in uploaded_gemini_files:
                try: client.files.delete(name=f.name)
                except: pass
            for p in tmp_file_paths:
                try: 
                    if os.path.exists(p): os.remove(p)
                except: pass

# --- BULGULARI VE RAPORU GÖSTERME (SESSION STATE ÜZERİNDEN) ---
if st.session_state["analysis_results"] is not None:
    parsed_data = st.session_state["analysis_results"]
    findings = parsed_data.get("findings", [])
    vessel_name_ref = st.session_state["analyzed_vessel_name"]
    
    st.markdown("## 📊 V2 Çapraz Kontrol ve Sörvey Raporu")
    st.info(f"**Yapay Zeka Değerlendirmesi:** {parsed_data.get('vessel_evaluation', '')}")
    
    c_crit = sum(1 for f in findings if f.get("severity") == "critical")
    c_err = sum(1 for f in findings if f.get("severity") == "error")
    c_warn = sum(1 for f in findings if f.get("severity") == "warning")
    c_succ = sum(1 for f in findings if f.get("severity") == "success")
    c_info = sum(1 for f in findings if f.get("severity") == "info")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🚨 Kırmızı Alarm (Kritik)", c_crit)
    m2.metric("❌ Uygunsuzluklar", c_err)
    m3.metric("⚠️ Uyarılar", c_warn)
    m4.metric("✅ Uygun Maddeler", c_succ)
    
    st.write("---")
    
    excel_data = generate_excel(findings, vessel_name_ref)
    st.download_button(
        label="📥 Raporu Excel Olarak İndir (Geçmiş Kayıtlar İçin)",
        data=excel_data,
        file_name=f"{vessel_name_ref}_Survey_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    st.write("---")
    
    # Yardımcı bulgu kartı çizim fonksiyonu
    def render_finding(f):
        sev = f.get("severity", "info").lower()
        title = f.get("title", "")
        rule = f.get("rule", "Kural Belirtilmemiş")
        desc = f.get("description", "")
        status = f.get("status", "")
        item_no = f.get("item_no", "-")
        
        icon = "🚨" if sev == "critical" else "❌" if sev == "error" else "⚠️" if sev == "warning" else "✅" if sev == "success" else "ℹ️"
        
        st.markdown(f"""
        <div class="finding-card card-{sev}">
            <span class="finding-rule">{rule}</span>
            <div class="finding-title">{item_no}. {icon} {title} ({status})</div>
            <div class="finding-desc">{desc}</div>
        </div>
        """, unsafe_allow_html=True)

    # --- KATEGORİLERE AYRILMIŞ SEKME (TAB) GÖRÜNÜMÜ ---
    tab_all, tab_bakilmali, tab_warning, tab_success = st.tabs([
        f"🔍 Tüm Bulgular ({len(findings)})",
        f"🚨 Bakılmalı / Sorunlu ({c_crit + c_err + c_info})",
        f"⚠️ Uyarılar ({c_warn})",
        f"✅ Uygun Maddeler ({c_succ})"
    ])
    
    with tab_all:
        if not findings:
            st.info("Gösterilecek bulgu bulunamadı.")
        else:
            for f in findings:
                render_finding(f)
                
    with tab_bakilmali:
        bakilmali_findings = [f for f in findings if f.get("severity") in ["critical", "error", "info"]]
        if not bakilmali_findings:
            st.success("Bakılması gereken kritik bir durum, hata veya eksik bilgi bulunmamaktadır.")
        else:
            for f in bakilmali_findings:
                render_finding(f)
                
    with tab_warning:
        warn_findings = [f for f in findings if f.get("severity") == "warning"]
        if not warn_findings:
            st.success("Herhangi bir uyarı derecesinde bulgu bulunmamaktadır.")
        else:
            for f in warn_findings:
                render_finding(f)
                
    with tab_success:
        succ_findings = [f for f in findings if f.get("severity") == "success"]
        if not succ_findings:
            st.info("Herhangi bir uygun olarak onaylanmış madde bulunmamaktadır.")
        else:
            for f in succ_findings:
                render_finding(f)

