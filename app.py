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

# --- PDF İÇİN HTML RAPOR OLUŞTURUCU ---
def generate_html_report(findings, vessel_name):
    html_content = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <title>{vessel_name} Sörvey Raporu</title>
        <style>
            body {{ font-family: Arial, sans-serif; padding: 20px; color: #333; }}
            h1 {{ color: #1e293b; border-bottom: 2px solid #ccc; padding-bottom: 10px; }}
            .item {{ border: 1px solid #ddd; padding: 15px; margin-bottom: 10px; border-radius: 8px; }}
            .status {{ font-weight: bold; padding: 3px 8px; border-radius: 4px; }}
        </style>
    </head>
    <body>
        <h1>🚢 {vessel_name} - Çapraz Kontrol ve Sörvey Raporu</h1>
    """
    for f in findings:
        html_content += f"""
        <div class="item">
            <div style="font-size: 18px; margin-bottom: 5px;"><b>{f.get('item_no', '-')}</b>. {f.get('title', '')}</div>
            <div><b>Kural:</b> {f.get('rule', '-')}</div>
            <div><b>Durum:</b> <span class="status">{f.get('status', '-')}</span></div>
            <div style="margin-top: 10px;"><b>Açıklama:</b> {f.get('description', '')}</div>
        </div>
        """
    html_content += "</body></html>"
    return html_content.encode('utf-8')

# --- HAFIZA VE EKRAN SIFIRLAMA (STATE MANAGEMENT) ---
if 'analysis_data' not in st.session_state:
    st.session_state['analysis_data'] = None
if 'vessel_name' not in st.session_state:
    st.session_state['vessel_name'] = ""
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

# --- GÜVENLİ API BAĞLANTISI ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    st.error("🚨 API Anahtarı bulunamadı! Lütfen Streamlit ayarlarından 'Secrets' kısmına GEMINI_API_KEY ekleyin.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# --- SOL MENÜ ---
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
    <h1 class="header-title">🚢 Klas Kuruluşu Sörveyörü V2.1</h1>
    <p class="header-subtitle">Multi-Document Çapraz Kontrol & Özel Talimat Motoru</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1.2, 1])

with col1:
    st.subheader("📋 Gemi ve Denetim Bilgileri")
    vessel_name = st.text_input("Gemi Adı (Referans)", value=st.session_state['vessel_name'])
    
    col_a, col_b = st.columns(2)
    with col_a:
        imo_number = st.text_input("IMO Numarası")
    with col_b:
        grt_dwt = st.text_input("GT / DWT")
        
    vessel_type = st.selectbox("Gemi Türü", ["Seçiniz", "General Cargo", "Bulk Carrier", "Oil Tanker", "Chemical Tanker", "Container", "Diğer"])
    
    surveyor_notes = st.text_area("✍️ Sörveyör Özel Notları / Talimatları", 
                                  placeholder="Örn: Raporu kontrol et ama özellikle 1.3 maddesine dikkat edilsin, 'unmarked' olabilir.")

with col2:
    st.subheader("📁 Çapraz Kontrol İçin Belgeleri Yükle")
    st.info("Sörvey raporu, Narrative Report ve Sertifikaları (PDF) aynı anda yükleyebilirsiniz.")
    
    uploaded_files = st.file_uploader("Çoklu PDF Yükleme", type=["pdf"], accept_multiple_files=True, key=f"uploader_{st.session_state['uploader_key']}")
    
    if st.button("🔄 Ekranı Temizle / Yeni Belge Yükle", use_container_width=True):
        st.session_state['analysis_data'] = None
        st.session_state['vessel_name'] = ""
        st.session_state['uploader_key'] += 1 
        st.rerun()

analyze_btn = st.button("🚀 Belgeleri Oku ve Çapraz Kontrol Yap", type="primary", use_container_width=True)

# --- MASTER PROMPT ---
system_instruction = f"""
Sen uluslararası IACS standartlarında çalışan, son derece titiz bir 'Baş Klas Sörveyörü'sün.
Amacın sana verilen belgeleri teker teker incelemek ve belgeler arası ÇAPRAZ KONTROL (Double-Check) yapmaktır.

⚠️ SÖRVEYÖRÜN ÖZEL TALİMATI / NOTLARI:
"{surveyor_notes}"
(LÜTFEN DİKKAT: Raporun TÜM maddelerini eksiksiz inceleyeceksin, ancak yukarıdaki notta belirtilen maddelere veya konulara EKSTRA ÖZEN göster. Bulgularında bu özel duruma değin.)

⚠️ DİKKAT (KATI KURAL): 
SADECE VE SADECE sana yüklenen PDF belgelerinin içindeki maddeleri incele! 
Kendi hafızandan, eski klas dökümanlarından veya genel kurallardan ASLA yeni maddeler (örneğin 110 maddelik şablonlar) UYDURMA! Belgede kaç madde varsa sadece onları listele ve HİÇBİR MADDEYİ ATLAMA.

AŞAĞIDAKİ ALTIN KURALLARA KESİNLİKLE UYACAKSIN:
1. BOŞ KUTULAR (☐): Eğer bir formda kutu boş bırakılmışsa bunu 'Düzeltilmeli' olarak değil, "Uygun Değil" veya "Gözden Kaçmış" olarak değerlendir.
2. SEE ATTACHMENT: Sörveyör eke atıf yapmışsa (SEE ATTACHMENT), bunu 'Uygun' kabul et.
3. TİK İŞARETİ (☑): Tablolardaki Tik (☑) işaretini daima 'Uygun' olarak kabul et.

⚠️ DURUM (STATUS) KATEGORİSİ KURALI:
Çıktıdaki her bulgunun "status" alanına SADECE şu üç kelimeden birini yazabilirsin:
- "Uygun"
- "Uygun Değil"
- "Düzeltilmeli"
Başka hiçbir kelime kullanma!

Çıktıyı SADECE aşağıdaki JSON formatında ver, başka hiçbir metin ekleme:
{{
  "cross_check_status": "Başarılı / Başarısız",
  "vessel_evaluation": "Genel değerlendirme...",
  "compliance_score": 85,
  "findings": [
    {{
      "item_no": "Madde Sıra No (örn: 1, 2, 3...)",
      "title": "Madde Başlığı",
      "rule": "İlgili Kural (Örn: SOLAS Ch. II-2 Reg. 10)",
      "status": "Uygun | Uygun Değil | Düzeltilmeli",
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
        uploaded_gemini_files = []
        tmp_file_paths = []
        
        try:
            with st.spinner("📤 Belgeler güvenli sunucuya yükleniyor..."):
                for uploaded_file in uploaded_files:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(uploaded_file.getvalue())
                        tmp_file_paths.append(tmp.name)
                        
                    file_obj = client.files.upload(file=tmp.name)
                    uploaded_gemini_files.append(file_obj)
                
                time.sleep(3) 

            prompt_text = f"Kullanıcının Girdiği Gemi Referans Bilgileri:\nAdı: {vessel_name}\nIMO: {imo_number}\nGT/DWT: {grt_dwt}\nTür: {vessel_type}\nÖzel Notlar: {surveyor_notes}\nLütfen belgelerdeki İSTİSNASIZ TÜM maddeleri incele ve JSON dön."
            
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
            
            clean_response = response.text.replace('```json', '').replace('
```', '').strip()
            parsed_data = robust_json_parser(clean_response)
            
            if parsed_data and "findings" in parsed_data:
                st.session_state['analysis_data'] = parsed_data
                st.session_state['vessel_name'] = vessel_name
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


# --- EĞER HAFIZADA VERİ VARSA EKRANA BASTIR ---
if st.session_state['analysis_data']:
    parsed_data = st.session_state['analysis_data']
    findings = parsed_data["findings"]
    current_vessel = st.session_state['vessel_name']
    
    st.markdown("## 📊 V2 Çapraz Kontrol ve Sörvey Raporu")
    st.info(f"**Yapay Zeka Değerlendirmesi:** {parsed_data.get('vessel_evaluation', '')}")
    
    c_crit = sum(1 for f in findings if f.get("severity") == "critical")
    c_err = sum(1 for f in findings if f.get("severity") == "error")
    c_warn = sum(1 for f in findings if f.get("severity") == "warning")
    c_succ = sum(1 for f in findings if f.get("severity") == "success")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🚨 Kırmızı Alarm (Kritik)", c_crit)
    m2.metric("❌ Uygunsuzluklar", c_err)
    m3.metric("⚠️ Uyarılar", c_warn)
    m4.metric("✅ Uygun Maddeler", c_succ)
    
    st.write("---")
    
    # BUTONLAR YANYANA
    btn_col1, btn_col2 = st.columns(2)
    
    with btn_col1:
        excel_data = generate_excel(findings, current_vessel)
        st.download_button(
            label="📥 Raporu Excel Olarak İndir",
            data=excel_data,
            file_name=f"{current_vessel}_Survey_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
    with btn_col2:
        html_data = generate_html_report(findings, current_vessel)
        st.download_button(
            label="📄 PDF/Yazdırılabilir Çıktı (HTML) İndir",
            data=html_data,
            file_name=f"{current_vessel}_Rapor.html",
            mime="text/html",
            type="secondary",
            use_container_width=True
        )
        st.caption("İnen dosyaya tıklayıp tarayıcıda açın ve `Ctrl+P` yaparak mükemmel bir PDF alabilirsiniz.")
    
    st.write("---")
    
    # --- KATEGORİLERE GÖRE AYRILMIŞ SEKMELER (TABS) ---
    st.markdown("### 📋 Bulgu Kategorileri")
    tab1, tab2, tab3 = st.tabs(["✅ Uygun Olanlar", "❌ Uygun Olmayanlar", "⚠️ Düzeltilmesi Gerekenler"])
    
    def render_cards(filtered_findings):
        if not filtered_findings:
            st.success("Bu kategoride bulgu yok.")
            return
            
        for f in filtered_findings:
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

    with tab1:
        render_cards([f for f in findings if f.get("status") == "Uygun"])
        
    with tab2:
        render_cards([f for f in findings if f.get("status") == "Uygun Değil"])
        
    with tab3:
        render_cards([f for f in findings if f.get("status") == "Düzeltilmeli"])
