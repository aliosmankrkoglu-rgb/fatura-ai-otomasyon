import streamlit as st
import json
import pandas as pd
from google import genai
from google.genai import types

st.set_page_config(
    page_title="FinAuto AI - Fatura & Fiş Otomasyonu", 
    page_icon="💼", 
    layout="wide"
)

st.title("💼 FinAuto AI - Sıfır Veri Girişi Fatura Ayrıştırıcı")
st.markdown("Faturalarınızı ve fişlerinizi yükleyin; yapay zeka saniyeler içinde muhasebe tablosuna dönüştürsün.")

# API İstemcisi (Güvenli Kasadan Çeker)
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# Dosya yükleme alanı
yuklenen_dosyalar = st.file_uploader(
    "Fatura veya makbuzları sürükleyip bırakın (PDF, PNG, JPG)", 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    st.info(f"Toplam **{len(yuklenen_dosyalar)}** adet dosya seçildi.")
    
    if st.button("🚀 Faturaları Otomatik İşle", type="primary"):
        tum_veriler = []
        
        # İlerleme çubuğu
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        toplam_dosya = len(yuklenen_dosyalar)
        
        for index, dosya in enumerate(yuklenen_dosyalar):
            status_text.text(f"İşleniyor ({index + 1}/{toplam_dosya}): {dosya.name}...")
            
            dosya_baytlari = dosya.read()
            mime_tipi = dosya.type if dosya.type else "application/pdf"
            
            prompt = """
            Sen uzman bir muhasebe ve finansal veri ayrıştırma uzmanısın. 
            Görseldeki belgeyi (fatura, perakende satış fişi, makbuz veya dekont olabilir) dikkatle incele.

            Aşağıdaki kurallara göre bilgileri çıkar ve SADECE saf bir JSON objesi döndür:
            1. "Fatura No": Belgede geçen Fatura No, Fiş No, Belge No veya Belge Seri/Sıra numarasını yaz. Bulamazsan onay kodunu yaz.
            2. "Tarih": Gün-Ay-Yıl formatında işlem tarihini yaz.
            3. "Satıcı": Faturayı kesen işletmenin, dükkanın veya şirketin adını yaz (en üstteki başlık).
            4. "Alıcı": Belge kime kesilmişse yaz. Perakende fiş ise "Nihai Tüketici" yaz.
            5. "Para Birimi": TRY, USD, EUR vb. para birimini belirt.
            6. "KDV": Belgedeki toplam KDV tutarını yaz (örn: 130.00). Bulamazsan 0 yaz.
            7. "Toplam Tutar": Ödenen nihai genel toplam tutarı yaz (örn: 780.00).

            JSON Şablonu:
            {
              "Dosya Adı": "...",
              "Fatura No": "...",
              "Tarih": "...",
              "Satıcı": "...",
              "Alıcı": "...",
              "Para Birimi": "...",
              "KDV": "...",
              "Toplam Tutar": "..."
            }
            Asla markdown (```json) veya fazladan açıklama yazma.
            """
            
            try:
                yanit = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=[
                        types.Part.from_bytes(
                            data=dosya_baytlari,
                            mime_type=mime_tipi
                        ),
                        prompt
                    ]
                )
                
                temiz_metin = yanit.text.replace("```json", "").replace("```", "").strip()
                veri = json.loads(temiz_metin)
                veri["Dosya Adı"] = dosya.name
                tum_veriler.append(veri)
            except Exception as e:
                st.error(f"{dosya.name} okunurken bir hata oluştu: {e}")
            
            # Çubuğu güncelle
            progress_bar.progress((index + 1) / toplam_dosya)
        
        status_text.text("İşlem başarıyla tamamlandı!")
        
        if tum_veriler:
            df = pd.DataFrame(tum_veriler)
            
            st.divider()
            st.subheader("📊 Ayrıştırılan Veri Tablosu")
            st.dataframe(df, use_container_width=True)
            
            csv = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                label="📥 Tabloyu Excel / CSV Olarak İndir",
                data=csv,
                file_name="muhasebe_aktarim_listesi.csv",
                mime="text/csv"
            )
            import time

# İstek gönderme kısmını döngüye alıyoruz:
maksimum_deneme = 3
yanit = None

for deneme in range(maksimum_deneme):
    try:
        yanit = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=[
                types.Part.from_bytes(
                    data=dosya_baytlari,
                    mime_type=mime_tipi
                ),
                prompt
            ]
        )
        break  # Başarılı olursa döngüden çık
    except Exception as e:
        if "503" in str(e) and deneme < maksimum_deneme - 1:
            time.sleep(2)  # 2 saniye bekle ve tekrar dene
            continue
        else:
            raise e
