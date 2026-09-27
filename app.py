import streamlit as st
import json
import time
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

# API İstemcisi
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
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        toplam_dosya = len(yuklenen_dosyalar)
        
        for index, dosya in enumerate(yuklenen_dosyalar):
            status_text.text(f"İşleniyor ({index + 1}/{toplam_dosya}): {dosya.name}...")
            
            dosya_baytlari = dosya.read()
            mime_tipi = dosya.type if dosya.type else "application/pdf"
            
            prompt = """
            Sen uzman bir muhasebe ve finansal veri ayrıştırma uzmanısın. 
            Görseldeki belgeyi incele. Aşağıdaki alanları çıkarıp SADECE saf bir JSON objesi döndür:
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
            Markdown etiketi kullanma, doğrudan saf JSON döndür.
            """
            
            # Hataya dayanıklı yeniden deneme döngüsü (Retry)
            maksimum_deneme = 3
            basarili = False
            
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
                    
                    temiz_metin = yanit.text.replace("```json", "").replace("```", "").strip()
                    veri = json.loads(temiz_metin)
                    veri["Dosya Adı"] = dosya.name
                    tum_veriler.append(veri)
                    basarili = True
                    break
                    
                except Exception as e:
                    hata_metni = str(e)
                    # 503 Yoğunluk veya 429 Kota uyarısında bekle ve tekrar dene
                    if ("503" in hata_metni or "429" in hata_metni) and deneme < maksimum_deneme - 1:
                        bekleme_suresi = 10 * (deneme + 1)
                        status_text.text(f"Sunucu yoğunluğu/kota sınırı. {bekleme_suresi} sn bekleniyor...")
                        time.sleep(bekleme_suresi)
                        continue
                    else:
                        st.error(f"{dosya.name} işlenirken bir sorun oluştu: {hata_metni[:150]}...")
                        break
            
            progress_bar.progress((index + 1) / toplam_dosya)
            # Çoklu yüklemelerde dakikalık kota sınırına takılmamak için kısa bir nefes payı
            time.sleep(1)
        
        status_text.text("İşlem tamamlandı!")
        
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
