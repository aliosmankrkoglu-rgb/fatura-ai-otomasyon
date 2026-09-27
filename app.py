import streamlit as st
import json
import time
import io
import pandas as pd
from google import genai
from google.genai import types
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="FinAuto AI - Fatura & Fiş Otomasyonu", 
    page_icon="💼", 
    layout="wide"
)

# API İstemcisi
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# --- DÜZENLİ VE KURUMSAL EXCEL OLUŞTURUCU ---
def excel_tablosu_olustur(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Faturalar')
        worksheet = writer.sheets['Faturalar']

        # Başlık stili (Koyu mavi zemin, beyaz kalın yazı)
        baslik_dolgu = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        baslik_yazi = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        govde_yazi = Font(name="Calibri", size=10)
        
        ince_kenarlik = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        # Başlık satırını biçimlendir
        for col_idx in range(1, len(df.columns) + 1):
            hucre = worksheet.cell(row=1, column=col_idx)
            hucre.fill = baslik_dolgu
            hucre.font = baslik_yazi
            hucre.alignment = Alignment(horizontal="center", vertical="center")

        # Veri satırlarını biçimlendir ve sütun genişliklerini otomatik ayarla
        for col in worksheet.columns:
            maksimum_uzunluk = 0
            sutun_harfi = get_column_letter(col[0].column)
            
            for hucre in col:
                hucre.border = ince_kenarlik
                if hucre.row != 1:
                    hucre.font = govde_yazi
                    hucre.alignment = Alignment(vertical="center")
                
                deger_uzunlugu = len(str(hucre.value or ''))
                if deger_uzunlugu > maksimum_uzunluk:
                    maksimum_uzunluk = deger_uzunlugu
            
            # Sütunun rahat okunması için kenar boşluğu
            worksheet.column_dimensions[sutun_harfi].width = max(maksimum_uzunluk + 5, 14)

    return output.getvalue()

# --- ARAYÜZ ---
st.title("💼 FinAuto AI - Akıllı Fatura & Fiş Ayrıştırıcı")
st.markdown("Belgelerinizi yükleyin; yapay zeka verileri otomatik okuyup kurumsal Excel tablosuna dönüştürsün.")

yuklenen_dosyalar = st.file_uploader(
    "Fatura veya makbuz yükleyin (PDF, PNG, JPG)", 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    # KULLANIM SINIRI: Maksimum 2 dosya
    if len(yuklenen_dosyalar) > 2:
        st.error("🛑 Ücretsiz demo sürümünde aynı anda en fazla 2 fatura işleyebilirsiniz. Sınırsız kullanım ve muhasebe entegrasyonu için lütfen iletişime geçin.")
    else:
        st.info(f"İşlenecek belge sayısı: **{len(yuklenen_dosyalar)}**")
        
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
                Sen uzman bir finans ve muhasebe veri ayrıştırma uzmanısın. 
                Belgedeki bilgileri çıkar ve SADECE saf bir JSON objesi döndür:
                {
                  "Dosya Adı": "...",
                  "Fatura No": "...",
                  "Tarih": "...",
                  "Satıcı": "...",
                  "Alıcı": "...",
                  "Para Birimi": "...",
                  "KDV Tutarı": 0.0,
                  "Genel Toplam": 0.0
                }
                Tutar alanlarını sayısal olarak ver. Markdown etiketi (```json) kullanma, doğrudan saf JSON döndür.
                """
                
                maksimum_deneme = 3
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
                        break
                        
                    except Exception as e:
                        hata_metni = str(e)
                        if ("503" in hata_metni or "429" in hata_metni) and deneme < maksimum_deneme - 1:
                            bekleme = 10 * (deneme + 1)
                            status_text.text(f"Yoğunluk sebebiyle bekleniyor ({bekleme} sn)...")
                            time.sleep(bekleme)
                            continue
                        else:
                            st.error(f"{dosya.name} işlenemedi: {hata_metni[:120]}")
                            break
                
                progress_bar.progress((index + 1) / toplam_dosya)
                time.sleep(1)
            
            status_text.text("İşlem başarıyla tamamlandı!")
            
            if tum_veriler:
                df = pd.DataFrame(tum_veriler)
                st.divider()
                st.subheader("📊 Ayrıştırılan Veri Tablosu")
                st.dataframe(df, use_container_width=True)
                
                # Excel oluştur ve indirme butonunu hazırla
                excel_dosyasi = excel_tablosu_olustur(df)
                st.download_button(
                    label="📥 Kurumsal Excel Dosyasını İndir (.xlsx)",
                    data=excel_dosyasi,
                    file_name="islenmis_faturalar.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
