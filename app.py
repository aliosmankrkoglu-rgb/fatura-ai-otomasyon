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
    page_title="FinAuto AI - Akıllı Muhasebe Motoru", 
    page_icon="💼", 
    layout="wide"
)

# API İstemcisi
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# --- EXCEL OLUŞTURUCU ---
def excel_tablosu_olustur(df, sheet_name="Muhasebe_Fisi"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]

        baslik_dolgu = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        baslik_yazi = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        govde_yazi = Font(name="Calibri", size=10)
        
        ince_kenarlik = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for col_idx in range(1, len(df.columns) + 1):
            hucre = worksheet.cell(row=1, column=col_idx)
            hucre.fill = baslik_dolgu
            hucre.font = baslik_yazi
            hucre.alignment = Alignment(horizontal="center", vertical="center")

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
            
            worksheet.column_dimensions[sutun_harfi].width = max(maksimum_uzunluk + 5, 14)

    return output.getvalue()

# --- ARAYÜZ ---
st.title("💼 FinAuto AI - Akıllı Muhasebe Fiş Motoru")
st.markdown("Faturaları yükleyin; sistem harcama türünü (Akaryakıt, Yemek, Mal Alışı, Kargo vb.) otomatik tespit edip kurumsal yevmiye fişini çıkarsın.")

yuklenen_dosyalar = st.file_uploader(
    "Fatura veya fiş yükleyin (PDF, PNG, JPG)", 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    if len(yuklenen_dosyalar) > 2:
        st.error("🛑 Ücretsiz demo sürümünde aynı anda en fazla 2 fatura işleyebilirsiniz.")
    else:
        st.info(f"İşlenecek belge sayısı: **{len(yuklenen_dosyalar)}**")
        
        if st.button("🚀 Akıllı Fiş Analizini Başlat", type="primary"):
            ham_veriler = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            toplam_dosya = len(yuklenen_dosyalar)
            
            for index, dosya in enumerate(yuklenen_dosyalar):
                status_text.text(f"İşleniyor ({index + 1}/{toplam_dosya}): {dosya.name}...")
                
                dosya_baytlari = dosya.read()
                mime_tipi = dosya.type if dosya.type else "application/pdf"
                
                prompt = """
                Sen üst düzey bir mali müşavir ve muhasebe uzmanısın. Belgeyi incele.
                Harcama türünü ve Tek Düzen Hesap Planı'na göre en uygun hesap kodunu otomatik belirle:
                - Ticari mal alımı ise: "153.01 Ticari Mallar"
                - Akaryakıt/Yakıt ise: "770.01 Akaryakıt Giderleri"
                - Yemek/Temsil/Ağırlama ise: "770.02 Yemek ve Ağırlama"
                - Kırtasiye/Ofis Malzemesi ise: "770.03 Kırtasiye Giderleri"
                - Kargo/Nakliye ise: "770.04 Kargo ve Ulaşım"
                - Demirbaş/Ekipman alımı ise: "255.01 Demirbaşlar"
                - Diğer genel giderler için: "770.99 Genel Giderler"

                SADECE şu JSON şablonunu döndür:
                {
                  "Fatura No": "...",
                  "Tarih": "...",
                  "Satıcı": "...",
                  "VKN_TCKN": "...",
                  "Belge Türü": "Alış Faturası / Gider Fişi",
                  "Önerilen Hesap Kodu": "770.01",
                  "Hesap Adı": "Akaryakıt Giderleri",
                  "Matrah": 0.0,
                  "KDV Orani": 20,
                  "KDV Tutarı": 0.0,
                  "Genel Toplam": 0.0
                }
                Tutar alanlarını float yap. Markdown etiketi ekleme.
                """
                
                maksimum_deneme = 3
                for deneme in range(maksimum_deneme):
                    try:
                        yanit = client.models.generate_content(
                            model="gemini-3.5-flash-lite",
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
                        ham_veriler.append(veri)
                        break
                        
                    except Exception as e:
                        hata_metni = str(e)
                        if ("503" in hata_metni or "429" in hata_metni) and deneme < maksimum_deneme - 1:
                            bekleme = 3 * (deneme + 1)
                            time.sleep(bekleme)
                            continue
                        else:
                            st.error(f"{dosya.name} okunamadı: {hata_metni[:120]}")
                            break
                
                progress_bar.progress((index + 1) / toplam_dosya)
                time.sleep(1)
            
            status_text.text("Analiz başarıyla tamamlandı!")
            
            if ham_veriler:
                fis_satirlari = []
                fis_sira_no = 1
                
                for item in ham_veriler:
                    fatura_no = item.get("Fatura No", "")
                    tarih = item.get("Tarih", "")
                    satici = item.get("Satıcı", "")
                    hesap_kodu = item.get("Önerilen Hesap Kodu", "770.99")
                    hesap_adi = item.get("Hesap Adı", "Genel Giderler")
                    kdv_orani = item.get("KDV Orani", 20)
                    
                    matrah = float(item.get("Matrah", 0.0) or 0.0)
                    kdv = float(item.get("KDV Tutarı", 0.0) or 0.0)
                    genel_toplam = float(item.get("Genel Toplam", 0.0) or (matrah + kdv))
                    
                    # 1. Satır: Akıllı Tespit Edilen Gider veya Mal Hesabı (Borç)
                    fis_satirlari.append({
                        "Fiş No": fis_sira_no,
                        "Tarih": tarih,
                        "Hesap Kodu": hesap_kodu,
                        "Hesap Adı": hesap_adi,
                        "Açıklama": f"{satici} - Ftr No: {fatura_no}",
                        "Borç": matrah,
                        "Alacak": 0.0
                    })
                    
                    # 2. Satır: İndirilecek KDV (Borç)
                    if kdv > 0:
                        kdv_hesap_kodu = f"191.{int(kdv_orani):02d}" if kdv_orani else "191.20"
                        fis_satirlari.append({
                            "Fiş No": fis_sira_no,
                            "Tarih": tarih,
                            "Hesap Kodu": kdv_hesap_kodu,
                            "Hesap Adı": f"%{kdv_orani} İndirilecek KDV",
                            "Açıklama": f"{satici} - KDV",
                            "Borç": kdv,
                            "Alacak": 0.0
                        })
                    
                    # 3. Satır: Satıcı / Kasa (Alacak)
                    fis_satirlari.append({
                        "Fiş No": fis_sira_no,
                        "Tarih": tarih,
                        "Hesap Kodu": "320.01.001",
                        "Hesap Adı": satici,
                        "Açıklama": f"{satici} - Ftr No: {fatura_no}",
                        "Borç": 0.0,
                        "Alacak": genel_toplam
                    })
                    
                    fis_sira_no += 1
                
                df_sonuc = pd.DataFrame(fis_satirlari)
                
                st.divider()
                st.subheader("📊 Otomatik Sınıflandırılmış Yevmiye Fişi")
                st.dataframe(df_sonuc, use_container_width=True)
                
                # Bakiye kontrolü
                toplam_borc = df_sonuc["Borç"].sum()
                toplam_alacak = df_sonuc["Alacak"].sum()
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Toplam Borç", f"{toplam_borc:,.2f} TL")
                col2.metric("Toplam Alacak", f"{toplam_alacak:,.2f} TL")
                if abs(toplam_borc - toplam_alacak) < 0.05:
                    col3.success("✅ Fiş Bakiyesi Dengeli (Borç = Alacak)")
                else:
                    col3.warning("⚠️ Bakiye Farkı Var")

                excel_cikti = excel_tablosu_olustur(df_sonuc, sheet_name="Fis_Aktarim")
                st.download_button(
                    label="📥 Otomatik Muhasebe Fişini İndir (.xlsx)",
                    data=excel_cikti,
                    file_name="akilli_muhasebe_fisi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
