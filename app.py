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
    page_title="FinAuto AI - Muhasebe Fiş Entegratörü", 
    page_icon="💼", 
    layout="wide"
)

# API İstemcisi
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# --- YAN PANEL: FİRMA / PROGRAM AYARLARI ---
with st.sidebar:
    st.title("⚙️ Muhasebe Ayarları")
    program_secimi = st.selectbox(
        "Aktarım Formatı Seçin",
        ["ETA / Luca / Zirve Uyumlu (Fiş Aktarımı)", "Standart Fatura Listesi"]
    )
    st.markdown("---")
    st.markdown("#### Varsayılan Hesap Kodları")
    varsayilan_gider_kodu = st.text_input("Gider / Mal Hesabı", value="770.01.001")
    varsayilan_kdv_kodu = st.text_input("KDV Hesabı", value="191.20.001")
    varsayilan_cari_kodu = st.text_input("Satıcı / Kasa Hesabı", value="320.01.001")
    st.caption("Firma bazlı özel hesap planı eşleştirmesi entegrasyon paketinde otomatik tanımlanır.")

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

# --- ANA EKRAN ---
st.title("💼 FinAuto AI - Akıllı Muhasebe Fiş Aktarım Prototipi")
st.markdown("Faturaları yükleyin; sistem doğrudan muhasebe programınızın aktarım modülüne uygun Borç/Alacak yevmiye fişi üretsin.")

yuklenen_dosyalar = st.file_uploader(
    "Fatura veya makbuz yükleyin (PDF, PNG, JPG)", 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    if len(yuklenen_dosyalar) > 2:
        st.error("🛑 Ücretsiz demo sürümünde aynı anda en fazla 2 fatura işleyebilirsiniz.")
    else:
        st.info(f"İşlenecek belge sayısı: **{len(yuklenen_dosyalar)}**")
        
        if st.button("🚀 Muhasebe Fişini Otomatik Oluştur", type="primary"):
            ham_veriler = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            toplam_dosya = len(yuklenen_dosyalar)
            
            for index, dosya in enumerate(yuklenen_dosyalar):
                status_text.text(f"İşleniyor ({index + 1}/{toplam_dosya}): {dosya.name}...")
                
                dosya_baytlari = dosya.read()
                mime_tipi = dosya.type if dosya.type else "application/pdf"
                
                prompt = """
                Sen uzman bir muhasebe uzmanısın. Belgeyi dikkatle incele ve SADECE saf bir JSON objesi döndür:
                {
                  "Fatura No": "...",
                  "Tarih": "...",
                  "Satıcı": "...",
                  "Vergi No / TCKN": "...",
                  "Matrah": 0.0,
                  "KDV Orani": 20,
                  "KDV Tutarı": 0.0,
                  "Genel Toplam": 0.0
                }
                Tutar alanlarını kesinlikle sayısal (float) döndür. Markdown etiketi kullanma.
                """
                
                maksimum_deneme = 3
                for deneme in range(maksimum_deneme):
                    try:
                        yanit = client.models.generate_content(
                            model="gemini-2.5-flash",
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
                            status_text.text(f"Hızlı deneme yapılıyor ({bekleme} sn)...")
                            time.sleep(bekleme)
                            continue
                        else:
                            st.error(f"{dosya.name} okunamadı: {hata_metni[:120]}")
                            break
                
                progress_bar.progress((index + 1) / toplam_dosya)
                time.sleep(1)
            
            status_text.text("Fiş aktarımı hazırlandı!")
            
            if ham_veriler:
                if program_secimi == "ETA / Luca / Zirve Uyumlu (Fiş Aktarımı)":
                    fis_satirlari = []
                    fis_sira_no = 1
                    
                    for item in ham_veriler:
                        fatura_no = item.get("Fatura No", "")
                        tarih = item.get("Tarih", "")
                        satici = item.get("Satıcı", "")
                        matrah = float(item.get("Matrah", 0.0) or 0.0)
                        kdv = float(item.get("KDV Tutarı", 0.0) or 0.0)
                        genel_toplam = float(item.get("Genel Toplam", 0.0) or (matrah + kdv))
                        
                        # 1. Satır: Matrah (Borç)
                        fis_satirlari.append({
                            "Fiş No": fis_sira_no,
                            "Tarih": tarih,
                            "Hesap Kodu": varsayilan_gider_kodu,
                            "Hesap Adı": "Gider / Mal Alış Hesabı",
                            "Açıklama": f"{satici} - Ftr No: {fatura_no}",
                            "Borç": matrah,
                            "Alacak": 0.0
                        })
                        
                        # 2. Satır: KDV (Borç)
                        if kdv > 0:
                            fis_satirlari.append({
                                "Fiş No": fis_sira_no,
                                "Tarih": tarih,
                                "Hesap Kodu": varsayilan_kdv_kodu,
                                "Hesap Adı": "İndirilecek KDV",
                                "Açıklama": f"{satici} - KDV",
                                "Borç": kdv,
                                "Alacak": 0.0
                            })
                        
                        # 3. Satır: Satıcı / Kasa (Alacak)
                        fis_satirlari.append({
                            "Fiş No": fis_sira_no,
                            "Tarih": tarih,
                            "Hesap Kodu": varsayilan_cari_kodu,
                            "Hesap Adı": satici,
                            "Açıklama": f"{satici} - Ftr No: {fatura_no}",
                            "Borç": 0.0,
                            "Alacak": genel_toplam
                        })
                        
                        fis_sira_no += 1
                    
                    df_sonuc = pd.DataFrame(fis_satirlari)
                else:
                    df_sonuc = pd.DataFrame(ham_veriler)

                st.divider()
                st.subheader("📊 Oluşturulan Muhasebe Fiş Tablosu")
                st.dataframe(df_sonuc, use_container_width=True)
                
                excel_cikti = excel_tablosu_olustur(df_sonuc, sheet_name="Fis_Aktarim")
                st.download_button(
                    label="📥 Muhasebe Aktarım Excel'ini İndir (.xlsx)",
                    data=excel_cikti,
                    file_name="muhasebe_fis_aktarim.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
