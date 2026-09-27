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

# --- YAN PANEL: ESNEK HESAP PLANI VE KURAL AYARLARI ---
with st.sidebar:
    st.title("⚙️ Hesap Planı & Format")
    
    st.markdown("### 1. Özel Hesap Planı Yükleme (Opsiyonel)")
    hesap_plani_dosyasi = st.file_uploader(
        "Firma Hesap Planı (Excel veya CSV)", 
        type=["xlsx", "xls", "csv"],
        help="Hesap Kodu ve Hesap Adı sütunlarını içeren dosyanızı yükleyin."
    )
    
    firma_hesap_ozeti = ""
    if hesap_plani_dosyasi:
        try:
            if hesap_plani_dosyasi.name.endswith(".csv"):
                df_plan = pd.read_csv(hesap_plani_dosyasi)
            else:
                df_plan = pd.read_excel(hesap_plani_dosyasi)
            
            # İlk 2 sütunu baz alarak hızlı özet çıkar
            cols = df_plan.columns[:2]
            ornek_kodlar = df_plan[cols].dropna().head(100).to_dict(orient="records")
            firma_hesap_ozeti = json.dumps(ornek_kodlar, ensure_ascii=False)
            st.success(f"✅ {len(df_plan)} satırlık hesap planı yüklendi!")
        except Exception as e:
            st.error(f"Hesap planı okunamadı: {str(e)[:60]}")
    
    st.markdown("---")
    st.markdown("### 2. Standart Format Ayarları")
    st.caption("Hesap planı yüklemediyseniz veya listede eşleşme yoksa bu kurallar uygulanır:")
    
    cari_kodlama_tipi = st.selectbox(
        "Cari (320) Kodlama Formatı",
        [
            "Vergi Numarası Bazlı (Örn: 320.VKN)", 
            "Firma Adı Bazlı (Örn: 320.SHELL)", 
            "Standart Sıralı (320.01.001)"
        ]
    )
    
    alt_hesap_stili = st.selectbox(
        "Gider Kodlama Formatı",
        ["3 Kademeli (770.01.001)", "2 Kademeli (770.01)", "Ana Hesap (770)"]
    )

# --- EXCEL OLUŞTURUCU (KURUMSAL FORMAT) ---
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
st.title("💼 FinAuto AI - Esnek Muhasebe Motoru")
st.markdown("Belgeleri yükleyin; yapay zeka firmanızın hesap planı yapısına göre Borç/Alacak yevmiye fişini çıkarsın.")

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
                
                ek_talimat = ""
                if firma_hesap_ozeti:
                    ek_talimat = f"ÖNCELİK: Firmanın özel hesap planı listesinden en uygun kodu seç: {firma_hesap_ozeti}"
                else:
                    ek_talimat = f"Format Tercihleri: Gider kodlarını {alt_hesap_stili} formatında aç. Cari hesap formatı: {cari_kodlama_tipi}."

                prompt = f"""
                Sen üst düzey bir muhasebe denetçisisin. Belgeyi incele ve bilgileri çıkar.
                {ek_talimat}

                SADECE şu JSON objesini döndür:
                {{
                  "Fatura No": "...",
                  "Tarih": "...",
                  "Satıcı": "...",
                  "VKN_TCKN": "...",
                  "Gider Türü": "Akaryakıt / Ticari Mal / Yemek / Kırtasiye / Diğer",
                  "Önerilen Hesap Kodu": "...",
                  "Hesap Adı": "...",
                  "Matrah": 0.0,
                  "KDV Orani": 20,
                  "KDV Tutarı": 0.0,
                  "Genel Toplam": 0.0
                }}
                Tutar alanlarını kesinlikle float döndür. Markdown etiketi kullanma, doğrudan saf JSON dön.
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
                    vkn = item.get("VKN_TCKN", "").strip()
                    gider_kodu = item.get("Önerilen Hesap Kodu", "770.01")
                    hesap_adi = item.get("Hesap Adı", "Genel Gider")
                    kdv_orani = item.get("KDV Orani", 20)
                    
                    matrah = float(item.get("Matrah", 0.0) or 0.0)
                    kdv = float(item.get("KDV Tutarı", 0.0) or 0.0)
                    genel_toplam = float(item.get("Genel Toplam", 0.0) or (matrah + kdv))
                    
                    # Cari Kod Belirleme
                    if "Vergi Numarası" in cari_kodlama_tipi and vkn:
                        cari_kod = f"320.{vkn}"
                    elif "Firma Adı" in cari_kodlama_tipi and satici:
                        temiz_isim = "".join(ch for ch in satici[:10] if ch.isalnum()).upper()
                        cari_kod = f"320.{temiz_isim}"
                    else:
                        cari_kod = "320.01.001"
                    
                    # 1. Gider/Mal Satırı (Borç)
                    fis_satirlari.append({
                        "Fiş No": fis_sira_no,
                        "Tarih": tarih,
                        "Hesap Kodu": gider_kodu,
                        "Hesap Adı": hesap_adi,
                        "Açıklama": f"{satici} - Ftr: {fatura_no}",
                        "Borç": matrah,
                        "Alacak": 0.0
                    })
                    
                    # 2. KDV Satırı (Borç)
                    if kdv > 0:
                        kdv_kodu = f"191.{int(kdv_orani):02d}" if kdv_orani else "191.20"
                        fis_satirlari.append({
                            "Fiş No": fis_sira_no,
                            "Tarih": tarih,
                            "Hesap Kodu": kdv_kodu,
                            "Hesap Adı": f"%{kdv_orani} İndirilecek KDV",
                            "Açıklama": f"{satici} - KDV",
                            "Borç": kdv,
                            "Alacak": 0.0
                        })
                    
                    # 3. Satıcı/Cari Satırı (Alacak)
                    fis_satirlari.append({
                        "Fiş No": fis_sira_no,
                        "Tarih": tarih,
                        "Hesap Kodu": cari_kod,
                        "Hesap Adı": satici,
                        "Açıklama": f"{satici} - Ftr: {fatura_no}",
                        "Borç": 0.0,
                        "Alacak": genel_toplam
                    })
                    
                    fis_sira_no += 1
                
                df_sonuc = pd.DataFrame(fis_satirlari)
                st.session_state["sonuc_tablosu"] = df_sonuc

if "sonuc_tablosu" in st.session_state:
    st.divider()
    st.subheader("📊 Oluşturulan Muhasebe Fiş Tablosu (Düzenlenebilir)")
    st.info("💡 Tablodaki herhangi bir hücreye çift tıklayarak kod veya açıklamaları manuel olarak değiştirebilirsiniz. Değişiklikler anında Excel çıktısına yansır.")
    
    # Ekranda interaktif Excel düzenleyici
    guncel_df = st.data_editor(
        st.session_state["sonuc_tablosu"], 
        use_container_width=True, 
        num_rows="dynamic"
    )
    
    toplam_borc = guncel_df["Borç"].sum()
    toplam_alacak = guncel_df["Alacak"].sum()
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Toplam Borç", f"{toplam_borc:,.2f} TL")
    col2.metric("Toplam Alacak", f"{toplam_alacak:,.2f} TL")
    if abs(toplam_borc - toplam_alacak) < 0.05:
        col3.success("✅ Fiş Dengeli (Borç = Alacak)")
    else:
        col3.warning("⚠️ Bakiye Farkı Var!")

    excel_cikti = excel_tablosu_olustur(guncel_df, sheet_name="Fis_Aktarim")
    st.download_button(
        label="📥 Düzenlenmiş Muhasebe Excel'ini İndir (.xlsx)",
        data=excel_cikti,
        file_name="akilli_muhasebe_fisi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
