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
    page_title="LedgerAI - Akıllı Muhasebe Fiş Motoru", 
    page_icon="💼", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- DİL SÖZLÜĞÜ ---
DIL_SECENEKLERI = {
    "Türkçe 🇹🇷": {
        "title": "💼 LedgerAI — Otonom Muhasebe & Fiş Motoru",
        "subtitle": "Faturaları yükleyin; sistem harcama türünü (153, 770, 740 vb.) algılasın, KDV ayrımını yapsın ve dengeli yevmiye fişini çıkarsın.",
        "upload_label": "Fatura veya Fiş Yükleyin (PDF, PNG, JPG)",
        "process_btn": "🚀 Muhasebe Fişini Oluştur",
        "file_limit_err": "🛑 Ücretsiz demo sürümünde aynı anda en fazla 5 fatura işleyebilirsiniz.",
        "success_msg": "✓ Tüm fişler başarıyla oluşturuldu ve dengelendi.",
        "partial_msg": "⚠️ {success} belge işlendi, {failed} belge okunamadı.",
        "failed_msg": "❌ Belgeler işlenemedi. Lütfen dosya netliğini kontrol edin.",
        "preview_title": "📊 Muhasebe Yevmiye Fişi (Ekranda Düzenlenebilir)",
        "preview_tip": "💡 Hücrelere çift tıklayarak kod veya tutarları değiştirebilirsiniz. İndirilen Excel'e doğrudan yansır.",
        "total_deb": "Toplam Borç",
        "total_crd": "Toplam Alacak",
        "balanced": "✅ Fiş Bakiyesi Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "download_btn": "📥 Kurumsal Excel'i İndir (.xlsx)",
        "settings_title": "⚙️ Muhasebe & Görünüm Ayarları",
        "bg_toggle": "🖼️ Arka Plan Görselini Aç",
        "exp_header": "Hesap Grubu Tercihi (Borç)",
        "vat_header": "KDV Hesabı Tercihi",
        "ap_header": "Cari Hesap (320) Formatı",
        "ap_opt_vkn": "Vergi Numarası Bazlı (Örn: 320.1234567890)",
        "ap_opt_name": "Firma Adı Bazlı (Örn: 320.SHELL)",
        "ap_opt_std": "Standart Sıralı (Örn: 320.01.001)",
        "custom_plan_title": "📁 Özel Hesap Planı Yükle (Opsiyonel)"
    },
    "English 🇺🇸": {
        "title": "💼 LedgerAI — Autonomous Accounting Engine",
        "subtitle": "Upload invoices/receipts; AI auto-detects accounts (Inventory, OpEx, Assets), handles VAT/Tax, and balances journal vouchers.",
        "upload_label": "Upload Invoices or Receipts (PDF, PNG, JPG)",
        "process_btn": "🚀 Generate Journal Voucher",
        "file_limit_err": "🛑 Demo allows up to 5 documents per batch.",
        "success_msg": "✓ All journal vouchers successfully generated and balanced.",
        "partial_msg": "⚠️ {success} processed, {failed} failed.",
        "failed_msg": "❌ Files could not be parsed.",
        "preview_title": "📊 Journal Voucher Table (Editable)",
        "preview_tip": "💡 Double-click any cell to edit accounts or amounts directly.",
        "total_deb": "Total Debit",
        "total_crd": "Total Credit",
        "balanced": "✅ Voucher Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Unbalanced Voucher!",
        "download_btn": "📥 Download Formatted Excel (.xlsx)",
        "settings_title": "⚙️ Accounting & Theme Settings",
        "bg_toggle": "🖼️ Enable Background Wallpaper",
        "exp_header": "Primary Account Class (Debit)",
        "vat_header": "Tax / VAT Account",
        "ap_header": "Accounts Payable (AP) Format",
        "ap_opt_vkn": "Tax ID Based (e.g. AP-TAXID)",
        "ap_opt_name": "Vendor Name Based (e.g. AP-AMAZON)",
        "ap_opt_std": "Standard Sequential (e.g. 2000-01)",
        "custom_plan_title": "📁 Custom Chart of Accounts (Optional)"
    }
}

# --- YAN PANEL YAPILANDIRMASI ---
with st.sidebar:
    st.markdown("### 🌐 Dil / Language")
    secilen_dil = st.selectbox("", list(DIL_SECENEKLERI.keys()), label_visibility="collapsed")
    L = DIL_SECENEKLERI[secilen_dil]
    
    st.markdown("---")
    st.markdown(f"### {L['settings_title']}")
    
    # Arka plan görseli aç/kapa ayarı
    arka_plan_aktif = st.toggle(L["bg_toggle"], value=True)
    
    st.markdown("---")
    # Hesap Seçimleri (Anlaşılır ve Butonlu)
    st.markdown(f"**{L['exp_header']}**")
    ana_hesap_tercihi = st.selectbox(
        "",
        [
            "⚡ Otomatik (Yapay Zeka Karar Versin)",
            "153 - Ticari Mallar (Alış)",
            "770 - Genel Yönetim Giderleri",
            "740 - Hizmet Üretim Maliyeti",
            "760 - Pazarlama Satış Dağıtım",
            "255 - Demirbaşlar (Sabit Kıymet)",
            "730 - Genel Üretim Gideri"
        ],
        label_visibility="collapsed"
    )
    
    st.markdown(f"**{L['vat_header']}**")
    kdv_hesap_tercihi = st.radio(
        "",
        ["191 - İndirilecek KDV (Alışlar için)", "391 - Hesaplanan KDV (Satışlar için)"],
        label_visibility="collapsed"
    )
    
    st.markdown(f"**{L['ap_header']}**")
    cari_format = st.radio(
        "",
        [L["ap_opt_vkn"], L["ap_opt_name"], L["ap_opt_std"]],
        label_visibility="collapsed"
    )

    with st.expander(L["custom_plan_title"]):
        hesap_plani = st.file_uploader("Excel/CSV", type=["xlsx", "xls", "csv"], label_visibility="collapsed")
        hesap_ozeti = ""
        if hesap_plani:
            try:
                df_p = pd.read_csv(hesap_plani) if hesap_plani.name.endswith(".csv") else pd.read_excel(hesap_plani)
                c = df_p.columns[:2]
                hesap_ozeti = json.dumps(df_p[c].dropna().head(100).to_dict(orient="records"), ensure_ascii=False)
                st.success(f"✓ {len(df_p)} hesap listelendi")
            except:
                pass

# --- GÖRSEL TEMA VE CSS YÖNETİMİ ---
bg_css = """
    .stApp {
        background: linear-gradient(rgba(10, 15, 29, 0.88), rgba(10, 15, 29, 0.88)), 
                    url('https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?q=80&w=2070&auto=format&fit=crop');
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }
""" if arka_plan_aktif else """
    .stApp {
        background-color: #0B1120;
    }
"""

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    {bg_css}
    .stApp {{ color: #F8FAFC; }}
    div[data-testid="stFileUploader"] {{
        background: rgba(15, 23, 42, 0.75);
        border: 2px dashed #3B82F6;
        border-radius: 12px;
        backdrop-filter: blur(8px);
        padding: 24px;
    }}
    .stMetric {{
        background: rgba(15, 23, 42, 0.85);
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 12px 18px;
    }}
</style>
""", unsafe_allow_html=True)

# API İstemcisi
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# --- EXCEL OLUŞTURUCU FONKSİYON ---
def excel_olustur(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Fis_Aktarim')
        ws = writer.sheets['Fis_Aktarim']

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        body_font = Font(name="Calibri", size=10)
        border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )

        for col_idx in range(1, len(df.columns) + 1):
            c = ws.cell(row=1, column=col_idx)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
            m_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                cell.border = border
                if cell.row != 1:
                    cell.font = body_font
                    cell.alignment = Alignment(vertical="center")
            ws.column_dimensions[col_letter].width = max(m_len + 5, 14)

    return output.getvalue()

# --- ANA EKRAN ---
st.title(L["title"])
st.markdown(L["subtitle"])

yuklenen_dosyalar = st.file_uploader(
    L["upload_label"], 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    if len(yuklenen_dosyalar) > 5:
        st.error(L["file_limit_err"])
    else:
        st.write(f"📁 **{len(yuklenen_dosyalar)}** adet belge yüklendi.")
        
        if st.button(L["process_btn"], type="primary", use_container_width=True):
            ham_veriler = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            toplam_dosya = len(yuklenen_dosyalar)
            
            # KDV Tercihi
            kdv_koku = "191" if "191" in kdv_hesap_tercihi else "391"
            
            # Hesap Talimatı
            if "153" in ana_hesap_tercihi:
                hesap_talimati = "Kullanıcı tercihi gereği gider/mal hesabını kesinlikle '153.01 Ticari Mallar' olarak ata."
            elif "770" in ana_hesap_tercihi:
                hesap_talimati = "Kullanıcı tercihi gereği gider hesabını kesinlikle '770.01 Genel Yönetim Giderleri' olarak ata."
            elif "740" in ana_hesap_tercihi:
                hesap_talimati = "Kullanıcı tercihi gereği gider hesabını kesinlikle '740.01 Hizmet Üretim Maliyeti' olarak ata."
            elif "760" in ana_hesap_tercihi:
                hesap_talimati = "Kullanıcı tercihi gereği gider hesabını kesinlikle '760.01 Pazarlama Satış Dağıtım' olarak ata."
            elif "255" in ana_hesap_tercihi:
                hesap_talimati = "Kullanıcı tercihi gereği alımı kesinlikle '255.01 Demirbaşlar' hesabına ata."
            else:
                hesap_talimati = """
                Belgedeki ürünleri inceleyip Tek Düzen Hesap Planına göre mantıklı ata:
                - Satılacak ticari mal ise: 153.01
                - Akaryakıt, yemek, kırtasiye, genel ofis masrafı ise: 770 grubu (örn: 770.01, 770.02)
                - Bilgisayar, telefon, masa gibi demirbaş ise: 255.01
                - Nakliye, kargo ise: 770.04
                """

            if hesap_ozeti:
                hesap_talimati += f"\nÖNCELİK: Firmanın şu özel hesap planından uygun olanı seç: {hesap_ozeti}"

            for index, dosya in enumerate(yuklenen_dosyalar):
                status_text.text(f"İşleniyor ({index + 1}/{toplam_dosya}): {dosya.name}...")
                dosya_baytlari = dosya.read()
                mime_tipi = dosya.type if dosya.type else "application/pdf"
                
                prompt = f"""
                Sen uzman bir mali müşavirsin. Faturayı dikkatle oku.
                {hesap_talimati}

                SADECE şu JSON şablonunu döndür:
                {{
                  "fatura_no": "...",
                  "tarih": "YYYY-MM-DD",
                  "satici": "...",
                  "vkn": "...",
                  "hesap_kodu": "...",
                  "hesap_adi": "...",
                  "matrah": 0.0,
                  "kdv_orani": 20,
                  "kdv_tutari": 0.0,
                  "toplam": 0.0
                }}
                Tutar alanlarını sayısal (float) döndür. Markdown etiketi kullanma, doğrudan saf JSON dön.
                """
                
                maksimum_deneme = 3
                for deneme in range(maksimum_deneme):
                    try:
                        yanit = client.models.generate_content(
                            model="gemini-3.5-flash-lite",
                            contents=[types.Part.from_bytes(data=dosya_baytlari, mime_type=mime_tipi), prompt]
                        )
                        temiz = yanit.text.replace("```json", "").replace("```", "").strip()
                        veri = json.loads(temiz)
                        veri["dosya_adi"] = dosya.name
                        ham_veriler.append(veri)
                        break
                    except Exception as e:
                        hata_metni = str(e)
                        if ("503" in hata_metni or "429" in hata_metni) and deneme < maksimum_deneme - 1:
                            time.sleep(3 * (deneme + 1))
                            continue
                        else:
                            st.warning(f"⚠️ {dosya.name}: {hata_metni[:80]}")
                            break
                
                progress_bar.progress((index + 1) / toplam_dosya)
            
            # Durum mesajı
            if len(ham_veriler) == toplam_dosya:
                status_text.success(L["success_msg"])
            elif len(ham_veriler) > 0:
                status_text.warning(L["partial_msg"].format(success=len(ham_veriler), failed=toplam_dosya - len(ham_veriler)))
            else:
                status_text.error(L["failed_msg"])

            if ham_veriler:
                fis_satirlari = []
                fis_no = 1
                
                for item in ham_veriler:
                    f_no = str(item.get("fatura_no") or "").strip()
                    tarih = str(item.get("tarih") or "").strip()
                    satici = str(item.get("satici") or "Satıcı").strip()
                    vkn = str(item.get("vkn") or "").strip()
                    gider_kodu = str(item.get("hesap_kodu") or "770.01").strip()
                    hesap_adi = str(item.get("hesap_adi") or "Gider Hesabı").strip()
                    
                    matrah = float(item.get("matrah") or 0.0)
                    kdv = float(item.get("kdv_tutari") or 0.0)
                    toplam = float(item.get("toplam") or (matrah + kdv))
                    kdv_orani = item.get("kdv_orani") or 20
                    
                    temiz_ad = "".join(c for c in satici[:10] if c.isalnum()).upper() or "SATICI"
                    
                    # Cari Kod Belirleme
                    if L["ap_opt_vkn"] in cari_format and vkn:
                        cari_kodu = f"320.{vkn}"
                    elif L["ap_opt_name"] in cari_format:
                        cari_kodu = f"320.{temiz_ad}"
                    else:
                        cari_kodu = "320.01.001"

                    # 1. Gider/Mal Satırı (Borç)
                    fis_satirlari.append({
                        "Fiş No": fis_no, "Tarih": tarih, "Hesap Kodu": gider_kodu,
                        "Hesap Adı": hesap_adi, "Açıklama": f"{satici} - {f_no}",
                        "Borç": matrah, "Alacak": 0.0
                    })
                    
                    # 2. KDV Satırı (Borç)
                    if kdv > 0:
                        kdv_kod_tam = f"{kdv_koku}.{int(kdv_orani):02d}"
                        fis_satirlari.append({
                            "Fiş No": fis_no, "Tarih": tarih, "Hesap Kodu": kdv_kod_tam,
                            "Hesap Adı": f"%{kdv_orani} KDV", "Açıklama": f"{satici} - KDV",
                            "Borç": kdv, "Alacak": 0.0
                        })
                    
                    # 3. Satıcı / Cari Satırı (Alacak)
                    fis_satirlari.append({
                        "Fiş No": fis_no, "Tarih": tarih, "Hesap Kodu": cari_kodu,
                        "Hesap Adı": satici, "Açıklama": f"{satici} - {f_no}",
                        "Borç": 0.0, "Alacak": toplam
                    })
                    
                    fis_no += 1

                st.session_state["out_df"] = pd.DataFrame(fis_satirlari)

if "out_df" in st.session_state:
    st.divider()
    st.subheader(L["preview_title"])
    st.info(L["preview_tip"])
    
    guncel_df = st.data_editor(
        st.session_state["out_df"],
        use_container_width=True,
        num_rows="dynamic"
    )
    
    toplam_b = guncel_df["Borç"].sum()
    toplam_a = guncel_df["Alacak"].sum()
    
    col1, col2, col3 = st.columns(3)
    col1.metric(L["total_deb"], f"₺ {toplam_b:,.2f}")
    col2.metric(L["total_crd"], f"₺ {toplam_a:,.2f}")
    if abs(toplam_b - toplam_a) < 0.05:
        col3.success(L["balanced"])
    else:
        col3.warning(L["unbalanced"])
        
    excel_dosya = excel_olustur(guncel_df)
    st.download_button(
        label=L["download_btn"],
        data=excel_dosya,
        file_name="muhasebe_yevmiye_fisi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
