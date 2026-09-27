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
    page_title="LedgerAI — Autonomous Accounting", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- DİL SÖZLÜKLERİ ---
LANG_DATA = {
    "🇹🇷 Türkçe": {
        "title": "⚡ LedgerAI",
        "subtitle": "Otonom Belge Okuma • Akıllı Hesap Eşleme • Yevmiye Fişi",
        "upload_label": "Faturaları veya fişleri buraya bırakın",
        "process_btn": "⚡ Muhasebe Fişini Oluştur",
        "limit_err": "🛑 Demo sürümünde aynı anda en fazla 5 fatura işleyebilirsiniz.",
        "ready_count": "İşlenecek belge sayısı: **{count}**",
        "success": "✓ Fişler başarıyla oluşturuldu ve Borç/Alacak dengelendi.",
        "partial": "⚠️ {success} belge işlendi, {failed} belge okunamadı.",
        "failed": "❌ Belgeler işlenemedi. Lütfen dosya netliğini kontrol edin.",
        "preview_title": "📊 Muhasebe Yevmiye Fişi (Ekranda Düzenlenebilir)",
        "preview_tip": "💡 Kod veya tutarları değiştirmek için hücreye çift tıklayın. İndirilen Excel'e anında yansır.",
        "tot_deb": "Toplam Borç",
        "tot_crd": "Toplam Alacak",
        "balanced": "✅ Fiş Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "download_btn": "📥 Kurumsal Excel'i İndir (.xlsx)"
    },
    "🇺🇸 English": {
        "title": "⚡ LedgerAI",
        "subtitle": "Autonomous Receipt Parsing • Smart Mapping • Balanced Journal",
        "upload_label": "Drop receipts or invoices here",
        "process_btn": "⚡ Generate Journal Voucher",
        "limit_err": "🛑 Demo allows up to 5 documents per batch.",
        "ready_count": "Documents ready: **{count}**",
        "success": "✓ Journal vouchers generated and balanced.",
        "partial": "⚠️ {success} processed, {failed} failed.",
        "failed": "❌ Documents could not be parsed.",
        "preview_title": "📊 Journal Voucher Table (Editable)",
        "preview_tip": "💡 Double-click any cell to edit accounts or values directly.",
        "tot_deb": "Total Debit",
        "tot_crd": "Total Credit",
        "balanced": "✅ Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Unbalanced Voucher!",
        "download_btn": "📥 Download Clean Excel (.xlsx)"
    },
    "🇩🇪 Deutsch": {
        "title": "⚡ LedgerAI",
        "subtitle": "Autonome Belegerfassung • Intelligente Kontierung • Buchungssätze",
        "upload_label": "Belege oder Rechnungen hier ablegen",
        "process_btn": "⚡ Buchungssatz Generieren",
        "limit_err": "🛑 Demo-Limit: Maximal 5 Dokumente.",
        "ready_count": "Bereit: **{count}** Dokumente",
        "success": "✓ Buchungen erfolgreich erstellt und ausgeglichen.",
        "partial": "⚠️ {success} verarbeitet, {failed} fehlgeschlagen.",
        "failed": "❌ Dokumente konnten nicht gelesen werden.",
        "preview_title": "📊 Buchungszeilen (Bearbeitbar)",
        "preview_tip": "💡 Doppelklicken Sie auf ein Feld, um Konten oder Beträge zu ändern.",
        "tot_deb": "Soll Gesamt",
        "tot_crd": "Haben Gesamt",
        "balanced": "✅ Ausgeglichen (Soll = Haben)",
        "unbalanced": "⚠️ Differenz festgestellt!",
        "download_btn": "📥 Excel-Buchungsdatei Herunterladen (.xlsx)"
    },
    "🇫🇷 Français": {
        "title": "⚡ LedgerAI",
        "subtitle": "Lecture Autonome • Imputation Intelligente • Écritures Comptables",
        "upload_label": "Déposez vos factures ou reçus ici",
        "process_btn": "⚡ Générer les Écritures",
        "limit_err": "🛑 Limite démo: 5 documents maximum.",
        "ready_count": "Documents prêts: **{count}**",
        "success": "✓ Écritures générées et équilibrées.",
        "partial": "⚠️ {success} traités, {failed} échoués.",
        "failed": "❌ Impossible de lire les documents.",
        "preview_title": "📊 Journal Comptable (Modifiable)",
        "preview_tip": "💡 Double-cliquez sur une cellule pour modifier les comptes ou montants.",
        "tot_deb": "Total Débit",
        "tot_crd": "Total Crédit",
        "balanced": "✅ Équilibré (Débit = Crédit)",
        "unbalanced": "⚠️ Déséquilibre Détecté!",
        "download_btn": "📥 Télécharger le Journal Excel (.xlsx)"
    }
}

# --- ALT YÜZEN BAR DURUM YÖNETİMİ ---
if "active_lang" not in st.session_state:
    st.session_state["active_lang"] = "🇹🇷 Türkçe"
if "active_theme" not in st.session_state:
    st.session_state["active_theme"] = "✨ Aurora Animasyon"

# --- LÜKS ARKA PLAN TEMA CSS ---
if st.session_state["active_theme"] == "🖼️ Finans Görseli":
    bg_style = """
        .stApp {
            background: linear-gradient(rgba(7, 11, 22, 0.85), rgba(7, 11, 22, 0.90)), 
                        url('https://images.unsplash.com/photo-1554224155-8d04cb21cd6c?q=80&w=2070&auto=format&fit=crop');
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }
    """
elif st.session_state["active_theme"] == "✨ Aurora Animasyon":
    bg_style = """
        @keyframes auroraWave {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        .stApp {
            background: radial-gradient(at 0% 0%, rgba(30, 27, 75, 0.8) 0px, transparent 50%),
                        radial-gradient(at 100% 0%, rgba(15, 23, 42, 0.9) 0px, transparent 50%),
                        radial-gradient(at 50% 100%, rgba(20, 83, 45, 0.3) 0px, transparent 50%),
                        linear-gradient(135deg, #070A12, #0E1726, #111C35, #070A12);
            background-size: 300% 300%;
            animation: auroraWave 18s ease infinite;
            background-attachment: fixed;
        }
    """
else:
    bg_style = """
        .stApp {
            background-color: #070A12;
        }
    """

# --- GLOBAL STİL VE YÜZEN ALT DOCK ---
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Plus Jakarta Sans', sans-serif;
    }}
    
    {bg_style}
    
    /* Yan paneli tamamen görünmez yap */
    [data-testid="stSidebar"] {{
        display: none !important;
    }}
    
    .stApp {{
        color: #F8FAFC;
        padding-bottom: 110px; /* Alt bar için boşluk */
    }}
    
    /* Modern Kartlar & Yükleme Alanı */
    div[data-testid="stFileUploader"] {{
        background: rgba(15, 23, 42, 0.6);
        border: 1px dashed rgba(99, 102, 241, 0.5);
        border-radius: 16px;
        backdrop-filter: blur(12px);
        padding: 30px;
        transition: all 0.3s ease;
    }}
    div[data-testid="stFileUploader"]:hover {{
        border-color: #6366F1;
        box-shadow: 0 0 25px rgba(99, 102, 241, 0.2);
    }}
    
    .stMetric {{
        background: rgba(15, 23, 42, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        backdrop-filter: blur(10px);
        padding: 14px 20px;
    }}
    
    /* Buton Tasarımı */
    div.stButton > button:first-child {{
        background: linear-gradient(135deg, #4F46E5, #3B82F6);
        border: none;
        border-radius: 12px;
        font-weight: 600;
        letter-spacing: 0.3px;
        padding: 12px 24px;
        box-shadow: 0 4px 20px rgba(79, 70, 229, 0.35);
        transition: all 0.3s;
    }}
    div.stButton > button:first-child:hover {{
        box-shadow: 0 6px 28px rgba(79, 70, 229, 0.55);
        transform: translateY(-1px);
    }}

    /* Alt Yüzen Cam Bar (Dock) */
    .dock-container {{
        position: fixed;
        bottom: 24px;
        left: 50%;
        transform: translateX(-50%);
        z-index: 999999;
        background: rgba(15, 23, 42, 0.75);
        border: 1px solid rgba(255, 255, 255, 0.15);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border-radius: 40px;
        padding: 8px 24px;
        box-shadow: 0 10px 40px rgba(0, 0, 0, 0.5), inset 0 1px 0 rgba(255, 255, 255, 0.1);
        display: flex;
        align-items: center;
        gap: 16px;
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
        df.to_excel(writer, index=False, sheet_name='Journal')
        ws = writer.sheets['Journal']

        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
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

# Aktif Dil Tanımı
T = LANG_DATA[st.session_state["active_lang"]]

# --- ANA MERKEZİ EKRAN ---
st.title(T["title"])
st.caption(T["subtitle"])

yuklenen_dosyalar = st.file_uploader(
    T["upload_label"], 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    if len(yuklenen_dosyalar) > 5:
        st.error(T["limit_err"])
    else:
        st.write(T["ready_count"].format(count=len(yuklenen_dosyalar)))
        
        if st.button(T["process_btn"], type="primary", use_container_width=True):
            ham_veriler = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            toplam_dosya = len(yuklenen_dosyalar)
            
            for index, dosya in enumerate(yuklenen_dosyalar):
                status_text.text(f"İşleniyor ({index + 1}/{toplam_dosya}): {dosya.name}...")
                dosya_baytlari = dosya.read()
                mime_tipi = dosya.type if dosya.type else "application/pdf"
                
                prompt = """
                Sen otonom bir muhasebe denetçisisin. Belgeyi incele:
                1. Belge ülkesini/dilini tespit et (TR, DE, FR, US).
                2. Harcama türüne göre Tek Düzen / GAAP kodunu otomatik ata:
                   - Ticari mal alımı: 153.01
                   - Akaryakıt: 770.01
                   - Yemek / Ağırlama: 770.02
                   - Kırtasiye / Ofis: 770.03
                   - Kargo / Nakliye: 770.04
                   - Demirbaş: 255.01
                   - Genel Masraf: 770.99
                3. KDV oranını ve tutarını doğru ayıkla.
                4. Satıcı için cari kod türet (320.VKN veya 320.AD).

                SADECE şu JSON objesini döndür:
                {
                  "doc_country": "TR",
                  "currency": "TL",
                  "invoice_no": "...",
                  "date": "YYYY-MM-DD",
                  "vendor": "...",
                  "tax_id": "...",
                  "account_code": "...",
                  "account_name": "...",
                  "net": 0.0,
                  "tax_rate": 20,
                  "tax": 0.0,
                  "total": 0.0
                }
                Sayısal alanlar float olmalı. Markdown etiketi ekleme.
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
                        hata_msg = str(e)
                        if ("503" in hata_msg or "429" in hata_msg) and deneme < maksimum_deneme - 1:
                            time.sleep(3 * (deneme + 1))
                            continue
                        else:
                            st.warning(f"⚠️ {dosya.name}: {hata_msg[:70]}")
                            break
                
                progress_bar.progress((index + 1) / toplam_dosya)
            
            if len(ham_veriler) == toplam_dosya:
                status_text.success(T["success"])
            elif len(ham_veriler) > 0:
                status_text.warning(T["partial"].format(success=len(ham_veriler), failed=toplam_dosya - len(ham_veriler)))
            else:
                status_text.error(T["failed"])

            if ham_veriler:
                fis_satirlari = []
                fis_no = 1
                
                for item in ham_veriler:
                    curr = item.get("currency", "TL")
                    inv_no = str(item.get("invoice_no") or "").strip()
                    date_val = str(item.get("date") or "").strip()
                    vendor = str(item.get("vendor") or "Satıcı").strip()
                    tax_id = str(item.get("tax_id") or "").strip()
                    acc_code = str(item.get("account_code") or "770.01").strip()
                    acc_name = str(item.get("account_name") or "Gider Hesabı").strip()
                    
                    net = float(item.get("net") or 0.0)
                    tax = float(item.get("tax") or 0.0)
                    total = float(item.get("total") or (net + tax))
                    tax_rate = item.get("tax_rate") or 20
                    
                    clean_name = "".join(c for c in vendor[:10] if c.isalnum()).upper() or "SATICI"
                    cari_kod = f"320.{tax_id}" if tax_id else f"320.{clean_name}"
                    kdv_kod = f"191.{int(tax_rate):02d}"

                    # Başlıklar
                    if "Türkçe" in st.session_state["active_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Fiş No", "Tarih", "Hesap Kodu", "Hesap Adı", "Açıklama", "Borç", "Alacak"
                        kdv_adi = f"%{tax_rate} İndirilecek KDV"
                    elif "Deutsch" in st.session_state["active_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Beleg", "Datum", "Konto", "Bezeichnung", "Text", "Soll", "Haben"
                        kdv_adi = f"Vorsteuer {tax_rate}%"
                    elif "Français" in st.session_state["active_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Pièce", "Date", "Compte", "Libellé", "Détail", "Débit", "Crédit"
                        kdv_adi = f"TVA {tax_rate}%"
                    else:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Voucher #", "Date", "Account Code", "Account Name", "Memo", "Debit", "Credit"
                        kdv_adi = f"Tax ({tax_rate}%)"

                    # 1. Gider/Mal Satırı
                    fis_satirlari.append({
                        h_v: fis_no, h_d: date_val, h_c: acc_code, h_n: acc_name,
                        h_m: f"{vendor} - {inv_no}", "Para Birimi": curr, h_deb: net, h_crd: 0.0
                    })
                    
                    # 2. KDV Satırı
                    if tax > 0:
                        fis_satirlari.append({
                            h_v: fis_no, h_d: date_val, h_c: kdv_kod, h_n: kdv_adi,
                            h_m: f"{vendor} - KDV", "Para Birimi": curr, h_deb: tax, h_crd: 0.0
                        })
                    
                    # 3. Satıcı Satırı
                    fis_satirlari.append({
                        h_v: fis_no, h_d: date_val, h_c: cari_kod, h_n: vendor,
                        h_m: f"{vendor} - {inv_no}", "Para Birimi": curr, h_deb: 0.0, h_crd: total
                    })
                    
                    fis_no += 1

                st.session_state["out_df"] = pd.DataFrame(fis_satirlari)
                st.session_state["h_deb"] = h_deb
                st.session_state["h_crd"] = h_crd

if "out_df" in st.session_state:
    st.divider()
    st.subheader(T["preview_title"])
    st.info(T["preview_tip"])
    
    guncel_df = st.data_editor(
        st.session_state["out_df"],
        use_container_width=True,
        num_rows="dynamic"
    )
    
    deb_key = st.session_state["h_deb"]
    crd_key = st.session_state["h_crd"]
    
    tot_deb = guncel_df[deb_key].sum()
    tot_crd = guncel_df[crd_key].sum()
    
    col1, col2, col3 = st.columns(3)
    col1.metric(T["tot_deb"], f"{tot_deb:,.2f}")
    col2.metric(T["tot_crd"], f"{tot_crd:,.2f}")
    if abs(tot_deb - tot_crd) < 0.05:
        col3.success(T["balanced"])
    else:
        col3.warning(T["unbalanced"])
        
    excel_dosya = excel_olustur(guncel_df)
    st.download_button(
        label=T["download_btn"],
        data=excel_dosya,
        file_name="muhasebe_yevmiye_fisi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# --- ALT YÜZEN KONTROL DOCK'U (FLOATING GLASS DOCK) ---
st.markdown("<div style='height: 80px;'></div>", unsafe_allow_html=True)

# Alt barı sayfaya sabitleyen kapsayıcı kolonlar
dock_col1, dock_col2, dock_col3, dock_col4 = st.columns([2, 3, 3, 2])

with dock_col2:
    yeni_dil = st.selectbox(
        "🌐 Dil",
        list(LANG_DATA.keys()),
        index=list(LANG_DATA.keys()).index(st.session_state["active_lang"]),
        label_visibility="collapsed"
    )
    if yeni_dil != st.session_state["active_lang"]:
        st.session_state["active_lang"] = yeni_dil
        st.rerun()

with dock_col3:
    yeni_tema = st.selectbox(
        "🎨 Görünüm",
        ["✨ Aurora Animasyon", "🖼️ Finans Görseli", "🌑 Minimal Koyu"],
        index=["✨ Aurora Animasyon", "🖼️ Finans Görseli", "🌑 Minimal Koyu"].index(st.session_state["active_theme"]),
        label_visibility="collapsed"
    )
    if yeni_tema != st.session_state["active_theme"]:
        st.session_state["active_theme"] = yeni_tema
        st.rerun()
