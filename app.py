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
    page_title="LedgerAI", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- DİL SÖZLÜKLERİ (BAYRAKLI & EKSİKSİZ) ---
LANG_DATA = {
    "🇹🇷 Türkçe": {
        "title": "⚡ LedgerAI",
        "subtitle": "Otonom Fatura & Fiş Muhasebeleştirme Motoru",
        "upload_label": "Fatura veya fişleri yükleyin (PDF, PNG, JPG)",
        "process_btn": "⚡ Muhasebe Fişini Oluştur",
        "limit_err": "🛑 Demo sürümünde aynı anda en fazla 5 fatura işleyebilirsiniz.",
        "ready_count": "İşlenecek belge sayısı: **{count}**",
        "success": "✓ Fişler başarıyla oluşturuldu ve Borç/Alacak dengelendi.",
        "partial": "⚠️ {success} belge işlendi, {failed} belge okunamadı.",
        "failed": "❌ Belgeler işlenemedi. Lütfen görsel netliğini kontrol edin.",
        "preview_title": "📊 Muhasebe Yevmiye Fişi (Düzenlenebilir)",
        "preview_tip": "💡 Kod veya tutarları değiştirmek için hücreye çift tıklayın. İndirilen Excel'e anında yansır.",
        "tot_deb": "Toplam Borç",
        "tot_crd": "Toplam Alacak",
        "balanced": "✅ Fiş Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "download_btn": "📥 Kurumsal Excel'i İndir (.xlsx)",
        "theme_header": "🎨 Arayüz Görünümü",
        "theme_label": "Arka Plan Stili",
        "opt_image": "🖼️ Finans Görseli",
        "opt_anim": "✨ Akıcı Animasyon",
        "opt_dark": "🌑 Minimal Koyu (Sade)"
    },
    "🇺🇸 English": {
        "title": "⚡ LedgerAI",
        "subtitle": "Autonomous Invoice & Receipt Accounting Engine",
        "upload_label": "Upload receipts or invoices (PDF, PNG, JPG)",
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
        "download_btn": "📥 Download Clean Excel (.xlsx)",
        "theme_header": "🎨 Interface Appearance",
        "theme_label": "Background Mode",
        "opt_image": "🖼️ Finance Image",
        "opt_anim": "✨ Fluid Animation",
        "opt_dark": "🌑 Minimal Dark (Clean)"
    },
    "🇩🇪 Deutsch": {
        "title": "⚡ LedgerAI",
        "subtitle": "Autonome Buchhaltungs- und Beleg-Engine",
        "upload_label": "Belege oder Rechnungen hochladen (PDF, PNG, JPG)",
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
        "download_btn": "📥 Excel-Buchungsdatei Herunterladen (.xlsx)",
        "theme_header": "🎨 Oberflächendesign",
        "theme_label": "Hintergrundmodus",
        "opt_image": "🖼️ Finanz-Bild",
        "opt_anim": "✨ Fluid-Animation",
        "opt_dark": "🌑 Minimal Dunkel (Schlicht)"
    },
    "🇫🇷 Français": {
        "title": "⚡ LedgerAI",
        "subtitle": "Moteur Autonome d'Écritures Comptables",
        "upload_label": "Déposer des factures ou reçus (PDF, PNG, JPG)",
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
        "download_btn": "📥 Télécharger le Journal Excel (.xlsx)",
        "theme_header": "🎨 Apparence",
        "theme_label": "Mode d'arrière-plan",
        "opt_image": "🖼️ Image Finance",
        "opt_anim": "✨ Animation Fluide",
        "opt_dark": "🌑 Sombre Épuré (Simple)"
    }
}

# --- YAN PANEL: DİL VE ARKA PLAN YÖNETİMİ ---
with st.sidebar:
    st.markdown("### 🌐 Dil / Language")
    secilen_dil = st.selectbox("", list(LANG_DATA.keys()), label_visibility="collapsed")
    T = LANG_DATA[secilen_dil]
    
    st.markdown("---")
    st.markdown(f"### {T['theme_header']}")
    arka_plan_modu = st.radio(
        T["theme_label"],
        [T["opt_image"], T["opt_anim"], T["opt_dark"]],
        label_visibility="collapsed"
    )

# --- CSS VE TEMA ENJEKSİYONU ---
if arka_plan_modu == T["opt_image"]:
    bg_style = """
        .stApp {
            background: linear-gradient(rgba(11, 15, 25, 0.88), rgba(11, 15, 25, 0.88)), 
                        url('https://images.unsplash.com/photo-1554224155-8d04cb21cd6c?q=80&w=2070&auto=format&fit=crop');
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }
    """
elif arka_plan_modu == T["opt_anim"]:
    bg_style = """
        @keyframes gradientBG {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        .stApp {
            background: linear-gradient(-45deg, #090D16, #0F172A, #1E1B4B, #090D16);
            background-size: 400% 400%;
            animation: gradientBG 15s ease infinite;
        }
    """
else:
    bg_style = """
        .stApp {
            background-color: #090D16;
        }
    """

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    {bg_style}
    .stApp {{ color: #F1F5F9; }}
    div[data-testid="stFileUploader"] {{
        background: rgba(15, 23, 42, 0.7);
        border: 2px dashed #3B82F6;
        border-radius: 12px;
        backdrop-filter: blur(10px);
        padding: 24px;
    }}
    .stMetric {{
        background: rgba(15, 23, 42, 0.8);
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

# --- ANA EKRAN ---
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
                2. Harcama türüne göre uygun hesap kodunu Tek Düzen / Standart Hesap Planına göre otomatik belirle:
                   - Ticari mal alımı ise: 153.01
                   - Akaryakıt: 770.01
                   - Yemek / Ağırlama: 770.02
                   - Kırtasiye / Ofis: 770.03
                   - Kargo / Nakliye: 770.04
                   - Demirbaş / Cihaz: 255.01
                   - Genel Masraf: 770.99
                3. KDV oranını ve tutarını doğru ayıkla.
                4. Satıcı için cari kod türet (Varsayılan: 320.VKN veya 320.AD).

                SADECE şu saf JSON objesini döndür:
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
                Sayısal alanları float dön. Markdown etiketi kullanma.
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
            
            # Sonuç Mesajı
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
                    doc_c = item.get("doc_country", "TR")
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
                    if "Türkçe" in secilen_dil:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Fiş No", "Tarih", "Hesap Kodu", "Hesap Adı", "Açıklama", "Borç", "Alacak"
                        kdv_adi = f"%{tax_rate} İndirilecek KDV"
                    elif "Deutsch" in secilen_dil:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Beleg", "Datum", "Konto", "Bezeichnung", "Text", "Soll", "Haben"
                        kdv_adi = f"Vorsteuer {tax_rate}%"
                    elif "Français" in secilen_dil:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Pièce", "Date", "Compte", "Libellé", "Détail", "Débit", "Crédit"
                        kdv_adi = f"TVA {tax_rate}%"
                    else:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Voucher #", "Date", "Account Code", "Account Name", "Memo", "Debit", "Credit"
                        kdv_adi = f"Tax ({tax_rate}%)"

                    # 1. Gider/Mal Satırı (Borç)
                    fis_satirlari.append({
                        h_v: fis_no, h_d: date_val, h_c: acc_code, h_n: acc_name,
                        h_m: f"{vendor} - {inv_no}", "Para Birimi": curr, h_deb: net, h_crd: 0.0
                    })
                    
                    # 2. KDV Satırı (Borç)
                    if tax > 0:
                        fis_satirlari.append({
                            h_v: fis_no, h_d: date_val, h_c: kdv_kod, h_n: kdv_adi,
                            h_m: f"{vendor} - KDV", "Para Birimi": curr, h_deb: tax, h_crd: 0.0
                        })
                    
                    # 3. Satıcı / Cari Satırı (Alacak)
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
