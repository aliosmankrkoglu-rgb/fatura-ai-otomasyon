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

# --- ÇOK DİLLİ GLOBAL SÖZLÜK ---
LANG_DATA = {
    "🇹🇷 TR": {
        "badge": "OTONOM MUHASEBE MOTORU",
        "title": "LedgerAI",
        "subtitle": "Belgeleri yükleyin; firma yapınıza ve ülkenize göre otomatik kodlanmış yevmiye fişini alın.",
        "upload_label": "Fatura veya fişleri sürükleyin ya da seçin (PDF, PNG, JPG)",
        "process_btn": "⚡ Fişi Oluştur",
        "limit_err": "🛑 Demo sürümünde en fazla 5 fatura işlenebilir.",
        "ready_count": "İşlenecek belge: **{count}**",
        "success": "✓ Fişler başarıyla oluşturuldu.",
        "failed": "❌ Belgeler işlenemedi.",
        "preview_title": "📊 Yevmiye Fişi (Canlı Düzenlenebilir)",
        "preview_tip": "💡 Hücrelere çift tıklayarak kod veya açıklamaları değiştirebilirsiniz.",
        "tot_deb": "Toplam Borç",
        "tot_crd": "Toplam Alacak",
        "balanced": "✅ Fiş Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "download_btn": "📥 Excel'i İndir (.xlsx)",
        "industry_label": "Firma Türü:",
        "industries": ["⚡ Otomatik (AI)", "🛒 Ticaret / Al-Sat (153 Ağırlıklı)", "🏢 Hizmet & Danışmanlık (770/740)", "🏭 Üretim & İmalat (150/730)"],
        "about_btn": "ℹ️ Nasıl Çalışır?",
        "about_title": "LedgerAI Mimarisi",
        "about_content": """
        **LedgerAI**, fatura ve fişleri doğrudan muhasebe programınızın aktarım formatına çevirir.
        
        * **Sektörel Mantık:** Firma türünüze göre ürünleri mal alışı (153), üretim hammaddesi (150) veya masraf (770) olarak dinamik ayırır.
        * **Çoklu Ülke Uyumu:** Belgenin diline ve para birimine göre (TR Tek Düzen, Almanya Datev, Fransa PCG, ABD GAAP) otonom çalışır.
        * **Kusursuz Bakiye:** Borç = Alacak eşitliğini kontrol etmeden aktarım vermez.
        """
    },
    "🇺🇸 EN": {
        "badge": "AUTONOMOUS ACCOUNTING ENGINE",
        "title": "LedgerAI",
        "subtitle": "Upload documents; get balanced ERP journal vouchers mapped to your business model.",
        "upload_label": "Drag and drop receipts or invoices (PDF, PNG, JPG)",
        "process_btn": "⚡ Process & Generate",
        "limit_err": "🛑 Demo allows up to 5 documents per batch.",
        "ready_count": "Documents: **{count}**",
        "success": "✓ Journal vouchers generated successfully.",
        "failed": "❌ Documents could not be parsed.",
        "preview_title": "📊 Journal Voucher (Live Editable)",
        "preview_tip": "💡 Double-click any cell to adjust accounts or amounts directly.",
        "tot_deb": "Total Debit",
        "tot_crd": "Total Credit",
        "balanced": "✅ Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Unbalanced Voucher!",
        "download_btn": "📥 Download Excel (.xlsx)",
        "industry_label": "Industry:",
        "industries": ["⚡ Auto (AI)", "🛒 Retail / Wholesale (Inventory)", "🏢 Services / SaaS (OpEx)", "🏭 Manufacturing (COGS)"],
        "about_btn": "ℹ️ How it Works?",
        "about_title": "LedgerAI Architecture",
        "about_content": """
        **LedgerAI** transforms incoming receipts directly into balanced journal vouchers.
        
        * **Business Context:** Differentiates inventory from operational expenses based on your business type.
        * **Multi-GAAP:** Supports US GAAP, German Datev (SKR03/04), French PCG, and Turkish standards.
        * **Zero Discrepancy:** Verifies Total Debit = Total Credit before generating export.
        """
    },
    "🇩🇪 DE": {
        "badge": "AUTONOME BUCHHALTUNGS-ENGINE",
        "title": "LedgerAI",
        "subtitle": "Belege automatisch erfassen und Datev-konform kontieren.",
        "upload_label": "Belege oder Rechnungen hier ablegen (PDF, PNG, JPG)",
        "process_btn": "⚡ Buchungssatz Generieren",
        "limit_err": "🛑 Maximal 5 Dokumente im Demo-Modus.",
        "ready_count": "Bereit: **{count}**",
        "success": "✓ Buchungen erfolgreich erstellt.",
        "failed": "❌ Belege konnten nicht gelesen werden.",
        "preview_title": "📊 Buchungszeilen (Bearbeitbar)",
        "preview_tip": "💡 Doppelklick auf Zellen zum Bearbeiten.",
        "tot_deb": "Soll Gesamt",
        "tot_crd": "Haben Gesamt",
        "balanced": "✅ Ausgeglichen (Soll = Haben)",
        "unbalanced": "⚠️ Differenz festgestellt!",
        "download_btn": "📥 Excel Herunterladen (.xlsx)",
        "industry_label": "Branche:",
        "industries": ["⚡ Automatisch (KI)", "🛒 Handel / Wareneinkauf", "🏢 Dienstleistung / IT", "🏭 Produktion / Fertigung"],
        "about_btn": "ℹ️ Info",
        "about_title": "LedgerAI Architektur",
        "about_content": "Vollautomatische Belegkontierung nach Datev SKR03/04 Richtlinien."
    },
    "🇫🇷 FR": {
        "badge": "MOTEUR COMPTABLE AUTONOME",
        "title": "LedgerAI",
        "subtitle": "Écritures comptables générées selon votre secteur d'activité.",
        "upload_label": "Déposer les factures (PDF, PNG, JPG)",
        "process_btn": "⚡ Générer les Écritures",
        "limit_err": "🛑 Limite: 5 documents maximum.",
        "ready_count": "Prêts: **{count}**",
        "success": "✓ Écritures générées avec succès.",
        "failed": "❌ Impossible de lire les documents.",
        "preview_title": "📊 Journal Comptable (Modifiable)",
        "preview_tip": "💡 Double-cliquez pour modifier une cellule.",
        "tot_deb": "Total Débit",
        "tot_crd": "Total Crédit",
        "balanced": "✅ Équilibré (Débit = Crédit)",
        "unbalanced": "⚠️ Déséquilibre Détecté!",
        "download_btn": "📥 Télécharger Excel (.xlsx)",
        "industry_label": "Secteur:",
        "industries": ["⚡ Auto (IA)", "🛒 Négoce / Vente", "🏢 Services / Conseil", "🏭 Production / Industrie"],
        "about_btn": "ℹ️ Info",
        "about_title": "Architecture LedgerAI",
        "about_content": "Génération d'écritures conforme au Plan Comptable Général (PCG)."
    }
}

# --- DURUM YÖNETİMİ ---
if "user_lang" not in st.session_state:
    st.session_state["user_lang"] = "🇹🇷 TR"
if "industry_choice" not in st.session_state:
    st.session_state["industry_choice"] = 0
if "theme_mode" not in st.session_state:
    st.session_state["theme_mode"] = "aurora"

T = LANG_DATA[st.session_state["user_lang"]]

# --- CSS VE HAFİF SÜZÜLEN IŞILTI DALGALARI ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    [data-testid="stSidebar"] { display: none !important; }
    
    /* Canlı, Derin ve Ultra Lüks Mesh Gradient */
    @keyframes floatLight {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    .stApp {
        background: radial-gradient(circle at 10% 20%, rgba(99, 102, 241, 0.22), transparent 40%),
                    radial-gradient(circle at 90% 30%, rgba(6, 182, 212, 0.18), transparent 45%),
                    radial-gradient(circle at 50% 80%, rgba(139, 92, 246, 0.15), transparent 50%),
                    linear-gradient(135deg, #05070E, #090E1A, #0C1527, #05070E);
        background-size: 250% 250%;
        animation: floatLight 20s ease infinite;
        background-attachment: fixed;
        color: #F8FAFC;
        padding-bottom: 90px;
    }

    /* Üst İnce Rozet */
    .top-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 12px;
        background: rgba(99, 102, 241, 0.12);
        border: 1px solid rgba(99, 102, 241, 0.3);
        border-radius: 99px;
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 1px;
        color: #A5B4FC;
        margin-bottom: 8px;
    }

    .hero-title {
        font-size: 2.4rem;
        font-weight: 800;
        letter-spacing: -0.8px;
        background: linear-gradient(135deg, #FFFFFF 40%, #94A3B8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 4px;
    }
    
    .hero-sub {
        font-size: 0.95rem;
        color: #94A3B8;
        margin-bottom: 22px;
    }

    /* Minimal Cam Dosya Yükleyici */
    div[data-testid="stFileUploader"] {
        background: rgba(13, 19, 33, 0.65);
        border: 1px dashed rgba(99, 102, 241, 0.4);
        border-radius: 16px;
        backdrop-filter: blur(16px);
        padding: 24px;
        transition: all 0.3s ease;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #818CF8;
        box-shadow: 0 8px 30px rgba(99, 102, 241, 0.2);
    }
    
    /* İşlem Butonu */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #4F46E5, #0EA5E9);
        border: none;
        border-radius: 10px;
        font-weight: 600;
        padding: 10px 22px;
        box-shadow: 0 4px 20px rgba(79, 70, 229, 0.35);
        transition: all 0.25s;
    }
    div.stButton > button:first-child:hover {
        box-shadow: 0 6px 28px rgba(79, 70, 229, 0.55);
        transform: translateY(-1px);
    }

    /* MİNİMAL LİKİT CAM ALT DOCK (YÜZEN BAR) */
    .dock-wrapper {
        position: fixed;
        bottom: 16px;
        left: 0;
        right: 0;
        margin: auto;
        width: fit-content;
        max-width: 90vw;
        z-index: 999999;
    }
    
    .dock-box {
        background: rgba(13, 19, 33, 0.78);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 50px;
        backdrop-filter: blur(24px);
        -webkit-backdrop-filter: blur(24px);
        padding: 4px 14px;
        box-shadow: 0 12px 35px rgba(0, 0, 0, 0.5);
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    /* Alt bardaki selectbox ve butonları mikro boyuta indirme */
    .dock-box div[data-testid="stSelectbox"] > div {
        min-height: 32px !important;
        height: 32px !important;
        font-size: 0.8rem !important;
        border-radius: 20px !important;
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
    }
    
    .dock-box div.stButton > button {
        height: 32px !important;
        padding: 4px 12px !important;
        font-size: 0.78rem !important;
        border-radius: 20px !important;
        background: rgba(255, 255, 255, 0.08) !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        box-shadow: none !important;
    }
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
            ws.column_dimensions[col_letter].width = max(m_len + 4, 12)

    return output.getvalue()

# --- HERO ALANI ---
st.markdown(f"<div class='top-badge'>● {T['badge']}</div>", unsafe_allow_html=True)
st.markdown(f"<div class='hero-title'>{T['title']}</div>", unsafe_allow_html=True)
st.markdown(f"<div class='hero-sub'>{T['subtitle']}</div>", unsafe_allow_html=True)

# --- YÜKLEME ALANI ---
yuklenen_dosyalar = st.file_uploader(
    T["upload_label"], 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True,
    label_visibility="collapsed"
)

if yuklenen_dosyalar:
    if len(yuklenen_dosyalar) > 5:
        st.error(T["limit_err"])
    else:
        st.write(T["ready_count"].format(count=len(yuklenen_dosyalar)))
        
        if st.button(T["process_btn"], use_container_width=True):
            ham_veriler = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            toplam_dosya = len(yuklenen_dosyalar)
            
            # Seçilen sektör mantığı direktifi
            sektor_secimi = T["industries"][st.session_state["industry_choice"]]
            sektor_direktifi = f"Firma Faaliyet Türü: {sektor_secimi}. "
            if "Ticaret" in sektor_secimi or "Retail" in sektor_secimi:
                sektor_direktifi += "Firma al-sat ticaret firmasıdır. Satışa konu olan ana ürünler '153.01 Ticari Mallar' (veya GAAP 1200 Inventory) hesabına işlenmelidir. Sadece akaryakıt, yemek, kırtasiye gibi şirket içi tüketimler 770'e gider."
            elif "Hizmet" in sektor_secimi or "Services" in sektor_secimi:
                sektor_direktifi += "Firma hizmet/ofis firmasıdır. Ürün alımları doğrudan işin maliyeti (740) veya genel gider (770) olarak kodlanmalıdır."
            elif "Üretim" in sektor_secimi or "Manufacturing" in sektor_secimi:
                sektor_direktifi += "Firma imalat firmasıdır. Hammadde ve malzeme alımları '150 İlk Madde Malzeme', fabrika giderleri '730', ofis giderleri '770' olarak kodlanmalıdır."
            else:
                sektor_direktifi += "Belgedeki kalemleri incele; ticari ürün ise 153, ofis/masraf ise 770, demirbaş ise 255'e mantıklı ata."

            for index, dosya in enumerate(yuklenen_dosyalar):
                status_text.text(f"İşleniyor ({index + 1}/{toplam_dosya}): {dosya.name}...")
                dosya_baytlari = dosya.read()
                mime_tipi = dosya.type if dosya.type else "application/pdf"
                
                prompt = f"""
                Sen kıdemli bir otonom mali müşavir ve ERP denetçisisin.
                {sektor_direktifi}
                
                Belge ülkesini (TR, DE, FR, US) ve para birimini otomatik tespit et.
                - Türkiye için Tek Düzen (153/150/770/740, 191 KDV, 320 Cari).
                - Almanya için Datev SKR03/04.
                - Fransa için PCG.
                - Global/ABD için US GAAP (1200 Inventory, 6000 OpEx, 2000 AP).

                SADECE şu JSON şablonunu döndür:
                {{
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
                }}
                Rakamlar float olmalıdır. Markdown etiketi ekleme.
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
                status_text.warning(f"✓ {len(ham_veriler)} / {toplam_dosya} işlendi.")
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
                    if "TR" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Fiş No", "Tarih", "Hesap Kodu", "Hesap Adı", "Açıklama", "Borç", "Alacak"
                        kdv_adi = f"%{tax_rate} İndirilecek KDV"
                    elif "DE" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Beleg", "Datum", "Konto", "Bezeichnung", "Text", "Soll", "Haben"
                        kdv_adi = f"Vorsteuer {tax_rate}%"
                    elif "FR" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Pièce", "Date", "Compte", "Libellé", "Détail", "Débit", "Crédit"
                        kdv_adi = f"TVA {tax_rate}%"
                    else:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Voucher #", "Date", "Account Code", "Account Name", "Memo", "Debit", "Credit"
                        kdv_adi = f"Tax ({tax_rate}%)"

                    # 1. Gider/Mal Satırı
                    fis_satirlari.append({
                        h_v: fis_no, h_d: date_val, h_c: acc_code, h_n: acc_name,
                        h_m: f"{vendor} - {inv_no}", "Para": curr, h_deb: net, h_crd: 0.0
                    })
                    
                    # 2. KDV Satırı
                    if tax > 0:
                        fis_satirlari.append({
                            h_v: fis_no, h_d: date_val, h_c: kdv_kod, h_n: kdv_adi,
                            h_m: f"{vendor} - KDV", "Para": curr, h_deb: tax, h_crd: 0.0
                        })
                    
                    # 3. Satıcı Satırı
                    fis_satirlari.append({
                        h_v: fis_no, h_d: date_val, h_c: cari_kod, h_n: vendor,
                        h_m: f"{vendor} - {inv_no}", "Para": curr, h_deb: 0.0, h_crd: total
                    })
                    
                    fis_no += 1

                st.session_state["out_df"] = pd.DataFrame(fis_satirlari)
                st.session_state["h_deb"] = h_deb
                st.session_state["h_crd"] = h_crd

# --- TABLO ALANI ---
if "out_df" in st.session_state:
    st.divider()
    st.subheader(T["preview_title"])
    st.caption(T["preview_tip"])
    
    guncel_df = st.data_editor(st.session_state["out_df"], use_container_width=True, num_rows="dynamic")
    
    deb_key = st.session_state["h_deb"]
    crd_key = st.session_state["h_crd"]
    
    tot_deb = guncel_df[deb_key].sum()
    tot_crd = guncel_df[crd_key].sum()
    
    c1, c2, c3 = st.columns(3)
    c1.metric(T["tot_deb"], f"{tot_deb:,.2f}")
    c2.metric(T["tot_crd"], f"{tot_crd:,.2f}")
    if abs(tot_deb - tot_crd) < 0.05:
        c3.success(T["balanced"])
    else:
        c3.warning(T["unbalanced"])
        
    excel_dosya = excel_olustur(guncel_df)
    st.download_button(
        label=T["download_btn"],
        data=excel_dosya,
        file_name="ledger_journal_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# --- MİKRO VE ŞIK LİKİT CAM ALT DOCK (FLOATING BAR) ---
st.markdown("<div style='height: 70px;'></div>", unsafe_allow_html=True)
st.markdown("<div class='dock-wrapper'><div class='dock-box'>", unsafe_allow_html=True)

col_d1, col_d2, col_d3 = st.columns([1.5, 3.5, 1.5])

with col_d1:
    yeni_dil = st.selectbox(
        "", 
        list(LANG_DATA.keys()), 
        index=list(LANG_DATA.keys()).index(st.session_state["user_lang"]),
        label_visibility="collapsed"
    )
    if yeni_dil != st.session_state["user_lang"]:
        st.session_state["user_lang"] = yeni_dil
        st.rerun()

with col_d2:
    secilen_sektor = st.selectbox(
        "", 
        T["industries"],
        index=st.session_state["industry_choice"],
        label_visibility="collapsed"
    )
    yeni_sektor_idx = T["industries"].index(secilen_sektor)
    if yeni_sektor_idx != st.session_state["industry_choice"]:
        st.session_state["industry_choice"] = yeni_sektor_idx
        st.rerun()

with col_d3:
    with st.popover(T["about_btn"]):
        st.markdown(f"#### {T['about_title']}")
        st.markdown(T["about_content"])

st.markdown("</div></div>", unsafe_allow_html=True)
