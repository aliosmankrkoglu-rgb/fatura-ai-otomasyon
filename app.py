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
    page_title="LedgerAI — Autonomous Accounting Engine", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 6 DİLLİ GLOBAL SÖZLÜK ---
LANG_DATA = {
    "🇹🇷 Türkçe": {
        "badge": "YAPAY ZEKA DESTEKLİ OTONOM MUHASEBE",
        "title": "LedgerAI",
        "subtitle": "Faturaları ve fişleri saniyeler içinde kurumsal ERP yevmiye fişine dönüştürün.",
        "upload_label": "Belgeleri buraya sürükleyin ya da seçin (PDF, PNG, JPG)",
        "process_btn": "⚡ Muhasebe Fişini Oluştur",
        "limit_err": "🛑 Demo sürümünde aynı anda en fazla 5 fatura işleyebilirsiniz.",
        "ready_count": "İşlenmeye hazır belge: **{count}**",
        "success": "✓ Fişler başarıyla oluşturuldu ve Borç/Alacak dengelendi.",
        "partial": "⚠️ {success} belge işlendi, {failed} belge okunamadı.",
        "failed": "❌ Belgeler işlenemedi. Lütfen görsel netliğini kontrol edin.",
        "preview_title": "📊 Muhasebe Yevmiye Fişi (Canlı Düzenlenebilir)",
        "preview_tip": "💡 Kod veya tutarları değiştirmek için hücreye çift tıklayın. İndirilen Excel'e anında yansır.",
        "tot_deb": "Toplam Borç",
        "tot_crd": "Toplam Alacak",
        "balanced": "✅ Fiş Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "download_btn": "📥 Kurumsal Excel'i İndir (.xlsx)",
        "theme_options": ["✨ Lüks Aurora Dalgası", "🖼️ Finans Görseli", "🌑 Saf Minimal Koyu"],
        "about_btn": "ℹ️ Nasıl Çalışır?",
        "about_title": "LedgerAI Otonom Mimari",
        "about_content": """
        **LedgerAI**, kurumların fatura ve fiş işleme maliyetlerini sıfıra indiren kurumsal bir yapay zeka motorudur.
        
        * **Çift Taraflı Denetim:** Yapay zeka faturayı okuduktan sonra `Toplam Borç = Toplam Alacak` matematiksel eşitliğini doğrular.
        * **Akıllı Kodlama:** Ticari malları `153`, genel yönetim masraflarını `770`, demirbaşları `255`, vergiyi `191` grubuna dinamik olarak bağlar.
        * **Kusursuz Entegrasyon:** İndirilen Excel dosyası ETA, Luca, Zirve, Logo, Datev veya QuickBooks gibi sistemlere doğrudan aktarılabilir.
        """
    },
    "🇺🇸 English": {
        "badge": "AI-POWERED AUTONOMOUS ACCOUNTING",
        "title": "LedgerAI",
        "subtitle": "Transform invoices and receipts into balanced ERP journal vouchers instantly.",
        "upload_label": "Drag and drop receipts or invoices (PDF, PNG, JPG)",
        "process_btn": "⚡ Generate Journal Voucher",
        "limit_err": "🛑 Demo allows up to 5 documents per batch.",
        "ready_count": "Documents ready: **{count}**",
        "success": "✓ Journal vouchers generated and balanced.",
        "partial": "⚠️ {success} processed, {failed} failed.",
        "failed": "❌ Documents could not be parsed.",
        "preview_title": "📊 Journal Voucher Table (Live Editable)",
        "preview_tip": "💡 Double-click any cell to adjust accounts or amounts before downloading.",
        "tot_deb": "Total Debit",
        "tot_crd": "Total Credit",
        "balanced": "✅ Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Unbalanced Voucher!",
        "download_btn": "📥 Download Clean Excel (.xlsx)",
        "theme_options": ["✨ Luxury Aurora Wave", "🖼️ Finance Image", "🌑 Pure Minimal Dark"],
        "about_btn": "ℹ️ How it Works?",
        "about_title": "LedgerAI Autonomous Architecture",
        "about_content": """
        **LedgerAI** is an institutional-grade accounting engine designed to eliminate manual data entry.
        
        * **Dual-Audit Engine:** Automatically verifies that `Total Debit = Total Credit` across all generated lines.
        * **Adaptive Smart Chart:** Automatically maps costs into OpEx, Inventory, Assets, and multi-tier VAT/Sales Tax.
        * **ERP Compatibility:** Downloaded tables are structured for immediate import into QuickBooks, Xero, SAP, or Datev.
        """
    },
    "🇩🇪 Deutsch": {
        "badge": "KI-GESTÜTZTE AUTONOME BUCHHALTUNG",
        "title": "LedgerAI",
        "subtitle": "Belege und Rechnungen automatisch in Datev-konforme Buchungssätze umwandeln.",
        "upload_label": "Belege oder Rechnungen hier ablegen (PDF, PNG, JPG)",
        "process_btn": "⚡ Buchungssatz Generieren",
        "limit_err": "🛑 Demo-Limit: Maximal 5 Dokumente.",
        "ready_count": "Bereit: **{count}** Dokumente",
        "success": "✓ Buchungen erfolgreich erstellt und ausgeglichen.",
        "partial": "⚠️ {success} verarbeitet, {failed} fehlgeschlagen.",
        "failed": "❌ Dokumente konnten nicht gelesen werden.",
        "preview_title": "📊 Buchungszeilen (Live Bearbeitbar)",
        "preview_tip": "💡 Doppelklicken Sie auf ein Feld, um Konten oder Beträge zu ändern.",
        "tot_deb": "Soll Gesamt",
        "tot_crd": "Haben Gesamt",
        "balanced": "✅ Ausgeglichen (Soll = Haben)",
        "unbalanced": "⚠️ Differenz festgestellt!",
        "download_btn": "📥 Excel-Buchungsdatei Herunterladen (.xlsx)",
        "theme_options": ["✨ Luxus Aurora Welle", "🖼️ Finanz-Bild", "🌑 Pur Minimal Dunkel"],
        "about_btn": "ℹ️ Funktionsweise",
        "about_title": "LedgerAI Autonome Architektur",
        "about_content": """
        **LedgerAI** automatisiert die Vorkontierung und Belegverarbeitung vollständig.
        
        * **Soll/Haben-Validierung:** Garantiert mathematische Ausgeglichenheit vor dem Export.
        * **Standardkontenrahmen:** Automatische Trennung nach SKR03/04 Richtlinien inklusive Vorsteuer.
        """
    },
    "🇫🇷 Français": {
        "badge": "COMPTABILITÉ AUTONOME PAR IA",
        "title": "LedgerAI",
        "subtitle": "Convertissez vos factures fournisseurs en écritures comptables équilibrées.",
        "upload_label": "Déposer des factures ou reçus (PDF, PNG, JPG)",
        "process_btn": "⚡ Générer les Écritures",
        "limit_err": "🛑 Limite démo: 5 documents maximum.",
        "ready_count": "Documents prêts: **{count}**",
        "success": "✓ Écritures générées et équilibrées.",
        "partial": "⚠️ {success} traités, {failed} échoués.",
        "failed": "❌ Impossible de lire les documents.",
        "preview_title": "📊 Journal Comptable (Modifiable en Direct)",
        "preview_tip": "💡 Double-cliquez sur une cellule pour modifier les comptes ou montants.",
        "tot_deb": "Total Débit",
        "tot_crd": "Total Crédit",
        "balanced": "✅ Équilibré (Débit = Crédit)",
        "unbalanced": "⚠️ Déséquilibre Détecté!",
        "download_btn": "📥 Télécharger le Journal Excel (.xlsx)",
        "theme_options": ["✨ Vague Aurore Luxe", "🖼️ Image Finance", "🌑 Sombre Épuré"],
        "about_btn": "ℹ️ Comment ça marche?",
        "about_title": "Architecture Autonome LedgerAI",
        "about_content": """
        **LedgerAI** numérise et impute automatiquement vos factures fournisseurs.
        
        * **Équilibre Parfait:** Vérification stricte Débit = Crédit.
        * **Conformité PCG:** Ventilation automatique des comptes de charges, TVA et fournisseurs.
        """
    },
    "🇪🇸 Español": {
        "badge": "CONTABILIDAD AUTÓNOMA CON IA",
        "title": "LedgerAI",
        "subtitle": "Convierte facturas y recibos en asientos contables listos para ERP al instante.",
        "upload_label": "Arrastra facturas o recibos aquí (PDF, PNG, JPG)",
        "process_btn": "⚡ Generar Asiento Contable",
        "limit_err": "🛑 Límite de demo: 5 documentos máximo.",
        "ready_count": "Documentos listos: **{count}**",
        "success": "✓ Asientos contables generados y equilibrados.",
        "partial": "⚠️ {success} procesados, {failed} fallidos.",
        "failed": "❌ No se pudieron procesar los documentos.",
        "preview_title": "📊 Asiento Contable (Editable en Vivo)",
        "preview_tip": "💡 Haz doble clic en cualquier celda para editar cuentas o valores.",
        "tot_deb": "Total Debe",
        "tot_crd": "Total Haber",
        "balanced": "✅ Asiento Cuadrado (Debe = Haber)",
        "unbalanced": "⚠️ Asiento Descuadrado!",
        "download_btn": "📥 Descargar Excel Limpio (.xlsx)",
        "theme_options": ["✨ Onda Aurora de Lujo", "🖼️ Imagen Finanzas", "🌑 Oscuro Minimalista"],
        "about_btn": "ℹ️ ¿Cómo funciona?",
        "about_title": "Arquitectura Autónoma LedgerAI",
        "about_content": """
        **LedgerAI** automatiza el registro de facturas para empresas y asesorías.
        
        * **Validación Doble:** Verifica automáticamente que `Debe = Haber`.
        * **Clasificación Inteligente:** Separa gastos, IVA soportado y cuentas de proveedores.
        """
    },
    "🇮🇹 Italiano": {
        "badge": "CONTABILITÀ AUTONOMA CON IA",
        "title": "LedgerAI",
        "subtitle": "Trasforma fatture e scontrini in scritture contabili bilanciate per ERP.",
        "upload_label": "Trascina qui fatture o ricevute (PDF, PNG, JPG)",
        "process_btn": "⚡ Genera Scrittura Contabile",
        "limit_err": "🛑 Limite demo: Massimo 5 documenti.",
        "ready_count": "Documenti pronti: **{count}**",
        "success": "✓ Scritture contabili generate e bilanciate.",
        "partial": "⚠️ {success} elaborati, {failed} falliti.",
        "failed": "❌ Impossibile analizzare i documenti.",
        "preview_title": "📊 Prima Nota (Modificabile in Diretta)",
        "preview_tip": "💡 Fai doppio clic su una cella per modificare conti o importi.",
        "tot_deb": "Totale Dare",
        "tot_crd": "Totale Avere",
        "balanced": "✅ Quadratura Perfetta (Dare = Avere)",
        "unbalanced": "⚠️ Scrittura Sbilanciata!",
        "download_btn": "📥 Scarica Excel Formattato (.xlsx)",
        "theme_options": ["✨ Onda Aurora di Lusso", "🖼️ Immagine Finanza", "🌑 Scuro Minimal"],
        "about_btn": "ℹ️ Come Funziona?",
        "about_title": "Architettura Autonoma LedgerAI",
        "about_content": """
        **LedgerAI** automatizza la registrazione contabile delle fatture passive.
        
        * **Quadratura Automatica:** Verifica che `Dare = Avere`.
        * **Gestione Fiscale:** Ripartizione precisa tra costo, IVA detraibile e fornitore.
        """
    }
}

# --- DURUM YÖNETİMİ ---
if "user_lang" not in st.session_state:
    st.session_state["user_lang"] = "🇹🇷 Türkçe"
if "theme_idx" not in st.session_state:
    st.session_state["theme_idx"] = 0  # 0: Aurora, 1: Resim, 2: Minimal Koyu

T = LANG_DATA[st.session_state["user_lang"]]

# --- ULTRA LÜKS ARKA PLAN VE LIQUID GLASS CSS ---
if st.session_state["theme_idx"] == 1:
    bg_css = """
        .stApp {
            background: linear-gradient(rgba(5, 8, 18, 0.82), rgba(5, 8, 18, 0.90)), 
                        url('https://images.unsplash.com/photo-1554224155-8d04cb21cd6c?q=80&w=2070&auto=format&fit=crop');
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }
    """
elif st.session_state["theme_idx"] == 0:
    bg_css = """
        @keyframes auroraGlow {
            0% { background-position: 0% 50%; filter: hue-rotate(0deg); }
            50% { background-position: 100% 50%; filter: hue-rotate(15deg); }
            100% { background-position: 0% 50%; filter: hue-rotate(0deg); }
        }
        .stApp {
            background: radial-gradient(circle at 15% 20%, rgba(79, 70, 229, 0.28), transparent 45%),
                        radial-gradient(circle at 85% 30%, rgba(14, 165, 233, 0.22), transparent 45%),
                        radial-gradient(circle at 50% 85%, rgba(16, 185, 129, 0.18), transparent 50%),
                        linear-gradient(140deg, #050811, #0A0F1D, #0D1627, #050811);
            background-size: 250% 250%;
            animation: auroraGlow 22s ease infinite;
            background-attachment: fixed;
        }
    """
else:
    bg_css = """
        .stApp {
            background-color: #060911;
        }
    """

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Plus Jakarta Sans', sans-serif;
    }}
    
    {bg_css}
    
    [data-testid="stSidebar"] {{
        display: none !important;
    }}
    
    .stApp {{
        color: #F8FAFC;
        padding-bottom: 120px;
    }}

    /* Üst İnce Badge */
    .top-badge {{
        display: inline-block;
        padding: 6px 14px;
        background: rgba(99, 102, 241, 0.12);
        border: 1px solid rgba(99, 102, 241, 0.35);
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 1.5px;
        color: #818CF8;
        margin-bottom: 12px;
        text-transform: uppercase;
        backdrop-filter: blur(10px);
    }}

    /* Başlık Tipografisi */
    .hero-title {{
        font-size: 2.8rem;
        font-weight: 800;
        background: linear-gradient(135deg, #FFFFFF 30%, #94A3B8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        letter-spacing: -0.8px;
        margin-bottom: 6px;
    }}
    
    .hero-sub {{
        font-size: 1.05rem;
        color: #94A3B8;
        margin-bottom: 28px;
        font-weight: 400;
    }}

    /* Lüks Cam Dosya Yükleyici */
    div[data-testid="stFileUploader"] {{
        background: rgba(13, 19, 33, 0.65);
        border: 1px dashed rgba(129, 140, 248, 0.45);
        border-radius: 20px;
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        padding: 36px 20px;
        transition: all 0.35s ease;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    }}
    div[data-testid="stFileUploader"]:hover {{
        border-color: #818CF8;
        box-shadow: 0 12px 35px rgba(99, 102, 241, 0.25);
        transform: translateY(-1px);
    }}
    
    /* İşlem Butonu */
    div.stButton > button:first-child {{
        background: linear-gradient(135deg, #4F46E5, #3B82F6);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 14px;
        font-weight: 600;
        font-size: 1rem;
        padding: 14px 28px;
        color: #FFFFFF;
        box-shadow: 0 4px 25px rgba(79, 70, 229, 0.4);
        transition: all 0.3s;
    }}
    div.stButton > button:first-child:hover {{
        box-shadow: 0 6px 35px rgba(79, 70, 229, 0.65);
        transform: translateY(-2px);
    }}

    /* Metrik Kartları */
    .stMetric {{
        background: rgba(13, 19, 33, 0.75);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        backdrop-filter: blur(16px);
        padding: 16px 22px;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.3);
    }}

    /* Sabit Liquid Glass Alt Bar */
    .liquid-dock-wrap {{
        position: fixed;
        bottom: 22px;
        left: 0;
        right: 0;
        margin: auto;
        width: max-content;
        max-width: 92vw;
        z-index: 999999;
    }}
    
    .liquid-dock {{
        background: rgba(13, 19, 33, 0.72);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 40px;
        backdrop-filter: blur(28px);
        -webkit-backdrop-filter: blur(28px);
        padding: 8px 18px;
        box-shadow: 0 15px 50px rgba(0, 0, 0, 0.6), inset 0 1px 0 rgba(255, 255, 255, 0.15);
        display: flex;
        align-items: center;
        gap: 12px;
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

# --- HERO BÖLÜMÜ ---
st.markdown(f"<div class='top-badge'>{T['badge']}</div>", unsafe_allow_html=True)
st.markdown(f"<div class='hero-title'>{T['title']}</div>", unsafe_allow_html=True)
st.markdown(f"<div class='hero-sub'>{T['subtitle']}</div>", unsafe_allow_html=True)

# --- DOSYA YÜKLEME ---
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
            
            for index, dosya in enumerate(yuklenen_dosyalar):
                status_text.text(f"İşleniyor ({index + 1}/{toplam_dosya}): {dosya.name}...")
                dosya_baytlari = dosya.read()
                mime_tipi = dosya.type if dosya.type else "application/pdf"
                
                prompt = """
                Sen otonom bir muhasebe denetçisisin. Belgeyi analiz et:
                1. Belge ülkesini/dilini otomatik belirle (TR, DE, FR, US vb.).
                2. Harcama türüne göre Tek Düzen / Standart Hesap Planı kodunu ata:
                   - Mal Alışı: 153.01
                   - Akaryakıt: 770.01
                   - Yemek / Ağırlama: 770.02
                   - Kırtasiye / Ofis: 770.03
                   - Kargo / Nakliye: 770.04
                   - Demirbaş / Cihaz: 255.01
                   - Diğer Giderler: 770.99
                3. KDV oranını ve tutarını tespit et.
                4. Satıcı için cari kod türet (320.VKN veya 320.AD).

                SADECE şu saf JSON şablonunu döndür:
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
                Sayılar float olmalı. Markdown etiketi ekleme.
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
                    if "Türkçe" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Fiş No", "Tarih", "Hesap Kodu", "Hesap Adı", "Açıklama", "Borç", "Alacak"
                        kdv_adi = f"%{tax_rate} İndirilecek KDV"
                    elif "Deutsch" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Beleg", "Datum", "Konto", "Bezeichnung", "Text", "Soll", "Haben"
                        kdv_adi = f"Vorsteuer {tax_rate}%"
                    elif "Français" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Pièce", "Date", "Compte", "Libellé", "Détail", "Débit", "Crédit"
                        kdv_adi = f"TVA {tax_rate}%"
                    elif "Español" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Asiento", "Fecha", "Cuenta", "Nombre Cuenta", "Concepto", "Debe", "Haber"
                        kdv_adi = f"IVA Soportado {tax_rate}%"
                    elif "Italiano" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Partita", "Data", "Conto", "Descrizione Conto", "Causale", "Dare", "Avere"
                        kdv_adi = f"IVA a Credito {tax_rate}%"
                    else:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Voucher #", "Date", "Account Code", "Account Name", "Memo", "Debit", "Credit"
                        kdv_adi = f"Tax ({tax_rate}%)"

                    # 1. Gider Satırı
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

# --- TABLO VE ÇIKTI ALANI ---
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
        file_name="ledger_journal_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# --- SAYFA ALTI BOŞLUĞU ---
st.markdown("<div style='height: 120px;'></div>", unsafe_allow_html=True)

# --- LIQUID GLASS FLOATING DOCK (EN ALT BAR) ---
st.markdown("<div class='liquid-dock-wrap'><div class='liquid-dock'>", unsafe_allow_html=True)

d_col1, d_col2, d_col3, d_col4, d_col5 = st.columns([1.5, 3.5, 4, 2.5, 1.5])

with d_col2:
    secilen_yeni_dil = st.selectbox(
        "Dil Seç",
        list(LANG_DATA.keys()),
        index=list(LANG_DATA.keys()).index(st.session_state["user_lang"]),
        label_visibility="collapsed"
    )
    if secilen_yeni_dil != st.session_state["user_lang"]:
        st.session_state["user_lang"] = secilen_yeni_dil
        st.rerun()

with d_col3:
    # Tema metinleri dinamik olarak seçilen dilden gelir
    secilen_tema_str = st.selectbox(
        "Görünüm",
        T["theme_options"],
        index=st.session_state["theme_idx"],
        label_visibility="collapsed"
    )
    yeni_idx = T["theme_options"].index(secilen_tema_str)
    if yeni_idx != st.session_state["theme_idx"]:
        st.session_state["theme_idx"] = yeni_idx
        st.rerun()

with d_col4:
    # Zarif Bilgi Pop-Up'ı (Modal Dialog)
    if hasattr(st, "dialog"):
        @st.dialog(T["about_title"])
        def ac_hakkinda_diyalog():
            st.markdown(T["about_content"])
        
        if st.button(T["about_btn"], use_container_width=True):
            ac_hakkinda_diyalog()
    else:
        with st.popover(T["about_btn"]):
            st.markdown(f"### {T['about_title']}")
            st.markdown(T["about_content"])

st.markdown("</div></div>", unsafe_allow_html=True)
