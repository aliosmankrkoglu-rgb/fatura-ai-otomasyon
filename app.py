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
    page_title="LedgerAI — Institutional Autonomous Accounting", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 6 DİLLİ VE KUSURSUZ ÇEVİRİ MERKEZİ ---
LANG_DATA = {
    "🇹🇷 TR": {
        "badge": "INSTITUTIONAL AI ENGINE",
        "title": "LedgerAI",
        "subtitle": "Otonom Belge Denetimi • Sektörel Kodlama • ERP Yevmiye Fişi",
        "upload_label": "Faturaları buraya sürükleyin (PDF, PNG, JPG)",
        "upload_hint": "Kurumsal ERP aktarımına hazır yevmiye fişleri otomatik üretilir",
        "process_btn": "⚡ Fişleri Otonom Muhasebeleştir",
        "limit_err": "🛑 Demo sürümünde oturum başına en fazla 5 belge işlenebilir.",
        "ready_count": "İşlenecek belge sayısı: **{count}**",
        "success": "✓ Fişler başarıyla oluşturuldu ve Borç/Alacak dengelendi.",
        "failed": "❌ Belgeler işlenemedi. Lütfen görsel netliğini kontrol edin.",
        "preview_title": "📊 Muhasebe Yevmiye Fişi (Canlı Hücre Düzenleme)",
        "preview_tip": "💡 Hücrelere çift tıklayarak kod veya açıklamaları değiştirebilirsiniz. İndirilen Excel'e anında yansır.",
        "tot_deb": "Toplam Borç",
        "tot_crd": "Toplam Alacak",
        "balanced": "✅ Fiş Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "download_btn": "📥 Kurumsal Excel'i İndir (.xlsx)",
        "industries": ["⚡ Otomatik Sektör (AI)", "🛒 Ticaret / Al-Sat (153 Ağırlıklı)", "🏢 Hizmet & Ofis (770/740)", "🏭 Üretim & Fabrika (150/730)"],
        "themes": ["✨ Ultra Aurora Mesh", "🖼️ Finans Görseli", "🌑 Saf Titanyum Koyu"],
        "about_btn": "ℹ️ Mimari",
        "about_title": "LedgerAI Kurumsal Mimarisi",
        "about_content": """
        **LedgerAI**, kurumların ön muhasebe ve fiş giriş maliyetlerini sıfıra indiren yeni nesil finans motorudur.
        
        * **Sektörel Zeka:** Faturadaki mal alımını şirketin yapısına göre (Ticarette 153, Üretimde 150, Hizmette 770/740) dinamik eşler.
        * **Kusursuz Bakiye:** Borç = Alacak matematiksel denetimini kuruşu kuruşuna yapar.
        * **Doğrudan Entegrasyon:** ETA, Luca, Zirve, Logo, Datev ve QuickBooks sistemlerine hazır Excel üretir.
        """
    },
    "🇺🇸 EN": {
        "badge": "INSTITUTIONAL AI ENGINE",
        "title": "LedgerAI",
        "subtitle": "Autonomous Document Audit • Industry Mapping • Balanced ERP Vouchers",
        "upload_label": "Drag and drop receipts or invoices (PDF, PNG, JPG)",
        "upload_hint": "Instant generation of balanced, ERP-ready journal vouchers",
        "process_btn": "⚡ Generate Balanced Vouchers",
        "limit_err": "🛑 Demo allows up to 5 documents per batch.",
        "ready_count": "Ready to audit: **{count}**",
        "success": "✓ Journal vouchers generated and balanced.",
        "failed": "❌ Documents could not be processed.",
        "preview_title": "📊 Journal Voucher Grid (Live Editable)",
        "preview_tip": "💡 Double-click any cell to adjust accounts or descriptions before export.",
        "tot_deb": "Total Debit",
        "tot_crd": "Total Credit",
        "balanced": "✅ Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Unbalanced Voucher!",
        "download_btn": "📥 Download Clean Excel (.xlsx)",
        "industries": ["⚡ Auto Industry (AI)", "🛒 Retail / Inventory (1200)", "🏢 Services / SaaS (OpEx)", "🏭 Manufacturing (COGS)"],
        "themes": ["✨ Ultra Aurora Mesh", "🖼️ Finance Image", "🌑 Pure Titanium Dark"],
        "about_btn": "ℹ️ Architecture",
        "about_title": "LedgerAI Core Engine",
        "about_content": """
        **LedgerAI** eliminates manual bookkeeping for modern enterprises.
        
        * **Contextual Mapping:** Distinguishes inventory from operating expenses based on entity profile.
        * **Dual Verification:** Guarantees Debit = Credit parity before release.
        * **Global Formats:** Direct compatibility with QuickBooks, Xero, NetSuite, SAP, and Datev.
        """
    },
    "🇩🇪 DE": {
        "badge": "INSTITUTIONELLE KI-ENGINE",
        "title": "LedgerAI",
        "subtitle": "Autonome Belegerfassung • Branchenspezifische Kontierung • Datev Export",
        "upload_label": "Belege oder Rechnungen ablegen (PDF, PNG, JPG)",
        "upload_hint": "Datev- und ERP-konforme Buchungssätze in Echtzeit",
        "process_btn": "⚡ Buchungssätze Erstellen",
        "limit_err": "🛑 Maximal 5 Dokumente im Demo-Modus.",
        "ready_count": "Bereit: **{count}**",
        "success": "✓ Buchungen erfolgreich erstellt.",
        "failed": "❌ Dokumente konnten nicht verarbeitet werden.",
        "preview_title": "📊 Buchungszeilen (Live Bearbeitbar)",
        "preview_tip": "💡 Doppelklick zum Anpassen von Konten oder Beträgen.",
        "tot_deb": "Soll Gesamt",
        "tot_crd": "Haben Gesamt",
        "balanced": "✅ Ausgeglichen (Soll = Haben)",
        "unbalanced": "⚠️ Differenz festgestellt!",
        "download_btn": "📥 Excel Herunterladen (.xlsx)",
        "industries": ["⚡ Automatisch (KI)", "🛒 Handel / Wareneinkauf", "🏢 Dienstleistung / IT", "🏭 Produktion / Fertigung"],
        "themes": ["✨ Ultra Aurora Mesh", "🖼️ Finanz-Bild", "🌑 Reines Titan Dunkel"],
        "about_btn": "ℹ️ Architektur",
        "about_title": "LedgerAI Architektur",
        "about_content": "Vollautomatische Belegkontierung nach Datev SKR03/04 Richtlinien."
    },
    "🇫🇷 FR": {
        "badge": "MOTEUR COMPTABLE INSTITUTIONNEL",
        "title": "LedgerAI",
        "subtitle": "Audit Documentaire • Imputation Sectorielle • Journal ERP",
        "upload_label": "Déposer vos pièces comptables (PDF, PNG, JPG)",
        "upload_hint": "Écritures équilibrées prêtes pour votre logiciel comptable",
        "process_btn": "⚡ Générer les Écritures",
        "limit_err": "🛑 Limite: 5 documents par lot.",
        "ready_count": "Prêts: **{count}**",
        "success": "✓ Écritures générées avec succès.",
        "failed": "❌ Échec du traitement.",
        "preview_title": "📊 Journal Comptable (Édition Directe)",
        "preview_tip": "💡 Double-cliquez pour ajuster les comptes ou libellés.",
        "tot_deb": "Total Débit",
        "tot_crd": "Total Crédit",
        "balanced": "✅ Équilibré (Débit = Crédit)",
        "unbalanced": "⚠️ Déséquilibre Détecté!",
        "download_btn": "📥 Télécharger Excel (.xlsx)",
        "industries": ["⚡ Auto (IA)", "🛒 Négoce / Stock", "🏢 Services / Conseil", "🏭 Production / Industrie"],
        "themes": ["✨ Ultra Aurora Mesh", "🖼️ Image Finance", "🌑 Titane Pur Sombre"],
        "about_btn": "ℹ️ Architecture",
        "about_title": "Architecture LedgerAI",
        "about_content": "Génération conforme au Plan Comptable Général (PCG)."
    },
    "🇪🇸 ES": {
        "badge": "MOTOR CONTABLE INSTITUCIONAL",
        "title": "LedgerAI",
        "subtitle": "Auditoría Autónoma • Clasificación Sectorial • Asientos ERP",
        "upload_label": "Arrastra facturas o recibos aquí (PDF, PNG, JPG)",
        "upload_hint": "Asientos contables cuadrados listos para importar",
        "process_btn": "⚡ Generar Asientos",
        "limit_err": "🛑 Máximo 5 documentos por lote.",
        "ready_count": "Documentos: **{count}**",
        "success": "✓ Asientos generados y equilibrados.",
        "failed": "❌ Error al procesar documentos.",
        "preview_title": "📊 Libro Diario (Editable)",
        "preview_tip": "💡 Doble clic en cualquier celda para editar.",
        "tot_deb": "Total Debe",
        "tot_crd": "Total Haber",
        "balanced": "✅ Asiento Cuadrado (Debe = Haber)",
        "unbalanced": "⚠️ Asiento Descuadrado!",
        "download_btn": "📥 Descargar Excel (.xlsx)",
        "industries": ["⚡ Automático (IA)", "🛒 Comercio / Inventario", "🏢 Servicios / Oficina", "🏭 Fabricación / Industria"],
        "themes": ["✨ Ultra Aurora Mesh", "🖼️ Imagen Finanzas", "🌑 Titanio Puro Oscuro"],
        "about_btn": "ℹ️ Arquitectura",
        "about_title": "Arquitectura LedgerAI",
        "about_content": "Automatización contable con validación estricta Debe = Haber."
    },
    "🇮🇹 IT": {
        "badge": "MOTORE CONTABILE ISTITUZIONALE",
        "title": "LedgerAI",
        "subtitle": "Controllo Documenti • Riconciliazione Settoriale • Prima Nota ERP",
        "upload_label": "Trascina fatture o ricevute qui (PDF, PNG, JPG)",
        "upload_hint": "Scritture contabili bilanciate pronte per il gestionale",
        "process_btn": "⚡ Genera Scritture",
        "limit_err": "🛑 Limite demo: 5 documenti per sessione.",
        "ready_count": "Documenti pronti: **{count}**",
        "success": "✓ Scritture generate e bilanciate.",
        "failed": "❌ Impossibile elaborare i documenti.",
        "preview_title": "📊 Prima Nota (Modifica Diretta)",
        "preview_tip": "💡 Fai doppio clic su una cella per modificare.",
        "tot_deb": "Totale Dare",
        "tot_crd": "Totale Avere",
        "balanced": "✅ Quadratura Perfetta (Dare = Avere)",
        "unbalanced": "⚠️ Sbilancio Rilevato!",
        "download_btn": "📥 Scarica Excel (.xlsx)",
        "industries": ["⚡ Automatico (IA)", "🛒 Commercio / Magazzino", "🏢 Servizi / Consulenza", "🏭 Manifattura / Produzione"],
        "themes": ["✨ Ultra Aurora Mesh", "🖼️ Immagine Finanza", "🌑 Titanio Puro Scuro"],
        "about_btn": "ℹ️ Architettura",
        "about_title": "Architettura LedgerAI",
        "about_content": "Generazione automatizzata di prima nota con quadratura fiscale."
    }
}

# --- GÜVENLİ DURUM YÖNETİMİ ---
if "user_lang" not in st.session_state or st.session_state["user_lang"] not in LANG_DATA:
    st.session_state["user_lang"] = "🇹🇷 TR"
if "theme_idx" not in st.session_state:
    st.session_state["theme_idx"] = 0
if "industry_idx" not in st.session_state:
    st.session_state["industry_idx"] = 0

T = LANG_DATA[st.session_state["user_lang"]]

# --- REVOLUT ULTRA x DYNAMIC MESH CSS ---
if st.session_state["theme_idx"] == 1:
    # 🖼️ Lüks Metalik Finans Arka Planı
    bg_css = """
        .stApp {
            background: linear-gradient(rgba(7, 10, 19, 0.84), rgba(7, 10, 19, 0.92)), 
                        url('https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?q=80&w=2070&auto=format&fit=crop');
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }
    """
elif st.session_state["theme_idx"] == 0:
    # ✨ Ultra Canlı Aurora Mesh (Revolut Ultra & Linear Havası)
    bg_css = """
        @keyframes dynamicMesh {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        .stApp {
            background: radial-gradient(circle at 10% 15%, rgba(99, 102, 241, 0.28), transparent 35%),
                        radial-gradient(circle at 88% 25%, rgba(6, 182, 212, 0.22), transparent 40%),
                        radial-gradient(circle at 50% 85%, rgba(139, 92, 246, 0.18), transparent 45%),
                        linear-gradient(135deg, #05070D, #080D18, #0B1426, #05070D);
            background-size: 260% 260%;
            animation: dynamicMesh 22s ease infinite;
            background-attachment: fixed;
        }
    """
else:
    # 🌑 Saf Titanyum Minimal Koyu
    bg_css = """
        .stApp {
            background-color: #05070D;
        }
    """

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Plus Jakarta Sans', sans-serif;
    }}
    
    [data-testid="stSidebar"] {{ display: none !important; }}
    
    {bg_css}
    
    .stApp {{
        color: #F8FAFC;
        padding-bottom: 95px;
    }}

    /* Lüks Kompakt Hero */
    .hero-container {{
        text-align: center;
        padding: 10px 0 20px 0;
        max-width: 650px;
        margin: 0 auto;
    }}

    .top-badge {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 12px;
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 99px;
        font-size: 0.7rem;
        font-weight: 700;
        letter-spacing: 1.2px;
        color: #94A3B8;
        margin-bottom: 10px;
        backdrop-filter: blur(12px);
    }}

    .hero-title {{
        font-size: 2.3rem;
        font-weight: 800;
        letter-spacing: -0.6px;
        background: linear-gradient(135deg, #FFFFFF 40%, #A1A1AA 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 4px;
        line-height: 1.1;
    }}
    
    .hero-sub {{
        font-size: 0.92rem;
        color: #94A3B8;
        font-weight: 400;
        line-height: 1.4;
    }}

    /* Cam Dosya Yükleme Paneli */
    div[data-testid="stFileUploader"] {{
        background: rgba(13, 18, 30, 0.65);
        border: 1px dashed rgba(255, 255, 255, 0.18);
        border-radius: 16px;
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        padding: 24px;
        transition: all 0.3s ease;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    }}
    div[data-testid="stFileUploader"]:hover {{
        border-color: rgba(99, 102, 241, 0.6);
        box-shadow: 0 12px 35px rgba(99, 102, 241, 0.2);
    }}

    /* Titan / Neon Buton */
    div.stButton > button:first-child {{
        background: linear-gradient(135deg, #4F46E5, #0284C7);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 10px;
        font-weight: 600;
        font-size: 0.92rem;
        padding: 10px 22px;
        color: #FFFFFF;
        box-shadow: 0 4px 20px rgba(79, 70, 229, 0.35);
        transition: all 0.25s;
    }}
    div.stButton > button:first-child:hover {{
        box-shadow: 0 6px 28px rgba(79, 70, 229, 0.55);
        transform: translateY(-1px);
    }}

    /* Metrik Kartları */
    .stMetric {{
        background: rgba(13, 18, 30, 0.75);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        backdrop-filter: blur(14px);
        padding: 14px 18px;
    }}

    /* KUSURSUZ SABİT LİKİT CAM ALT DOCK */
    .bottom-dock-fixed {{
        position: fixed;
        bottom: 18px;
        left: 0;
        right: 0;
        margin: 0 auto;
        width: fit-content;
        max-width: 94vw;
        z-index: 999999;
        background: rgba(13, 18, 30, 0.82);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 40px;
        backdrop-filter: blur(28px);
        -webkit-backdrop-filter: blur(28px);
        padding: 5px 14px;
        box-shadow: 0 12px 40px rgba(0, 0, 0, 0.6), inset 0 1px 0 rgba(255, 255, 255, 0.1);
    }}

    /* Alt Dock Giriş Kontrolleri */
    .bottom-dock-fixed div[data-testid="stSelectbox"] > div {{
        min-height: 32px !important;
        height: 32px !important;
        font-size: 0.8rem !important;
        border-radius: 20px !important;
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
    }}
    .bottom-dock-fixed div.stButton > button {{
        height: 32px !important;
        padding: 4px 12px !important;
        font-size: 0.78rem !important;
        border-radius: 20px !important;
        background: rgba(255, 255, 255, 0.08) !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        box-shadow: none !important;
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
            ws.column_dimensions[col_letter].width = max(m_len + 4, 12)

    return output.getvalue()

# --- KOMPAKT HERO ALANI ---
st.markdown(f"""
<div class='hero-container'>
    <div class='top-badge'>● {T['badge']}</div>
    <div class='hero-title'>{T['title']}</div>
    <div class='hero-sub'>{T['subtitle']}</div>
</div>
""", unsafe_allow_html=True)

# --- DOSYA YÜKLEME ALANI ---
yuklenen_dosyalar = st.file_uploader(
    T["upload_label"], 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True,
    label_visibility="collapsed",
    help=T["upload_hint"]
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
            sektor_secimi = T["industries"][st.session_state["industry_idx"]]
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
                    elif "ES" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Asiento", "Fecha", "Cuenta", "Nombre Cuenta", "Concepto", "Debe", "Haber"
                        kdv_adi = f"IVA Soportado {tax_rate}%"
                    elif "IT" in st.session_state["user_lang"]:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Partita", "Data", "Conto", "Descrizione", "Causale", "Dare", "Avere"
                        kdv_adi = f"IVA a Credito {tax_rate}%"
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

# --- KUSURSUZ LİKİT CAM ALT DOCK (FLOATING BAR) ---
st.markdown("<div class='bottom-dock-fixed'>", unsafe_allow_html=True)
col_b1, col_b2, col_b3, col_b4 = st.columns([1.8, 3.2, 3.2, 1.8])

with col_b1:
    dil_listesi = list(LANG_DATA.keys())
    mevcut_dil_idx = dil_listesi.index(st.session_state["user_lang"]) if st.session_state["user_lang"] in dil_listesi else 0
    yeni_dil = st.selectbox(
        "Dil",
        dil_listesi,
        index=mevcut_dil_idx,
        label_visibility="collapsed"
    )
    if yeni_dil != st.session_state["user_lang"]:
        st.session_state["user_lang"] = yeni_dil
        st.rerun()

with col_b2:
    secilen_tema = st.selectbox(
        "Görünüm",
        T["themes"],
        index=st.session_state["theme_idx"],
        label_visibility="collapsed"
    )
    yeni_t_idx = T["themes"].index(secilen_tema)
    if yeni_t_idx != st.session_state["theme_idx"]:
        st.session_state["theme_idx"] = yeni_t_idx
        st.rerun()

with col_b3:
    secilen_sektor = st.selectbox(
        "Sektör",
        T["industries"],
        index=st.session_state["industry_idx"],
        label_visibility="collapsed"
    )
    yeni_s_idx = T["industries"].index(secilen_sektor)
    if yeni_s_idx != st.session_state["industry_idx"]:
        st.session_state["industry_idx"] = yeni_s_idx
        st.rerun()

with col_b4:
    with st.popover(T["about_btn"]):
        st.markdown(f"#### {T['about_title']}")
        st.markdown(T["about_content"])

st.markdown("</div>", unsafe_allow_html=True)
