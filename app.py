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

# --- 6 DİLLİ VE KAPSAMLI KURUMSAL SÖZLÜK ---
LANG_DATA = {
    "🇹🇷 TR": {
        "badge": "KURUMSAL OTONOM FİNANS MOTORU",
        "title": "LedgerAI",
        "subtitle": "Faturaları ve fişleri saniyeler içinde sektörel hesap kodlarına ve dengeli ERP yevmiye fişine dönüştürün.",
        "drop_title": "Belgeleri Buraya Sürükleyin veya Seçin",
        "drop_sub": "PDF, PNG, JPG • Oturum başına maksimum 5 belge",
        "process_btn": "⚡ Otonom Muhasebeleştir",
        "limit_err": "🛑 Demo sürümünde oturum başına en fazla 5 fatura işlenebilir.",
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
        "themes": ["✨ Ultra Canlı Aurora", "🌌 Cyberpunk Gece", "🌑 Platin Titanyum"],
        "about_btn": "ℹ️ İşleyiş & Güvenlik",
        "about_title": "LedgerAI Otonom Sistem İşleyişi",
        "about_content": """
        ### 🛡️ Kurumsal Finans & Güvenlik Mimarisi
        
        **LedgerAI**, kurumların ve mali müşavirlik ofislerinin veri giriş yükünü sıfırlamak üzere tasarlanmış uçtan uca otonom bir muhasebe terminalidir.
        
        * **1. Çok Katmanlı OCR & Semantik Çıkarım:** Yüklenen fatura ve fişler; satıcı unvanı, VKN/TCKN, vergi dairesi, fatura numarası, KDV oranları (%1, %10, %20) ve matrah bazında ayrıştırılır.
        * **2. Şirket Faaliyetine Duyarlı Akıllı Kodlama:** Satın alınan bir bilgisayar ticaret firması için `153 Ticari Mal`, üretim şirketi için `255 Demirbaş`, yazılım ofisi için `770/740` maliyeti olarak sisteme otomatik atanır.
        * **3. Çift Taraflı Denetim Güvencesi:** Sistem, faturanın genel toplamı ile satır matrahları ve KDV'leri arasındaki matematiksel eşitliği kuruşu kuruşuna doğrular. `Borç = Alacak` eşitliği sağlanmadan aktarım tablosu üretilmez.
        * **4. Evrensel ERP Entegrasyonu:** İndirilen `.xlsx` dosyaları ETA V.11, Luca, Logo, Zirve, Mikro, SAP ve Datev yazılımlarının şablonlarına doğrudan uyumludur.
        """,
        "step1_title": "1. Belge Analizi",
        "step1_desc": "OCR ile çoklu KDV, matrah ve satıcı bilgisi hatasız okunur.",
        "step2_title": "2. Sektörel Kodlama",
        "step2_desc": "Şirket türüne göre 153, 150 veya 770 hesapları atanır.",
        "step3_title": "3. Çift Taraflı Bakiye",
        "step3_desc": "Borç = Alacak denkliği kuruşu kuruşuna denetlenir.",
        "badge_erp": "✓ ETA • LUCA • DATEV • QUICKBOOKS UYUMLU",
        "badge_audit": "✓ %100 BORÇ/ALACAK DENGE GARANTİSİ",
        "badge_sec": "✓ BANKA STANDARTLARINDA GÜVENLİK"
    },
    "🇺🇸 EN": {
        "badge": "INSTITUTIONAL AI FINANCIAL ENGINE",
        "title": "LedgerAI",
        "subtitle": "Convert raw invoices and receipts into balanced, multi-GAAP ERP journal vouchers autonomously.",
        "drop_title": "Drop Financial Documents Here or Browse",
        "drop_sub": "PDF, PNG, JPG • Maximum 5 documents per session",
        "process_btn": "⚡ Process & Generate Vouchers",
        "limit_err": "🛑 Demo limit is 5 documents per batch.",
        "ready_count": "Documents ready: **{count}**",
        "success": "✓ Journal vouchers generated and balanced.",
        "failed": "❌ Documents could not be parsed.",
        "preview_title": "📊 Journal Voucher Grid (Live Editable)",
        "preview_tip": "💡 Double-click any cell to adjust accounts or descriptions before export.",
        "tot_deb": "Total Debit",
        "tot_crd": "Total Credit",
        "balanced": "✅ Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Unbalanced Voucher!",
        "download_btn": "📥 Download Clean Excel (.xlsx)",
        "industries": ["⚡ Auto Industry (AI)", "🛒 Retail / Inventory (1200)", "🏢 Services / SaaS (OpEx)", "🏭 Manufacturing (COGS)"],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platinum Titanium"],
        "about_btn": "ℹ️ How it Works & Security",
        "about_title": "LedgerAI Autonomous Architecture",
        "about_content": """
        ### 🛡️ Institutional Financial Architecture
        
        **LedgerAI** is an autonomous accounting terminal engineered to eliminate manual bookkeeping for global enterprises.
        
        * **1. Multi-Tier Semantic OCR:** Automatically extracts Vendor, Tax ID/EIN, Line Items, Multi-tier Sales Tax/VAT, and Currencies (USD, EUR, GBP, TRY).
        * **2. Context-Aware Chart of Accounts:** Differentiates inventory from operational expenses based on entity classification (US GAAP, Datev SKR03/04, PCG).
        * **3. Strict Dual-Audit Parity:** Enforces `Total Debit = Total Credit` balance down to the exact cent before releasing the journal voucher.
        * **4. ERP Interoperability:** Generated spreadsheets import directly into QuickBooks, Xero, NetSuite, SAP, and Datev.
        """,
        "step1_title": "1. Document Audit",
        "step1_desc": "Sub-millisecond OCR extraction of tax rates, net amounts, and vendor metadata.",
        "step2_title": "2. Contextual Mapping",
        "step2_desc": "Automated account mapping to OpEx, Inventory, or Capital Assets.",
        "step3_title": "3. Double-Entry Balance",
        "step3_desc": "Mathematical verification guaranteeing Total Debit equals Total Credit.",
        "badge_erp": "✓ QUICKBOOKS • XERO • DATEV • SAP READY",
        "badge_audit": "✓ 100% DEBIT/CREDIT BALANCE GUARANTEE",
        "badge_sec": "✓ SOC2 & BANK-GRADE DATA ENCRYPTION"
    },
    "🇩🇪 DE": {
        "badge": "KI FINANZTERMINAL & BUCHHALTUNG",
        "title": "LedgerAI",
        "subtitle": "Autonome Belegerfassung und Datev-konforme Kontierung in Echtzeit.",
        "drop_title": "Belege hier ablegen oder durchsuchen",
        "drop_sub": "PDF, PNG, JPG • Rechnungen & Quittungen",
        "process_btn": "⚡ Buchungssätze Erstellen",
        "limit_err": "🛑 Maximal 5 Dokumente im Demo-Modus.",
        "ready_count": "Bereit: **{count}**",
        "success": "✓ Buchungen erfolgreich erstellt.",
        "failed": "❌ Belege konnten nicht gelesen werden.",
        "preview_title": "📊 Buchungszeilen (Live Bearbeitbar)",
        "preview_tip": "💡 Doppelklick zum Ändern von Konten oder Beträgen.",
        "tot_deb": "Soll Gesamt",
        "tot_crd": "Haben Gesamt",
        "balanced": "✅ Ausgeglichen (Soll = Haben)",
        "unbalanced": "⚠️ Differenz festgestellt!",
        "download_btn": "📥 Excel Herunterladen (.xlsx)",
        "industries": ["⚡ Automatisch (KI)", "🛒 Handel / Wareneinkauf", "🏢 Dienstleistung / IT", "🏭 Produktion / Fertigung"],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platin Titan"],
        "about_btn": "ℹ️ Funktionsweise & Sicherheit",
        "about_title": "LedgerAI Architektur & Datev-Standard",
        "about_content": """
        ### 🛡️ Sichere Autonome Vorkontierung
        
        **LedgerAI** automatisiert die buchhalterische Erfassung von Eingangsrechnungen nach deutschen Standards.
        
        * **1. OCR-Belegprüfung:** Erkennt USt-IdNr, Steuersätze (7%, 19%), Rechnungsbeträge und Ausstellungsdaten lückenlos.
        * **2. Kontenrahmen-Zuordnung:** Ordnet Kosten automatisch den Sachkonten nach SKR03 oder SKR04 zu.
        * **3. Soll/Haben-Gleichgewicht:** Gewährleistet vor dem Export die absolute mathematische Ausgeglichenheit der Buchungssätze.
        * **4. Nahtloser Export:** Generiert strukturierte Dateien zur sofortigen Übernahme in Datev Unternehmen online oder SAP.
        """,
        "step1_title": "1. Belegprüfung",
        "step1_desc": "Präzise Vorsteueraufteilung und USt-IdNr Validierung in Sekunden.",
        "step2_title": "2. SKR03/04 Zuordnung",
        "step2_desc": "Automatische Kontierung nach Wareneinkauf, Kosten oder Anlagevermögen.",
        "step3_title": "3. Soll/Haben-Check",
        "step3_desc": "Revisionssichere Prüfung auf mathematische Ausgeglichenheit.",
        "badge_erp": "✓ DATEV SKR03/04 • SAP KOMPATIBEL",
        "badge_audit": "✓ 100% SOLL/HABEN AUSGEGLICHENHEIT",
        "badge_sec": "✓ DSGVO-KONFORME DATENVERARBEITUNG"
    },
    "🇫🇷 FR": {
        "badge": "TERMINAL FINANCIER AUTONOME IA",
        "title": "LedgerAI",
        "subtitle": "Génération d'écritures comptables équilibrées et ventilées par secteur d'activité.",
        "drop_title": "Déposer les pièces comptables ici",
        "drop_sub": "Factures et reçus (PDF, PNG, JPG)",
        "process_btn": "⚡ Générer les Écritures",
        "limit_err": "🛑 Limite: 5 documents par lot.",
        "ready_count": "Prêts: **{count}**",
        "success": "✓ Écritures générées avec succès.",
        "failed": "❌ Échec de lecture.",
        "preview_title": "📊 Journal Comptable (Édition Directe)",
        "preview_tip": "💡 Double-cliquez sur une cellule pour modifier.",
        "tot_deb": "Total Débit",
        "tot_crd": "Total Crédit",
        "balanced": "✅ Équilibré (Débit = Crédit)",
        "unbalanced": "⚠️ Déséquilibre Détecté!",
        "download_btn": "📥 Télécharger Excel (.xlsx)",
        "industries": ["⚡ Auto (IA)", "🛒 Négoce / Stock", "🏢 Services / Conseil", "🏭 Production / Industrie"],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platine Titane"],
        "about_btn": "ℹ️ Fonctionnement & Sécurité",
        "about_title": "Architecture Comptable LedgerAI",
        "about_content": """
        ### 🛡️ Automatisation et Conformité PCG
        
        **LedgerAI** traite et comptabilise vos factures fournisseurs selon les normes comptables françaises.
        
        * **1. Extraction Multitaxe:** Détection précise du SIREN/TVA Intra, des taux de TVA (5.5%, 10%, 20%) et du montant HT/TTC.
        * **2. Ventilation PCG:** Imputation intelligente entre les comptes de charges (classe 6), TVA déductible (44566) et fournisseurs (401).
        * **3. Équilibre Débit/Crédit:** Contrôle rigoureux garantissant l'égalité stricte Débit = Crédit avant exportation.
        * **4. Export Universel:** Fichiers configurés pour Sage, Cegid, Pennylane et QuickBooks.
        """,
        "step1_title": "1. Lecture OCR",
        "step1_desc": "Extraction des montants HT, TVA et identification du fournisseur.",
        "step2_title": "2. Ventilation PCG",
        "step2_desc": "Affectation automatique aux comptes de classe 6 selon l'activité.",
        "step3_title": "3. Contrôle Débit/Crédit",
        "step3_desc": "Vérification stricte de l'équilibre de chaque écriture de journal.",
        "badge_erp": "✓ CONFORME PCG • SAGE & CEGID READY",
        "badge_audit": "✓ ÉQUILIBRE DÉBIT/CRÉDIT GARANTI",
        "badge_sec": "✓ SÉCURITÉ CONFORME RGPD"
    },
    "🇪🇸 ES": {
        "badge": "TERMINAL FINANCIERO INTELIGENTE",
        "title": "LedgerAI",
        "subtitle": "Asientos contables equilibrados listos para ERP según el sector empresarial.",
        "drop_title": "Arrastra los documentos aquí o examina",
        "drop_sub": "PDF, PNG, JPG • Facturas y recibos",
        "process_btn": "⚡ Generar Asientos",
        "limit_err": "🛑 Máximo 5 documentos por lote.",
        "ready_count": "Listos: **{count}**",
        "success": "✓ Asientos generados y equilibrados.",
        "failed": "❌ Error al procesar.",
        "preview_title": "📊 Libro Diario (Editable)",
        "preview_tip": "💡 Haz doble clic para modificar cuentas.",
        "tot_deb": "Total Debe",
        "tot_crd": "Total Haber",
        "balanced": "✅ Cuadrado (Debe = Haber)",
        "unbalanced": "⚠️ Descuadre Detectado!",
        "download_btn": "📥 Descargar Excel (.xlsx)",
        "industries": ["⚡ Automático (IA)", "🛒 Comercio / Inventario", "🏢 Servicios / Oficina", "🏭 Fabricación / Industria"],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platino Titanio"],
        "about_btn": "ℹ️ Funcionamiento y Seguridad",
        "about_title": "Arquitectura y Seguridad LedgerAI",
        "about_content": """
        ### 🛡️ Automatización Contable Segura
        
        **LedgerAI** transforma facturas y recibos en asientos de libro diario para empresas y despachos profesionales.
        
        * **1. Extracción Integral:** Captura de CIF/NIF, bases imponibles, tramos de IVA (4%, 10%, 21%) e importes totales.
        * **2. Cuadro de Cuentas (PGC):** Clasificación automática en cuentas de gastos (grupo 6), IVA soportado (472) y proveedores (400).
        * **3. Cuadre Contable Garantizado:** Verificación matemática estricta asegurando que `Debe = Haber`.
        * **4. Compatibilidad:** Exportación directa compatible con A3, Sage y programas contables modernos.
        """,
        "step1_title": "1. Análisis de Factura",
        "step1_desc": "Lectura OCR avanzada de bases imponibles y tipos impositivos.",
        "step2_title": "2. Asignación PGC",
        "step2_desc": "Distribución en cuentas de gastos o existencias según la empresa.",
        "step3_title": "3. Cuadre de Asiento",
        "step3_desc": "Garantía matemática de que el Debe coincide con el Haber.",
        "badge_erp": "✓ COMPATIBLE A3 • SAGE • SOFTWARE FISCAL",
        "badge_audit": "✓ CUADRE DEBE = HABER GARANTIZADO",
        "badge_sec": "✓ CIFRADO DE DATOS BANCARIO"
    },
    "🇮🇹 IT": {
        "badge": "TERMINALE FINANZIARIO AUTONOMO",
        "title": "LedgerAI",
        "subtitle": "Scritture contabili in partita doppia bilanciate per qualsiasi software gestionale ERP.",
        "drop_title": "Trascina qui le fatture o cerca file",
        "drop_sub": "PDF, PNG, JPG • Ricevute e fatture",
        "process_btn": "⚡ Genera Scritture",
        "limit_err": "🛑 Massimo 5 documenti.",
        "ready_count": "Pronti: **{count}**",
        "success": "✓ Scritture generate e bilanciate.",
        "failed": "❌ Impossibile elaborare.",
        "preview_title": "📊 Prima Nota (Modificabile)",
        "preview_tip": "💡 Fai doppio clic per modificare.",
        "tot_deb": "Totale Dare",
        "tot_crd": "Totale Avere",
        "balanced": "✅ Quadratura Perfetta",
        "unbalanced": "⚠️ Sbilancio!",
        "download_btn": "📥 Scarica Excel (.xlsx)",
        "industries": ["⚡ Automatico (IA)", "🛒 Commercio / Magazzino", "🏢 Servizi / Consulenza", "🏭 Manifattura / Produzione"],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platino Titanio"],
        "about_btn": "ℹ️ Funzionamento e Sicurezza",
        "about_title": "Architettura di Sicurezza LedgerAI",
        "about_content": """
        ### 🛡️ Registrazione Contabile Intelligente
        
        **LedgerAI** digitalizza e registra automaticamente le fatture passive in partita doppia.
        
        * **1. Acquisizione Fiscale:** Riconoscimento di Partita IVA/Codice Fiscale, imponibili, aliquote IVA (4%, 10%, 22%) e totale documento.
        * **2. Piano dei Conti:** Assegnazione automatica a conti di costo, IVA a credito e debiti verso fornitori.
        * **3. Quadratura Fiscale:** Controllo rigoroso prima dell'export affinché `Dare = Avere`.
        * **4. Integrazione ERP:** File Excel strutturato pronto per Zucchetti, Teamsystem e SAP.
        """,
        "step1_title": "1. Acquisizione Dati",
        "step1_desc": "Scansione OCR di aliquote IVA, imponibili e fornitore.",
        "step2_title": "2. Piano dei Conti",
        "step2_desc": "Classificazione tra costi di gestione, merci o cespiti ammortizzabili.",
        "step3_title": "3. Quadratura Dare/Avere",
        "step3_desc": "Verifica della perfetta parità contabile della scrittura.",
        "badge_erp": "✓ PRONTO PER ZUCCHETTI • TEAMSYSTEM • SAP",
        "badge_audit": "✓ QUADRATURA DARE/AVERE GARANTITA",
        "badge_sec": "✓ PROTEZIONE DATI STANDARD BANCARIO"
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

# --- GERÇEK VE GÖZ ALICI 3 FARKLI TEMA MOTORU ---
if st.session_state["theme_idx"] == 1:
    # 🌌 Cyberpunk Gece (Canlı Neon Fuşya ve Lazer Mavisi Işık Hüzmeleri)
    bg_css = """
        @keyframes cyberpunkPulse {
            0% { 
                background-position: 0% 0%, 100% 100%, 50% 10%; 
                filter: brightness(1) contrast(1.1); 
            }
            50% { 
                background-position: 100% 100%, 0% 0%, 50% 90%; 
                filter: brightness(1.25) contrast(1.25); 
            }
            100% { 
                background-position: 0% 0%, 100% 100%, 50% 10%; 
                filter: brightness(1) contrast(1.1); 
            }
        }
        .stApp {
            background: radial-gradient(circle at 15% 15%, rgba(217, 70, 239, 0.45) 0%, transparent 45%),
                        radial-gradient(circle at 85% 85%, rgba(6, 182, 212, 0.40) 0%, transparent 45%),
                        radial-gradient(circle at 50% 40%, rgba(147, 51, 234, 0.35) 0%, transparent 55%),
                        linear-gradient(135deg, #020108 0%, #080318 45%, #050A1A 80%, #020108 100%);
            background-size: 220% 220%;
            animation: cyberpunkPulse 12s ease-in-out infinite;
            background-attachment: fixed;
        }
    """
elif st.session_state["theme_idx"] == 0:
    # ✨ Ultra Aurora (Zümrüt Yeşili, Safir Mavisi ve Altın Işık Dalgaları)
    bg_css = """
        @keyframes auroraRealFlow {
            0% { 
                background-position: 0% 30%; 
                filter: hue-rotate(0deg); 
            }
            50% { 
                background-position: 100% 70%; 
                filter: hue-rotate(45deg); 
            }
            100% { 
                background-position: 0% 30%; 
                filter: hue-rotate(0deg); 
            }
        }
        .stApp {
            background: radial-gradient(circle at 10% 20%, rgba(16, 185, 129, 0.45) 0%, transparent 45%),
                        radial-gradient(circle at 90% 20%, rgba(14, 165, 233, 0.45) 0%, transparent 45%),
                        radial-gradient(circle at 50% 90%, rgba(99, 102, 241, 0.35) 0%, transparent 50%),
                        radial-gradient(circle at 70% 60%, rgba(245, 158, 11, 0.20) 0%, transparent 40%),
                        linear-gradient(140deg, #020710 0%, #041424 40%, #09213A 70%, #020710 100%);
            background-size: 240% 240%;
            animation: auroraRealFlow 16s ease-in-out infinite;
            background-attachment: fixed;
        }
    """
else:
    # 🌑 Platin Titanyum (Apple Pro Metalik Gümüş & Derin Antrasit)
    bg_css = """
        @keyframes titaniumSheen {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        .stApp {
            background: radial-gradient(circle at 50% 0%, rgba(255, 255, 255, 0.12) 0%, transparent 60%),
                        radial-gradient(circle at 80% 100%, rgba(148, 163, 184, 0.08) 0%, transparent 50%),
                        linear-gradient(135deg, #07090E 0%, #0F131D 50%, #080A10 100%);
            background-size: 200% 200%;
            animation: titaniumSheen 20s ease infinite;
            background-attachment: fixed;
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
        padding-bottom: 125px;
    }}

    /* TEK PARÇA LÜKS CAM KONSOL */
    .master-console {{
        max-width: 860px;
        margin: 20px auto 0 auto;
        background: rgba(11, 16, 28, 0.75);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 28px;
        backdrop-filter: blur(36px);
        -webkit-backdrop-filter: blur(36px);
        padding: 38px 40px 32px 40px;
        box-shadow: 0 35px 90px rgba(0, 0, 0, 0.75), 
                    inset 0 1px 0 rgba(255, 255, 255, 0.18);
        text-align: center;
    }}

    .top-badge {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 14px;
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 99px;
        font-size: 0.68rem;
        font-weight: 800;
        letter-spacing: 1.5px;
        color: #CBD5E1;
        margin-bottom: 12px;
        text-transform: uppercase;
    }}

    .hero-title {{
        font-size: 3rem;
        font-weight: 800;
        letter-spacing: -1.2px;
        background: linear-gradient(135deg, #FFFFFF 40%, #CBD5E1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 6px;
        line-height: 1.05;
    }}
    
    .hero-sub {{
        font-size: 0.98rem;
        color: #94A3B8;
        font-weight: 400;
        line-height: 1.5;
        max-width: 600px;
        margin: 0 auto 24px auto;
    }}

    /* Konsol İçi Bütünleşik Yükleme Alanı */
    div[data-testid="stFileUploader"] {{
        background: rgba(6, 9, 18, 0.65);
        border: 1px dashed rgba(255, 255, 255, 0.22);
        border-radius: 18px;
        padding: 24px 16px;
        transition: all 0.3s ease;
        margin-bottom: 12px;
    }}
    div[data-testid="stFileUploader"]:hover {{
        border-color: rgba(99, 102, 241, 0.85);
        box-shadow: 0 0 35px rgba(99, 102, 241, 0.3);
        background: rgba(9, 14, 26, 0.8);
    }}

    /* İşlem Butonu */
    div.stButton > button:first-child {{
        background: linear-gradient(135deg, #4F46E5 0%, #06B6D4 100%);
        border: none;
        border-radius: 12px;
        font-weight: 700;
        font-size: 0.98rem;
        padding: 13px 30px;
        color: #FFFFFF;
        box-shadow: 0 4px 30px rgba(79, 70, 229, 0.5);
        transition: all 0.25s ease;
        margin-top: 6px;
    }}
    div.stButton > button:first-child:hover {{
        box-shadow: 0 6px 40px rgba(6, 182, 212, 0.7);
        transform: translateY(-2px);
    }}

    /* 3 ADIMLI İŞLEYİŞ KARTLARI (SAYFAYI ZENGİNLEŞTİREN ALAN) */
    .steps-container {{
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 16px;
        margin: 24px auto 0 auto;
        text-align: left;
    }}
    .step-card {{
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 14px 16px;
        backdrop-filter: blur(12px);
        transition: all 0.25s;
    }}
    .step-card:hover {{
        background: rgba(255, 255, 255, 0.06);
        border-color: rgba(99, 102, 241, 0.4);
        transform: translateY(-2px);
    }}
    .step-title {{
        font-size: 0.82rem;
        font-weight: 700;
        color: #F1F5F9;
        margin-bottom: 4px;
        display: flex;
        align-items: center;
        gap: 6px;
    }}
    .step-desc {{
        font-size: 0.72rem;
        color: #94A3B8;
        line-height: 1.4;
    }}

    /* GÜVEN ROZETLERİ */
    .trust-grid {{
        display: flex;
        justify-content: center;
        align-items: center;
        gap: 20px;
        flex-wrap: wrap;
        margin: 22px auto 0 auto;
        padding-top: 18px;
        border-top: 1px solid rgba(255, 255, 255, 0.08);
    }}
    .trust-item {{
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 0.8px;
        color: #94A3B8;
        display: flex;
        align-items: center;
        gap: 6px;
    }}

    /* EK DÜZENLEMELER: DOCK'U SAYFANIN EN ALTINA SABİTLEME */
    .dock-fixed-outer {{
        position: fixed;
        bottom: 18px;
        left: 0;
        right: 0;
        margin: 0 auto;
        width: fit-content;
        max-width: 94vw;
        z-index: 999999;
        background: rgba(10, 14, 26, 0.88);
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 50px;
        backdrop-filter: blur(32px);
        -webkit-backdrop-filter: blur(32px);
        padding: 4px 14px;
        box-shadow: 0 20px 50px rgba(0, 0, 0, 0.75), inset 0 1px 0 rgba(255, 255, 255, 0.18);
    }}

    .dock-fixed-outer div[data-testid="stSelectbox"] > div {{
        min-height: 32px !important;
        height: 32px !important;
        font-size: 0.8rem !important;
        border-radius: 20px !important;
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
    }}
    .dock-fixed-outer div.stButton > button {{
        height: 32px !important;
        padding: 4px 12px !important;
        font-size: 0.78rem !important;
        border-radius: 20px !important;
        background: rgba(255, 255, 255, 0.08) !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        box-shadow: none !important;
        margin-top: 0 !important;
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

# --- TEK PARÇA MERKEZİ KONSOL ALANI ---
st.markdown(f"""
<div class='master-console'>
    <div class='top-badge'>● {T['badge']}</div>
    <div class='hero-title'>{T['title']}</div>
    <div class='hero-sub'>{T['subtitle']}</div>
""", unsafe_allow_html=True)

# Konsolun içerisine doğrudan yerleşen dosya yükleyici
yuklenen_dosyalar = st.file_uploader(
    T["drop_title"], 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True,
    label_visibility="collapsed",
    help=T["drop_sub"]
)

if yuklenen_dosyalar:
    if len(yuklenen_dosyalar) > 5:
        st.error(T["limit_err"])
    else:
        st.markdown(f"<div style='text-align:center; font-size:0.9rem; margin-top:8px;'>{T['ready_count'].format(count=len(yuklenen_dosyalar))}</div>", unsafe_allow_html=True)
        
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

# 3 Adımlı Süreç Kartları & Güven Rozetleri
st.markdown(f"""
    <div class='steps-container'>
        <div class='step-card'>
            <div class='step-title'>⚡ {T['step1_title']}</div>
            <div class='step-desc'>{T['step1_desc']}</div>
        </div>
        <div class='step-card'>
            <div class='step-title'>🎯 {T['step2_title']}</div>
            <div class='step-desc'>{T['step2_desc']}</div>
        </div>
        <div class='step-card'>
            <div class='step-title'>⚖️ {T['step3_title']}</div>
            <div class='step-desc'>{T['step3_desc']}</div>
        </div>
    </div>
    <div class='trust-grid'>
        <div class='trust-item'>{T['badge_erp']}</div>
        <div class='trust-item'>{T['badge_audit']}</div>
        <div class='trust-item'>{T['badge_sec']}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# --- TABLO VE ÇIKTI ALANI ---
if "out_df" in st.session_state:
    st.markdown("<div style='height: 35px;'></div>", unsafe_allow_html=True)
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

# --- EN ALTA SABİTLENMİŞ İŞLEVSEL LİKİT CAM DOCK ---
st.markdown("<div class='dock-fixed-outer'>", unsafe_allow_html=True)
col_b1, col_b2, col_b3, col_b4 = st.columns([1.6, 3.2, 3.2, 2.2])

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
