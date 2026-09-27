"""
================================================================================
LEDGERAI — ENTERPRISE GRADE AUTONOMOUS ACCOUNTING ENGINE
Architecture: Streamlit + Google Gemini GenAI SDK + Pandas + OpenPyXL
Design: Liquid Glass / Mesh Dynamics / Institutional FinTech Aesthetic
================================================================================
"""

import streamlit as st
import json
import time
import io
import datetime
import pandas as pd
from google import genai
from google.genai import types
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. CORE SYSTEM CONFIGURATION & INITIAL STATE
# ==============================================================================

st.set_page_config(
    page_title="LedgerAI — Autonomous Financial Terminal",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

SESSION_DEFAULTS = {
    "user_lang": "🇹🇷 TR",
    "theme_idx": 0,
    "industry_idx": 0,
    "chat_messages": [],
    "out_df": None,
    "h_deb": "Borç",
    "h_crd": "Alacak",
    "processed_docs_count": 0,
    "last_processing_time": 0.0
}

for key, default_val in SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default_val

if "GEMINI_API_KEY" in st.secrets:
    API_KEY = st.secrets["GEMINI_API_KEY"]
else:
    st.error("Missing GEMINI_API_KEY in Streamlit Secrets! Please configure.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# ==============================================================================
# 2. LOCALIZATION DATA DICTIONARY (6 GLOBAL STANDARDS)
# ==============================================================================

LANG_DATA = {
    "🇹🇷 TR": {
        "badge": "KURUMSAL OTONOM FİNANS TERMİNALİ",
        "title": "LedgerAI",
        "subtitle": "Faturaları saniyeler içinde sektörel hesap kodlarına ve dengeli ERP yevmiye fişine dönüştürün.",
        "drop_title": "Belgeleri Buraya Bırakın veya Seçin",
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
        "download_eta": "💾 ETA V.11 Uyumlu CSV",
        "industries": [
            "⚡ Otomatik Sektör (AI)",
            "🛒 Ticaret / Al-Sat (153 Ağırlıklı)",
            "🏢 Hizmet & Ofis (770/740)",
            "🏭 Üretim & Fabrika (150/730)"
        ],
        "themes": [
            "✨ Ultra Canlı Aurora",
            "🌌 Cyberpunk Gece",
            "🌑 Platin Titanyum"
        ],
        "about_btn": "ℹ️ İşleyiş & Güvenlik",
        "about_title": "LedgerAI Otonom Sistem Mimarisi",
        "about_content": """
        ### 🛡️ Kurumsal Finans & Güvenlik Mimarisi
        **LedgerAI**, kurumların fiş giriş maliyetlerini sıfıra indiren yeni nesil finans motorudur.
        * **1. Semantik OCR:** VKN, vergi dairesi, çoklu KDV oranları ve matrahlar kuruşu kuruşuna ayıklanır.
        * **2. Sektörel Mantık:** Alınan ürün; ticaret firmasında `153`, üretimde `150`, ofiste `770` olarak dinamik atanır.
        * **3. Çift Bakiye Garantisi:** Borç = Alacak denkliği sağlanmadan yevmiye fişi üretilmez.
        * **4. ERP Entegrasyonu:** ETA, Luca, Logo, Zirve, Datev ve QuickBooks'a doğrudan aktarılabilir formatta Excel çıkar.
        """,
        "step1_title": "1. Belge Analizi",
        "step1_desc": "OCR ile çoklu KDV, matrah ve satıcı bilgisi hatasız okunur.",
        "step2_title": "2. Sektörel Kodlama",
        "step2_desc": "Şirket türüne göre 153, 150 veya 770 hesapları atanır.",
        "step3_title": "3. Çift Taraflı Bakiye",
        "step3_desc": "Borç = Alacak denkliği kuruşu kuruşuna denetlenir.",
        "badge_erp": "✓ ETA • LUCA • DATEV • QUICKBOOKS UYUMLU",
        "badge_audit": "✓ %100 BORÇ/ALACAK DENGE GARANTİSİ",
        "badge_sec": "✓ OTONOM OCR & ÇİFT BAKİYE DENETİMİ",
        "bot_title": "👾 LedgerBot Finans Mentorü",
        "bot_welcome": "Selam! Ben finans asistanınım. Muhasebe öğrenmek veya pratik hesap kodlarını sormak için bana yazabilirsin. Kısa, net ve örnekle anlatırım!",
        "bot_placeholder": "Sorunu yaz (Örn: Laptop aldık nereye atayım? Borç-Alacak mantığı nedir?)...",
        "bot_clear": "🧹 Temizle",
        "headers": {
            "vouch": "Fiş No", "date": "Tarih", "code": "Hesap Kodu",
            "name": "Hesap Adı", "desc": "Açıklama", "curr": "Para Birimi",
            "deb": "Borç", "crd": "Alacak"
        }
    },
    "🇺🇸 EN": {
        "badge": "INSTITUTIONAL AI FINANCIAL TERMINAL",
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
        "download_eta": "💾 Generic CSV Format",
        "industries": [
            "⚡ Auto Industry (AI)", "🛒 Retail / Inventory (1200)",
            "🏢 Services / SaaS (OpEx)", "🏭 Manufacturing (COGS)"
        ],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platinum Titanium"],
        "about_btn": "ℹ️ How it Works & Security",
        "about_title": "LedgerAI Autonomous Architecture",
        "about_content": "Autonomous double-entry journal voucher generator compatible with US GAAP, Datev and PCG.",
        "step1_title": "1. Document Audit",
        "step1_desc": "Sub-millisecond OCR extraction of tax rates, net amounts, and vendor metadata.",
        "step2_title": "2. Contextual Mapping",
        "step2_desc": "Automated account mapping to OpEx, Inventory, or Capital Assets.",
        "step3_title": "3. Double-Entry Balance",
        "step3_desc": "Mathematical verification guaranteeing Total Debit equals Total Credit.",
        "badge_erp": "✓ QUICKBOOKS • XERO • DATEV • SAP READY",
        "badge_audit": "✓ 100% DEBIT/CREDIT BALANCE GUARANTEE",
        "badge_sec": "✓ SOC2 & BANK-GRADE DATA ENCRYPTION",
        "bot_title": "👾 LedgerBot Finance Mentor",
        "bot_welcome": "Hi! I am your AI finance mentor. Ask me any accounting concepts or codes. I reply concisely with direct practical examples!",
        "bot_placeholder": "Ask a question (e.g. How to book SaaS subscriptions? Debit vs Credit?)...",
        "bot_clear": "🧹 Clear",
        "headers": {
            "vouch": "Voucher #", "date": "Date", "code": "Account Code",
            "name": "Account Name", "desc": "Memo", "curr": "Currency",
            "deb": "Debit", "crd": "Credit"
        }
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
        "download_eta": "💾 Datev Format (CSV)",
        "industries": [
            "⚡ Automatisch (KI)", "🛒 Handel / Wareneinkauf",
            "🏢 Dienstleistung / IT", "🏭 Produktion / Fertigung"
        ],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platin Titan"],
        "about_btn": "ℹ️ Funktionsweise & Sicherheit",
        "about_title": "LedgerAI Architektur & Datev-Standard",
        "about_content": "Vollautomatisierte Buchungssatzerstellung nach Datev SKR03/04 Richtlinien.",
        "step1_title": "1. Belegprüfung",
        "step1_desc": "Präzise Vorsteueraufteilung und USt-IdNr Validierung in Sekunden.",
        "step2_title": "2. SKR03/04 Zuordnung",
        "step2_desc": "Automatische Kontierung nach Wareneinkauf, Kosten oder Anlagevermögen.",
        "step3_title": "3. Soll/Haben-Check",
        "step3_desc": "Revisionssichere Prüfung auf mathematische Ausgeglichenheit.",
        "badge_erp": "✓ DATEV SKR03/04 • SAP KOMPATIBEL",
        "badge_audit": "✓ 100% SOLL/HABEN AUSGEGLICHENHEIT",
        "badge_sec": "✓ DSGVO-KONFORME DATENVERARBEITUNG",
        "bot_title": "👾 LedgerBot Finanzmentor",
        "bot_welcome": "Hallo! Ich erkläre Buchhaltung kurz und präzise mit Beispielen für SKR03/04.",
        "bot_placeholder": "Frage eingeben...",
        "bot_clear": "🧹 Leeren",
        "headers": {
            "vouch": "Beleg", "date": "Datum", "code": "Konto",
            "name": "Bezeichnung", "desc": "Text", "curr": "Währung",
            "deb": "Soll", "crd": "Haben"
        }
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
        "download_eta": "💾 Format Standard PCG",
        "industries": [
            "⚡ Auto (IA)", "🛒 Négoce / Stock",
            "🏢 Services / Conseil", "🏭 Production / Industrie"
        ],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platine Titane"],
        "about_btn": "ℹ️ Fonctionnement & Sécurité",
        "about_title": "Architecture Comptable LedgerAI",
        "about_content": "Conformité Plan Comptable Général (PCG) avec vérification Débit = Crédit.",
        "step1_title": "1. Lecture OCR",
        "step1_desc": "Extraction des montants HT, TVA et identification du fournisseur.",
        "step2_title": "2. Ventilation PCG",
        "step2_desc": "Affectation automatique aux comptes de classe 6 selon l'activité.",
        "step3_title": "3. Contrôle Débit/Crédit",
        "step3_desc": "Vérification stricte de l'équilibre de chaque écriture de journal.",
        "badge_erp": "✓ CONFORME PCG • SAGE & CEGID READY",
        "badge_audit": "✓ ÉQUILIBRE DÉBIT/CRÉDIT GARANTI",
        "badge_sec": "✓ SÉCURITÉ CONFORME RGPD",
        "bot_title": "👾 LedgerBot Mentor",
        "bot_welcome": "Bonjour! Posez vos questions comptables PCG. Réponses claires, concises et illustrées.",
        "bot_placeholder": "Poser une question...",
        "bot_clear": "🧹 Effacer",
        "headers": {
            "vouch": "Pièce", "date": "Date", "code": "Compte",
            "name": "Libellé", "desc": "Détail", "curr": "Devise",
            "deb": "Débit", "crd": "Crédit"
        }
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
        "download_eta": "💾 Formato Contasol",
        "industries": [
            "⚡ Automático (IA)", "🛒 Comercio / Inventario",
            "🏢 Servicios / Oficina", "🏭 Fabricación / Industria"
        ],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platino Titanio"],
        "about_btn": "ℹ️ Funcionamiento y Seguridad",
        "about_title": "Arquitectura y Seguridad LedgerAI",
        "about_content": "Contabilidad autónoma con cuadre de Debe y Haber garantizado.",
        "step1_title": "1. Análisis de Factura",
        "step1_desc": "Lectura OCR avanzada de bases imponibles y tipos impositivos.",
        "step2_title": "2. Asignación PGC",
        "step2_desc": "Distribución en cuentas de gastos o existencias según la empresa.",
        "step3_title": "3. Cuadre de Asiento",
        "step3_desc": "Garantía matemática de que el Debe coincide con el Haber.",
        "badge_erp": "✓ COMPATIBLE A3 • SAGE • SOFTWARE FISCAL",
        "badge_audit": "✓ CUADRE DEBE = HABER GARANTIZADO",
        "badge_sec": "✓ CIFRADO DE DATOS BANCARIO",
        "bot_title": "👾 LedgerBot Mentor",
        "bot_welcome": "¡Hola! Pregúntame dudas contables del PGC. Respuestas directas, pedagógicas y breves.",
        "bot_placeholder": "Escribe tu duda...",
        "bot_clear": "🧹 Limpiar",
        "headers": {
            "vouch": "Asiento", "date": "Fecha", "code": "Cuenta",
            "name": "Nombre Cuenta", "desc": "Concepto", "curr": "Moneda",
            "deb": "Debe", "crd": "Haber"
        }
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
        "download_eta": "💾 Formato Zucchetti",
        "industries": [
            "⚡ Automatico (IA)", "🛒 Commercio / Magazzino",
            "🏢 Servizi / Consulenza", "🏭 Manifattura / Produzione"
        ],
        "themes": ["✨ Ultra Vivid Aurora", "🌌 Cyberpunk Night", "🌑 Platino Titanio"],
        "about_btn": "ℹ️ Funzionamento e Sicurezza",
        "about_title": "Architettura di Sicurezza LedgerAI",
        "about_content": "Generazione automatica di prima nota conforme ai principi contabili.",
        "step1_title": "1. Acquisizione Dati",
        "step1_desc": "Scansione OCR di aliquote IVA, imponibili e fornitore.",
        "step2_title": "2. Piano dei Conti",
        "step2_desc": "Classificazione tra costi di gestione, merci o cespiti ammortizzabili.",
        "step3_title": "3. Quadratura Dare/Avere",
        "step3_desc": "Verifica della perfetta parità contabile della scrittura.",
        "badge_erp": "✓ PRONTO PER ZUCCHETTI • TEAMSYSTEM • SAP",
        "badge_audit": "✓ QUADRATURA DARE/AVERE GARANTITA",
        "badge_sec": "✓ PROTEZIONE DATI STANDARD BANCARIO",
        "bot_title": "👾 LedgerBot Mentor",
        "bot_welcome": "Ciao! Chiedimi qualsiasi cosa sulla partita doppia. Risposte sintetiche e chiare con esempi!",
        "bot_placeholder": "Fai una domanda contabile...",
        "bot_clear": "🧹 Cancella",
        "headers": {
            "vouch": "Partita", "date": "Data", "code": "Conto",
            "name": "Descrizione", "desc": "Causale", "curr": "Valuta",
            "deb": "Dare", "crd": "Avere"
        }
    }
}

if st.session_state["user_lang"] not in LANG_DATA:
    st.session_state["user_lang"] = "🇹🇷 TR"

T = LANG_DATA[st.session_state["user_lang"]]

# ==============================================================================
# 3. DYNAMIC STYLING ENGINE (LUXURY FINTECH CSS)
# ==============================================================================

if st.session_state["theme_idx"] == 1:
    bg_style = """
        @keyframes cyberpunkPulse {
            0% { background-position: 0% 0%, 100% 100%, 50% 10%; filter: brightness(1) contrast(1.1); }
            50% { background-position: 100% 100%, 0% 0%, 50% 90%; filter: brightness(1.25) contrast(1.25); }
            100% { background-position: 0% 0%, 100% 100%, 50% 10%; filter: brightness(1) contrast(1.1); }
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
    bg_style = """
        @keyframes auroraRealFlow {
            0% { background-position: 0% 30%; filter: hue-rotate(0deg); }
            50% { background-position: 100% 70%; filter: hue-rotate(45deg); }
            100% { background-position: 0% 30%; filter: hue-rotate(0deg); }
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
    bg_style = """
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
    
    {bg_style}
    
    .stApp {{
        color: #F8FAFC;
        padding-bottom: 60px;
    }}

    /* MASTER GLASS TERMINAL */
    .master-console {{
        max-width: 920px;
        margin: 20px auto 0 auto;
        background: rgba(11, 16, 28, 0.78);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 28px;
        backdrop-filter: blur(36px);
        -webkit-backdrop-filter: blur(36px);
        padding: 38px 42px 28px 42px;
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
        font-size: 3.1rem;
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
        max-width: 620px;
        margin: 0 auto 24px auto;
    }}

    /* FILE UPLOADER REFINEMENT */
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

    /* ACTION BUTTON */
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

    /* 3 STEP PROCESS CARDS */
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
    }}
    .step-desc {{
        font-size: 0.72rem;
        color: #94A3B8;
        line-height: 1.4;
    }}

    /* TRUST BADGES */
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
    }}

    /* CONSOLE CONTROLS */
    .console-controls {{
        margin-top: 20px;
        padding-top: 16px;
        border-top: 1px solid rgba(255, 255, 255, 0.06);
    }}

    /* METRIC CARDS */
    .stMetric {{
        background: rgba(13, 18, 30, 0.75);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        backdrop-filter: blur(14px);
        padding: 14px 18px;
    }}

    /* SCROLLABLE CHAT CONTAINER (PREVENTS PAGE STRETCHING) */
    .chat-scroll-area {{
        max-height: 360px;
        overflow-y: auto;
        padding: 12px 14px;
        background: rgba(6, 9, 18, 0.55);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        margin-bottom: 14px;
    }}
    .chat-scroll-area::-webkit-scrollbar {{
        width: 6px;
    }}
    .chat-scroll-area::-webkit-scrollbar-track {{
        background: rgba(0, 0, 0, 0.2);
    }}
    .chat-scroll-area::-webkit-scrollbar-thumb {{
        background: rgba(99, 102, 241, 0.4);
        border-radius: 4px;
    }}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 4. INSTITUTIONAL EXCEL EXPORT ENGINE (OPENPYXL)
# ==============================================================================

def export_corporate_excel(df: pd.DataFrame, system_name: str = "Standard") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet_title = f"{system_name}_Journal"[:30]
        df.to_excel(writer, index=False, sheet_name=sheet_title)
        ws = writer.sheets[sheet_title]

        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        data_font = Font(name="Calibri", size=10)
        num_font = Font(name="Consolas", size=10)
        
        border_thin = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )

        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            col_name = str(col[0].value or '')

            for cell in col:
                cell.border = border_thin
                if cell.row != 1:
                    cell.alignment = Alignment(vertical="center")
                    if any(term in col_name.lower() for term in ["borç", "alacak", "debit", "credit", "soll", "haben"]):
                        cell.font = num_font
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.font = data_font

            ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    return output.getvalue()

def export_eta_csv(df: pd.DataFrame) -> bytes:
    headers = T["headers"]
    eta_df = pd.DataFrame()
    eta_df["FIS_NO"] = df[headers["vouch"]]
    eta_df["TARIH"] = df[headers["date"]]
    eta_df["HESAP_KODU"] = df[headers["code"]]
    eta_df["ACIKLAMA"] = df[headers["desc"]]
    eta_df["BORC"] = df[headers["deb"]].apply(lambda x: f"{x:.2f}".replace(".", ","))
    eta_df["ALACAK"] = df[headers["crd"]].apply(lambda x: f"{x:.2f}".replace(".", ","))
    
    return eta_df.to_csv(sep=";", index=False, encoding="utf-8-sig").encode("utf-8-sig")

# ==============================================================================
# 5. CORE AI RECOGNITION ENGINE (GEMINI MULTI-TIER AUDIT)
# ==============================================================================

def execute_document_audit(uploaded_files, sector_directive: str):
    results = []
    total = len(uploaded_files)
    progress_bar = st.progress(0)
    status_msg = st.empty()

    for idx, doc in enumerate(uploaded_files):
        status_msg.text(f"⚡ İşleniyor ({idx + 1}/{total}): {doc.name}...")
        raw_bytes = doc.read()
        mime_type = doc.type if doc.type else "application/pdf"

        prompt = f"""
        You are an elite autonomous financial auditor and ERP data extractor.
        {sector_directive}

        Examine the document carefully. Extract:
        1. Document Language & Origin: (TR, DE, FR, US, IT, ES).
        2. Currency: (TRY, USD, EUR, GBP).
        3. Vendor Information: Full legal name, Tax ID (VKN/EIN/SIRET/Steuernummer).
        4. Invoice Metadata: Official Invoice Number, Date (YYYY-MM-DD).
        5. Accounting Breakdown:
           - Correct Chart of Account Code based on country standards (TR Tek Düzen: 153, 770, 740, 150, 255).
           - Tax Breakdown: Net Amount, Tax Rate (1, 8, 10, 18, 20), Tax Amount.
           - Total Amount (Must equal Net Amount + Tax Amount).
        
        Respond ONLY with a valid JSON object. No markdown codeblocks:
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
        """

        max_retries = 3
        for attempt in range(max_retries):
            try:
                response = client.models.generate_content(
                    model="gemini-3.5-flash-lite",
                    contents=[
                        types.Part.from_bytes(data=raw_bytes, mime_type=mime_type),
                        prompt
                    ]
                )
                clean_text = response.text.replace("```json", "").replace("```", "").strip()
                data = json.loads(clean_text)
                data["filename"] = doc.name
                results.append(data)
                break
            except Exception as e:
                err_str = str(e)
                if ("503" in err_str or "429" in err_str) and attempt < max_retries - 1:
                    time.sleep(3 * (attempt + 1))
                    continue
                else:
                    st.warning(f"⚠️ {doc.name}: {err_str[:80]}")
                    break

        progress_bar.progress((idx + 1) / total)

    return results

# ==============================================================================
# 6. MASTER USER INTERFACE & LAYOUT
# ==============================================================================

# Central Master Console
st.markdown(f"""
<div class='master-console'>
    <div class='top-badge'>● {T['badge']}</div>
    <div class='hero-title'>{T['title']}</div>
    <div class='hero-sub'>{T['subtitle']}</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    T["drop_title"],
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    help=T["drop_sub"]
)

if uploaded_files:
    if len(uploaded_files) > 5:
        st.error(T["limit_err"])
    else:
        st.markdown(f"<div style='text-align:center; font-size:0.9rem; margin-top:8px;'>{T['ready_count'].format(count=len(uploaded_files))}</div>", unsafe_allow_html=True)

        if st.button(T["process_btn"], use_container_width=True):
            start_time = time.time()
            
            industry_name = T["industries"][st.session_state["industry_idx"]]
            directive = f"Company Profile: {industry_name}. "
            if "Ticaret" in industry_name or "Retail" in industry_name:
                directive += "Company operates in wholesale/retail trade. Core commercial goods MUST be classified as '153.01 Commercial Inventory' (or GAAP 1200). Office/fuel/meals are operating expenses (770)."
            elif "Hizmet" in industry_name or "Services" in industry_name:
                directive += "Company operates as a professional service/consulting provider. Classify project costs as 740 and overhead as 770."
            elif "Üretim" in industry_name or "Manufacturing" in industry_name:
                directive += "Company is a manufacturer. Raw material purchases MUST be '150 Raw Materials', factory expenses '730', administrative overhead '770'."
            else:
                directive += "Classify contextually: resale goods -> 153, operational supplies -> 770, capital equipment/computers -> 255."

            parsed_data = execute_document_audit(uploaded_files, directive)

            if parsed_data:
                headers = T["headers"]
                voucher_lines = []
                voucher_num = 1

                for item in parsed_data:
                    curr = item.get("currency", "TL")
                    inv_no = str(item.get("invoice_no") or "").strip()
                    date_val = str(item.get("date") or datetime.date.today().strftime("%Y-%m-%d")).strip()
                    vendor = str(item.get("vendor") or "Satıcı / Vendor").strip()
                    tax_id = str(item.get("tax_id") or "").strip()
                    acc_code = str(item.get("account_code") or "770.01").strip()
                    acc_name = str(item.get("account_name") or "Gider Hesabı").strip()

                    net = float(item.get("net") or 0.0)
                    tax = float(item.get("tax") or 0.0)
                    total = float(item.get("total") or (net + tax))
                    tax_rate = item.get("tax_rate") or 20

                    clean_name = "".join(c for c in vendor[:12] if c.isalnum()).upper() or "CARİ"
                    if "TR" in st.session_state["user_lang"]:
                        ap_code = f"320.{tax_id}" if tax_id else f"320.{clean_name}"
                        tax_code = f"191.{int(tax_rate):02d}"
                        tax_name = f"%{tax_rate} İndirilecek KDV"
                    else:
                        ap_code = f"2000-{tax_id}" if tax_id else f"VEND-{clean_name}"
                        tax_code = f"2200-TAX{tax_rate}"
                        tax_name = f"Sales Tax ({tax_rate}%)"

                    # 1. Debit Entry
                    voucher_lines.append({
                        headers["vouch"]: voucher_num,
                        headers["date"]: date_val,
                        headers["code"]: acc_code,
                        headers["name"]: acc_name,
                        headers["desc"]: f"{vendor} - {inv_no}",
                        headers["curr"]: curr,
                        headers["deb"]: net,
                        headers["crd"]: 0.0
                    })

                    # 2. Tax Entry
                    if tax > 0:
                        voucher_lines.append({
                            headers["vouch"]: voucher_num,
                            headers["date"]: date_val,
                            headers["code"]: tax_code,
                            headers["name"]: tax_name,
                            headers["desc"]: f"{vendor} - Tax",
                            headers["curr"]: curr,
                            headers["deb"]: tax,
                            headers["crd"]: 0.0
                        })

                    # 3. Credit Entry
                    voucher_lines.append({
                        headers["vouch"]: voucher_num,
                        headers["date"]: date_val,
                        headers["code"]: ap_code,
                        headers["name"]: vendor,
                        headers["desc"]: f"{vendor} - {inv_no}",
                        headers["curr"]: curr,
                        headers["deb"]: 0.0,
                        headers["crd"]: total
                    })

                    voucher_num += 1

                st.session_state["out_df"] = pd.DataFrame(voucher_lines)
                st.session_state["h_deb"] = headers["deb"]
                st.session_state["h_crd"] = headers["crd"]
                st.session_state["last_processing_time"] = round(time.time() - start_time, 2)
                st.session_state["processed_docs_count"] = len(parsed_data)
                st.success(f"{T['success']} ({st.session_state['last_processing_time']}s)")

# 3 Step Process Cards & Trust Badges
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
""", unsafe_allow_html=True)

# INTEGRATED IN-CONSOLE CONTROL DECK
st.markdown("<div class='console-controls'>", unsafe_allow_html=True)
c_ctrl1, c_ctrl2, c_ctrl3, c_ctrl4 = st.columns([1.8, 3.2, 3.2, 2.0])

with c_ctrl1:
    lang_keys = list(LANG_DATA.keys())
    curr_lang_idx = lang_keys.index(st.session_state["user_lang"]) if st.session_state["user_lang"] in lang_keys else 0
    new_lang = st.selectbox("Language / Dil", lang_keys, index=curr_lang_idx, label_visibility="collapsed")
    if new_lang != st.session_state["user_lang"]:
        st.session_state["user_lang"] = new_lang
        st.rerun()

with c_ctrl2:
    new_theme_str = st.selectbox("Theme / Görünüm", T["themes"], index=st.session_state["theme_idx"], label_visibility="collapsed")
    new_t_idx = T["themes"].index(new_theme_str)
    if new_t_idx != st.session_state["theme_idx"]:
        st.session_state["theme_idx"] = new_t_idx
        st.rerun()

with c_ctrl3:
    new_industry_str = st.selectbox("Industry / Sektör", T["industries"], index=st.session_state["industry_idx"], label_visibility="collapsed")
    new_i_idx = T["industries"].index(new_industry_str)
    if new_i_idx != st.session_state["industry_idx"]:
        st.session_state["industry_idx"] = new_i_idx
        st.rerun()

with c_ctrl4:
    with st.popover(T["about_btn"]):
        st.markdown(f"#### {T['about_title']}")
        st.markdown(T["about_content"])

st.markdown("</div></div>", unsafe_allow_html=True)

# ==============================================================================
# 7. INTERACTIVE JOURNAL VOUCHER GRID & EXPORT CENTER
# ==============================================================================

if st.session_state["out_df"] is not None:
    st.markdown("<div style='height: 35px;'></div>", unsafe_allow_html=True)
    st.subheader(T["preview_title"])
    st.caption(T["preview_tip"])

    edited_df = st.data_editor(
        st.session_state["out_df"],
        use_container_width=True,
        num_rows="dynamic"
    )

    deb_col = st.session_state["h_deb"]
    crd_col = st.session_state["h_crd"]

    tot_deb = edited_df[deb_col].sum()
    tot_crd = edited_df[crd_col].sum()
    diff = abs(tot_deb - tot_crd)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric(T["tot_deb"], f"{tot_deb:,.2f}")
    m2.metric(T["tot_crd"], f"{tot_crd:,.2f}")
    
    if diff < 0.05:
        m3.success(T["balanced"])
    else:
        m3.error(f"{T['unbalanced']} (Δ {diff:,.2f})")
        
    m4.metric("İşlem Süresi", f"{st.session_state['last_processing_time']} sn")

    exp_col1, exp_col2, exp_col3 = st.columns(3)
    
    with exp_col1:
        xlsx_data = export_corporate_excel(edited_df, system_name="LedgerAI")
        st.download_button(
            label=T["download_btn"],
            data=xlsx_data,
            file_name="ledger_journal_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with exp_col2:
        eta_data = export_eta_csv(edited_df)
        st.download_button(
            label=T["download_eta"],
            data=eta_data,
            file_name="eta_v11_aktarim.csv",
            mime="text/csv",
            use_container_width=True
        )

    with exp_col3:
        json_data = edited_df.to_json(orient="records", indent=2, force_ascii=False)
        st.download_button(
            label="💾 JSON Veri İndir",
            data=json_data,
            file_name="ledger_audit_data.json",
            mime="application/json",
            use_container_width=True
        )

# ==============================================================================
# 8. MENTOR FINANS ASİSTANI (OPTİMİZE EDİLMİŞ & KAYDIRMALI PENCERE)
# ==============================================================================

st.markdown("<div style='height: 35px;'></div>", unsafe_allow_html=True)
c_bot_l, c_bot_center, c_bot_r = st.columns([1, 4, 1])

with c_bot_center:
    with st.expander(T["bot_title"], expanded=False):
        top_col1, top_col2 = st.columns([5.5, 1.5])
        top_col1.caption(T["bot_welcome"])
        with top_col2:
            st.markdown("""
            <style>
                div[data-testid="stExpander"] div.stButton > button {
                    height: 28px !important;
                    min-height: 28px !important;
                    padding: 2px 10px !important;
                    font-size: 0.75rem !important;
                    border-radius: 8px !important;
                    white-space: nowrap !important;
                    margin-top: 0px !important;
                }
            </style>
            """, unsafe_allow_html=True)
            if st.button(T["bot_clear"], use_container_width=True):
                st.session_state["chat_messages"] = []
                st.rerun()

        # Sabit Yükseklikli, Kayan Chat Alanı
        st.markdown("<div class='chat-scroll-area'>", unsafe_allow_html=True)
        if not st.session_state["chat_messages"]:
            st.markdown(f"<div style='color: #64748B; font-size: 0.85rem; padding: 10px 0;'>💡 <i>{T['bot_placeholder']}</i></div>", unsafe_allow_html=True)
        else:
            for msg in st.session_state["chat_messages"][-4:]:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
        st.markdown("</div>", unsafe_allow_html=True)

        user_query = st.chat_input(T["bot_placeholder"])
        if user_query:
            st.session_state["chat_messages"].append({"role": "user", "content": user_query})
            
            prompt_bot = f"""
            Sen LedgerAI'ın kurumsal finans mentorü ve pratik muhasebe uzmanısın.
            Kullanıcı Dili: {st.session_state['user_lang']}
            Kullanıcı Sorusu: "{user_query}"

            TALİMATLAR:
            1. Asla lafı uzatma, gevezelik yapma, genel tanımlar yazma.
            2. MAKSİMUM 2-3 CÜMLEDE doğrudan ve net cevabı ver.
            3. Muhasebe öğrenmek isteyen birine anlatır gibi mantığını öğret:
               - "Şu hesaba gider, çünkü..." şeklinde kısaca sebebini söyle.
            4. Her cevabın sonuna tek satırlık somut fiş kaydı veya pratik örnek ekle:
               - Borç: 153 Ticari Mallar / 191 KDV
               - Alacak: 320 Satıcılar
            5. Türkiye için Tek Düzen kodlarını (153, 770, 740, 255 vb.), global için GAAP/Datev kodlarını kullan.
            """
            try:
                bot_resp = client.models.generate_content(
                    model="gemini-3.5-flash-lite",
                    contents=prompt_bot
                ).text.strip()
                st.session_state["chat_messages"].append({"role": "assistant", "content": bot_resp})
                st.rerun()
            except Exception:
                st.error("Asistan yanıt veremedi, lütfen tekrar deneyiniz.")
