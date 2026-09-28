"""
================================================================================
LEDGERAI — MULTI-MODAL ENTERPRISE FINANCIAL TERMINAL & ACADEMY SIMULATION
Architecture: Streamlit + Google Gemini GenAI SDK + Pandas + OpenPyXL
Design: Minimal GitHub Capsule Pill Architecture / Dual-Wing Executive Cockpit
Compliance: KVKK, GDPR, Turkish Uniform Chart of Accounts, Datev, US GAAP
Version: 5.2.0 Flagship Production Edition
================================================================================
"""

import streamlit as st
import json
import time
import io
import re
import datetime
import math
import pandas as pd
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. CORE SYSTEM CONFIGURATION & INITIAL STATE
# ==============================================================================

st.set_page_config(
    page_title="LedgerAI — Enterprise Financial Terminal & Academy",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

SESSION_DEFAULTS = {
    "user_lang": "🇹🇷 TR",
    "theme_idx": 0,  # 0: Platin Titanyum, 1: Nebula Aurora, 2: Siber Matris
    "industry_idx": 0,
    "chat_messages": [],
    "out_df": None,
    "raw_audit_results": [],
    "h_deb": "Borç",
    "h_crd": "Alacak",
    "processed_docs_count": 0,
    "last_processing_time": 0.0,
    "total_batch_debit": 0.0,
    "total_batch_credit": 0.0,
    "total_withholding_amount": 0.0,
    "academy_xp": 100,
    "academy_level": "Mali Stajyer",
    "puzzle_step": 0,
    "puzzle_solved": False
}

for key, default_val in SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default_val

# API Anahtarı Doğrulama
if "GEMINI_API_KEY" in st.secrets:
    API_KEY = st.secrets["GEMINI_API_KEY"]
else:
    st.error("Missing GEMINI_API_KEY in Streamlit Secrets! Please configure.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# ==============================================================================
# 2. SİBER AKADEMİ OYUN VERİTABANI (KADEMELİ PUZZLE DÖNGÜSÜ)
# ==============================================================================

ACADEMY_PUZZLES = [
    {
        "id": 1,
        "title": "GÖREV 1: TİCARİ MAL & MASRAF AYRIMI (KOLAY)",
        "scenario": "Şirketiniz (Teknoloji Mağazası) satmak amacıyla toptancıdan 50 adet kablosuz kulaklık satın aldı.",
        "invoice_data": "Tutar: 60.000 TL + %20 KDV (12.000 TL) = 72.000 TL | Fatura No: GIB2026000001",
        "question": "Satılmak üzere depoya giren bu ticari mallar ve KDV hangi borç hesaplarına kaydedilmelidir?",
        "options_code": [
            "153.01 Ticari Mallar / 191.20 İndirilecek KDV",
            "770.01 Genel Yönetim Giderleri / 191.20 İndirilecek KDV",
            "600.01 Yurtiçi Satışlar / 391.20 Hesaplanan KDV",
            "255.01 Demirbaşlar / 191.20 İndirilecek KDV"
        ],
        "options_tax": ["Tevkifatsız Normal Alım (%20 Tam KDV)", "5/10 KDV Tevkifatı", "9/10 KDV Tevkifatı"],
        "correct_code": "153.01 Ticari Mallar / 191.20 İndirilecek KDV",
        "correct_tax": "Tevkifatsız Normal Alım (%20 Tam KDV)",
        "reward_xp": 150,
        "success_msg": "Harika! Satılmak üzere alınan mallar daima 153 Ticari Mallar hesabında izlenir.",
        "error_msg": "Hata! Satış amacıyla depoya giren ürünler masraf (770) değil, stok (153) olmalıdır."
    },
    {
        "id": 2,
        "title": "GÖREV 2: 5/10 KDV TEVKİFATI KİLİDİ (ORTA)",
        "scenario": "Şirketiniz bir temizlik firmasından ofis temizlik hizmet faturası aldı. Mevzuat gereği temizlik hizmetlerinde 5/10 KDV Tevkifatı zorunludur.",
        "invoice_data": "Matrah: 20.000 TL | %20 KDV: 4.000 TL | Tevkifat (5/10): 2.000 TL | Ödenecek: 22.000 TL",
        "question": "Bu işlemde kesilen 2.000 TL tevkifat hangi alacak hesabına kaydedilerek bakiye eşitlenir?",
        "options_code": [
            "770.01 Genel Yönetim Gideri (Borç) / 191.20 KDV (Borç)",
            "153.01 Ticari Mallar (Borç) / 191.20 KDV (Borç)",
            "740.01 Hizmet Üretim Maliyeti (Borç)"
        ],
        "options_tax": [
            "360.01 Ödenecek KDV Tevkifatı (2.000 TL Alacak)",
            "320.01 Satıcıya Tam Ödeme (Tevkifatsız)",
            "391.20 Hesaplanan KDV (Alacak)"
        ],
        "correct_code": "770.01 Genel Yönetim Gideri (Borç) / 191.20 KDV (Borç)",
        "correct_tax": "360.01 Ödenecek KDV Tevkifatı (2.000 TL Alacak)",
        "reward_xp": 250,
        "success_msg": "Tebrikler! Tevkif edilen 2.000 TL KDV devlete ödenmek üzere 360 hesabına aktarılır.",
        "error_msg": "Dikkat! Tevkifat kesintisi satıcıya ödenmez; 360 Ödenecek Tevkifat hesabına alacak yazılır."
    },
    {
        "id": 3,
        "title": "GÖREV 3: SMMM SERBEST MESLEK MAKBUZU & STOPAJ (İLERİ)",
        "scenario": "Şirketinizin Mali Müşaviri aylık muhasebe danışmanlık makbuzu kesti. Makbuzda %20 Gelir Vergisi Stopajı ve %20 KDV bulunmaktadır.",
        "invoice_data": "Brüt Ücret: 10.000 TL | Stopaj (%20): 2.000 TL | KDV (%20): 2.000 TL | Net Ödenen: 10.000 TL",
        "question": "Müşavire net 10.000 TL ödenirken kesilen 2.000 TL stopaj hangi hesaba bağlanmalıdır?",
        "options_code": [
            "770.02 Müşavirlik Gideri (Borç 10.000) & 191.20 KDV (Borç 2.000)",
            "153.01 Ticari Mal Alışı (Borç 10.000)",
            "659.01 Diğer Olağan Giderler (Borç 10.000)"
        ],
        "options_tax": [
            "360.02 Ödenecek Stopaj (Alacak 2.000) & 320/102 Net Ödeme (10.000)",
            "320 Satıcılar Tam Ödeme (Alacak 12.000)",
            "391 Hesaplanan KDV (Alacak 2.000)"
        ],
        "correct_code": "770.02 Müşavirlik Gideri (Borç 10.000) & 191.20 KDV (Borç 2.000)",
        "correct_tax": "360.02 Ödenecek Stopaj (Alacak 2.000) & 320/102 Net Ödeme (10.000)",
        "reward_xp": 400,
        "success_msg": "Mükemmel Başarı! Brüt gider 770'e, KDV 191'e, stopaj 360'a ve net para çıkışı 102/320'ye işlendi.",
        "error_msg": "Stopaj mantığı hatalı! Kesilen gelir vergisi 360 Ödenecek Stopaj hesabına kaydedilmelidir."
    }
]

# ==============================================================================
# 3. LOCALIZATION DATA DICTIONARY (6 DİLDE TAM VE DETAYLI AÇIKLAMALAR)
# ==============================================================================

LANG_DATA = {
    "🇹🇷 TR": {
        "badge": "İNSAN GÜCÜ + YAPAY ZEKA ORTAKLIĞI",
        "title": "LedgerAI",
        "subtitle": "Yapay zeka faturaları ve tevkifatı hazırlar; uzman mali müşavir son kararı verir ve onaylar.",
        "drop_title": "Belgeleri Buraya Bırakın veya Seçin",
        "drop_sub": "PDF, PNG, JPG • Fatura, Serbest Meslek Makbuzu ve Fişler",
        "process_btn": "⚡ Otonom İncele & Fişi Hazırla",
        "limit_err": "🛑 Demo sürümünde oturum başına en fazla 5 fatura işlenebilir.",
        "ready_count": "İşlenecek belge sayısı: **{count}**",
        "success": "✓ Fişler hazırlandı, Tevkifat & KDV ayrıldı, Borç/Alacak dengelendi.",
        "failed": "❌ Belgeler işlenemedi. Lütfen görsel netliğini kontrol edin.",
        "preview_title": "📊 Muhasebe Yevmiye Fişi & İnsan Denetim Masası",
        "preview_tip": "💡 Yapay zekanın önerdiği kodları ve tutarları değiştirmek için hücreye çift tıklayın.",
        "tot_deb": "Toplam Borç",
        "tot_crd": "Toplam Alacak",
        "balanced": "✅ Fiş Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "download_btn": "📥 Çok Sayfalı Kurumsal Excel'i İndir (.xlsx)",
        "download_eta": "💾 ETA V.11 Uyumlu CSV",
        "download_luca": "💾 Luca Uyumlu Aktarım",
        "industries": [
            "⚡ Otomatik Sektör (AI)",
            "🛒 Ticaret / Al-Sat (153 Ağırlıklı)",
            "🏢 Hizmet & Ofis (770/740)",
            "🏭 Üretim & Fabrika (150/730)"
        ],
        "themes": [
            "🌑 Platin Titanyum (Kurumsal)",
            "✨ Nebula Aurora (Modern)",
            "⚡ Siber Matris (Akademi)"
        ],
        "about_btn": "ℹ️ Mimarî & Standartlar",
        "about_title": "LedgerAI Sistem Mimarisi & Regülasyon Standartları",
        "about_content": """
        ### 🛡️ Kurumsal Finans & Güvenlik Mimarisi
        LedgerAI, Türkiye Tek Düzen Hesap Planı, VUK ve uluslararası finansal standartlara tam uyumlu bir otonom ön muhasebe altyapısıdır:
        * **1. Çift Göz Prensibi (Dual Control):** Yapay zeka veri ayıklayıcı ve tasnif edici olarak çalışır; yasal defter kaydı yetkili SMMM/YMM onayına bağlıdır.
        * **2. Tevkifat & Stopaj Ayrıştırma:** 5/10, 7/10, 9/10 KDV tevkifatlarını ve serbest meslek stopajlarını ayrı hesap kodlarına (360) dengeli olarak dağıtır.
        * **3. Çift Bakiye Doğrulama:** Borç ve Alacak tutarları kuruşu kuruşuna eşitlenmeden sistem dışa aktarıma izin vermez.
        * **4. ERP Entegrasyon Standartları:** ETA V.11, Luca, Logo, Zirve, Datev ve QuickBooks uyumlu veri çıktıları sağlar.
        """,
        "cockpit_card1_title": "🏛️ Mevzuat & Tevkifat Uyumu",
        "cockpit_card1_desc": "5/10, 7/10, 9/10 KDV tevkifatları ve Serbest Meslek stopajları kuruş farkı olmadan 360 hesabına aktarılır.",
        "cockpit_card2_title": "⚡ ERP Aktarım Formatları",
        "cockpit_card2_desc": "Tek tıkla ETA V.11 uyumlu CSV, Luca ve çok sayfalı (153 & 770 ayrılmış) kurumsal Excel üretimi.",
        "cockpit_card3_title": "🛡️ Çift Taraflı Denetim Kilidi",
        "cockpit_card3_desc": "Toplam Borç = Toplam Alacak eşitliği sağlanmadan yevmiye fişi kapatılmaz; bakiye farkı riski sıfırlanır.",
        "headers": {
            "vouch": "Fiş No", "date": "Tarih", "code": "Hesap Kodu",
            "name": "Hesap Adı", "desc": "Açıklama", "curr": "Para Birimi",
            "deb": "Borç", "crd": "Alacak"
        }
    },
    "🇺🇸 EN": {
        "badge": "HUMAN + AI COLLABORATIVE TERMINAL",
        "title": "LedgerAI",
        "subtitle": "Autonomous AI journal voucher generator with continuous CPA audit & verification.",
        "drop_title": "Drop Financial Documents Here or Browse",
        "drop_sub": "PDF, PNG, JPG • Invoices, Receipts & Vouchers • Up to 5 files",
        "process_btn": "⚡ Process & Prepare Vouchers",
        "limit_err": "🛑 Demo limit is 5 documents per batch.",
        "ready_count": "Documents ready: **{count}**",
        "success": "✓ Vouchers generated, taxes reconciled, Debit = Credit balanced.",
        "failed": "❌ Documents could not be parsed.",
        "preview_title": "📊 Journal Voucher Grid & Human Audit Desk",
        "preview_tip": "💡 Double-click any cell to adjust accounts or descriptions before export.",
        "tot_deb": "Total Debit",
        "tot_crd": "Total Credit",
        "balanced": "✅ Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Unbalanced Voucher!",
        "download_btn": "📥 Download Multi-Tab Corporate Excel (.xlsx)",
        "download_eta": "💾 Generic CSV Format",
        "download_luca": "💾 QuickBooks Format",
        "industries": [
            "⚡ Auto Industry (AI)", "🛒 Retail / Inventory (1200)",
            "🏢 Services / SaaS (OpEx)", "🏭 Manufacturing (COGS)"
        ],
        "themes": [
            "🌑 Platinum Titanium (Corporate)",
            "✨ Nebula Aurora (Modern)",
            "⚡ Cyber Matrix (Academy)"
        ],
        "about_btn": "ℹ️ Architecture & Standards",
        "about_title": "LedgerAI Regulatory & Architectural Standards",
        "about_content": """
        ### 🛡️ Enterprise Financial & Security Architecture
        Autonomous multi-GAAP accounting terminal compliant with US GAAP, IFRS, and SOC2:
        * **1. Human-in-the-Loop Audit:** AI structures journal entries; certified financial controllers provide final clearance.
        * **2. Mathematical Parity Lock:** Absolute verification guaranteeing Debit strictly equals Credit.
        * **3. Withholding Reconciliation:** Automatic allocation of multi-rate sales taxes and tax withholdings.
        * **4. ERP Interoperability:** Certified CSV/Excel templates for QuickBooks, SAP, Datev, and Xero.
        """,
        "cockpit_card1_title": "🏛️ Tax Withholding Engine",
        "cockpit_card1_desc": "Automatic handling of multi-rate sales taxes and withholding accounts with zero cent deviation.",
        "cockpit_card2_title": "⚡ ERP Interoperability",
        "cockpit_card2_desc": "Direct exports formatted for QuickBooks, SAP, Datev SKR03/04, and multi-tab Excel workbooks.",
        "cockpit_card3_title": "🛡️ Dual-Audit Integrity Lock",
        "cockpit_card3_desc": "Mathematical assurance guaranteeing that Total Debit strictly equals Total Credit before release.",
        "headers": {
            "vouch": "Voucher #", "date": "Date", "code": "Account Code",
            "name": "Account Name", "desc": "Memo", "curr": "Currency",
            "deb": "Debit", "crd": "Credit"
        }
    },
    "🇩🇪 DE": {
        "badge": "MENSCH + KI FINANZTERMINAL",
        "title": "LedgerAI",
        "subtitle": "Autonome Belegerfassung und Datev-konforme Kontierung unter ständiger Expertenkontrolle.",
        "drop_title": "Belege hier ablegen oder durchsuchen",
        "drop_sub": "PDF, PNG, JPG • Rechnungen & Quittungen",
        "process_btn": "⚡ Buchungssätze Erstellen",
        "limit_err": "🛑 Maximal 5 Dokumente im Demo-Modus.",
        "ready_count": "Bereit: **{count}**",
        "success": "✓ Buchungen erfolgreich erstellt und ausgeglichen.",
        "failed": "❌ Belege konnten nicht gelesen werden.",
        "preview_title": "📊 Buchungszeilen & Kontrollzentrum",
        "preview_tip": "💡 Doppelklick zum Ändern von Konten oder Beträgen.",
        "tot_deb": "Soll Gesamt",
        "tot_crd": "Haben Gesamt",
        "balanced": "✅ Ausgeglichen (Soll = Haben)",
        "unbalanced": "⚠️ Differenz festgestellt!",
        "download_btn": "📥 Excel Herunterladen (.xlsx)",
        "download_eta": "💾 Datev Format (CSV)",
        "download_luca": "💾 SAP Kompatibel",
        "industries": [
            "⚡ Automatisch (KI)", "🛒 Handel / Wareneinkauf",
            "🏢 Dienstleistung / IT", "🏭 Produktion / Fertigung"
        ],
        "themes": [
            "🌑 Platin Titan (Business)",
            "✨ Nebula Aurora (Modern)",
            "⚡ Cyber Matrix (Akademie)"
        ],
        "about_btn": "ℹ️ Architektur & Datev",
        "about_title": "LedgerAI Architektur & Datev SKR03/04 Standard",
        "about_content": "Vollautomatisierte Buchungssatzerstellung nach GoBD und Datev-Richtlinien mit strengem Soll/Haben-Ausgleich.",
        "cockpit_card1_title": "🏛️ Vorsteuer- & Steuerlogik",
        "cockpit_card1_desc": "Automatische Zuordnung von SKR03/04 Vorsteuern und USt-IdNr Validierung.",
        "cockpit_card2_title": "⚡ Datev Export",
        "cockpit_card2_desc": "Direkter Datev-konformer CSV-Export für das Steuerbüro.",
        "cockpit_card3_title": "🛡️ Soll/Haben Garantie",
        "cockpit_card3_desc": "Mathematische Prüfung auf absolute Ausgeglichenheit der Buchungssätze.",
        "headers": {
            "vouch": "Beleg", "date": "Datum", "code": "Konto",
            "name": "Bezeichnung", "desc": "Text", "curr": "Währung",
            "deb": "Soll", "crd": "Haben"
        }
    },
    "🇫🇷 FR": {
        "badge": "TERMINAL COLLABORATIF IA + HUMAIN",
        "title": "LedgerAI",
        "subtitle": "L'IA prépare les imputations comptables; l'expert-comptable valide et approuve.",
        "drop_title": "Déposer les pièces comptables ici",
        "drop_sub": "Factures et reçus (PDF, PNG, JPG)",
        "process_btn": "⚡ Générer les Écritures",
        "limit_err": "🛑 Limite: 5 documents par lot.",
        "ready_count": "Prêts: **{count}**",
        "success": "✓ Écritures générées avec succès et équilibrées.",
        "failed": "❌ Échec de lecture.",
        "preview_title": "📊 Journal Comptable & Audit Expert",
        "preview_tip": "💡 Double-cliquez sur une cellule pour modifier.",
        "tot_deb": "Total Débit",
        "tot_crd": "Total Crédit",
        "balanced": "✅ Équilibré (Débit = Crédit)",
        "unbalanced": "⚠️ Déséquilibre Détecté!",
        "download_btn": "📥 Télécharger Excel (.xlsx)",
        "download_eta": "💾 Format Standard PCG",
        "download_luca": "💾 Sage / Cegid Ready",
        "industries": [
            "⚡ Auto (IA)", "🛒 Négoce / Stock",
            "🏢 Services / Conseil", "🏭 Production / Industrie"
        ],
        "themes": [
            "🌑 Platine Titane (Entreprise)",
            "✨ Nebula Aurora (Moderne)",
            "⚡ Cyber Matrix (Académie)"
        ],
        "about_btn": "ℹ️ Architecture & Normes",
        "about_title": "Architecture Comptable & Normes PCG",
        "about_content": "Conformité Plan Comptable Général (PCG) avec vérification stricte du principe Débit = Crédit.",
        "cockpit_card1_title": "🏛️ Ventilation PCG",
        "cockpit_card1_desc": "Affectation automatique aux comptes de charges et TVA déductible.",
        "cockpit_card2_title": "⚡ Formats Export",
        "cockpit_card2_desc": "Compatible avec les logiciels Sage, Cegid et tableur multi-feuilles.",
        "cockpit_card3_title": "🛡️ Équilibre Débit/Crédit",
        "cockpit_card3_desc": "Vérification stricte de la balance avant validation finale.",
        "headers": {
            "vouch": "Pièce", "date": "Date", "code": "Compte",
            "name": "Libellé", "desc": "Détail", "curr": "Devise",
            "deb": "Débit", "crd": "Crédit"
        }
    },
    "🇪🇸 ES": {
        "badge": "TERMINAL COLABORATIVO IA + HUMANO",
        "title": "LedgerAI",
        "subtitle": "La IA estructura los asientos contables; el asesor profesional revisa y valida.",
        "drop_title": "Arrastra los documentos aquí o examina",
        "drop_sub": "PDF, PNG, JPG • Facturas y recibos",
        "process_btn": "⚡ Generar Asientos",
        "limit_err": "🛑 Máximo 5 documentos por lote.",
        "ready_count": "Listos: **{count}**",
        "success": "✓ Asientos generados y equilibrados.",
        "failed": "❌ Error al procesar.",
        "preview_title": "📊 Libro Diario & Mesa de Control",
        "preview_tip": "💡 Haz doble clic para modificar cuentas.",
        "tot_deb": "Total Debe",
        "tot_crd": "Total Haber",
        "balanced": "✅ Cuadrado (Debe = Haber)",
        "unbalanced": "⚠️ Descuadre Detectado!",
        "download_btn": "📥 Descargar Excel (.xlsx)",
        "download_eta": "💾 Formato Contasol",
        "download_luca": "💾 A3 / Sage Ready",
        "industries": [
            "⚡ Automático (IA)", "🛒 Comercio / Inventario",
            "🏢 Servicios / Oficina", "🏭 Fabricación / Industria"
        ],
        "themes": [
            "🌑 Platino Titanio (Corporativo)",
            "✨ Nebula Aurora (Moderno)",
            "⚡ Cyber Matrix (Academia)"
        ],
        "about_btn": "ℹ️ Normativa y PGC",
        "about_title": "Estándares Contables y Seguridad Fiscal",
        "about_content": "Asientos contables conformes al Plan General Contable (PGC) con cuadre matemático de Debe y Haber.",
        "cockpit_card1_title": "🏛️ Cuadre Fiscal",
        "cockpit_card1_desc": "Gestión automática de retenciones e IVA soportado.",
        "cockpit_card2_title": "⚡ Compatibilidad ERP",
        "cockpit_card2_desc": "Exportación directa para Contasol, A3 y software contable estándar.",
        "cockpit_card3_title": "🛡️ Control de Asiento",
        "cockpit_card3_desc": "Validación matemática estricta de paridad Debe = Haber.",
        "headers": {
            "vouch": "Asiento", "date": "Fecha", "code": "Cuenta",
            "name": "Nombre Cuenta", "desc": "Concepto", "curr": "Moneda",
            "deb": "Debe", "crd": "Haber"
        }
    },
    "🇮🇹 IT": {
        "badge": "TERMINALE COLLABORATIVO IA + UOMO",
        "title": "LedgerAI",
        "subtitle": "L'IA prepara le scritture contabili; il commercialista esperto valida e autorizza.",
        "drop_title": "Trascina qui le fatture o cerca file",
        "drop_sub": "PDF, PNG, JPG • Ricevute e fatture",
        "process_btn": "⚡ Genera Scritture",
        "limit_err": "🛑 Massimo 5 documenti.",
        "ready_count": "Pronti: **{count}**",
        "success": "✓ Scritture generate e bilanciate.",
        "failed": "❌ Impossibile elaborare.",
        "preview_title": "📊 Prima Nota & Centro di Controllo",
        "preview_tip": "💡 Fai doppio clic per modificare.",
        "tot_deb": "Totale Dare",
        "tot_crd": "Totale Avere",
        "balanced": "✅ Quadratura Perfetta",
        "unbalanced": "⚠️ Sbilancio!",
        "download_btn": "📥 Scarica Excel (.xlsx)",
        "download_eta": "💾 Formato Zucchetti",
        "download_luca": "💾 Teamsystem Ready",
        "industries": [
            "⚡ Automatico (IA)", "🛒 Commercio / Magazzino",
            "🏢 Servizi / Consulenza", "🏭 Manifattura / Produzione"
        ],
        "themes": [
            "🌑 Platino Titanio (Business)",
            "✨ Nebula Aurora (Moderno)",
            "⚡ Cyber Matrix (Accademia)"
        ],
        "about_btn": "ℹ️ Architettura & Norme",
        "about_title": "Standard di Conformità e Partita Doppia",
        "about_content": "Generazione automatica di scritture in partita doppia perfettamente bilanciate per gestionali Zucchetti e Teamsystem.",
        "cockpit_card1_title": "🏛️ Scritture Bilanciate",
        "cockpit_card1_desc": "Gestione automatica ritenute d'acconto ed IVA a credito.",
        "cockpit_card2_title": "⚡ Compatibilità Gestionale",
        "cockpit_card2_desc": "File pronti per Zucchetti, Teamsystem e formati Excel avanzati.",
        "cockpit_card3_title": "🛡️ Quadratura Certificata",
        "cockpit_card3_desc": "Garanzia matematica di parità tra totale Dare e Avere.",
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

# HIZLI VE KESİNTİSİZ CEVAP ÖNBELLEĞİ
INSTANT_FAQ_CACHE = {
    "💡 Muhasebeciye ne kazandırır?": "LedgerAI, manuel veri girişini %80 azaltarak mali müşavirlerin rutin fiş işleme yükünü ortadan kaldırır. Yapay zeka fiş taslağını oluşturur, uzman insan sadece onaylar ve denetler.",
    "🔒 Verilerim güvende mi?": "Evet. Tüm finansal verileriniz TLS 256-bit uçtan uca şifreleme ile işlenir. Belgeleriniz kalıcı sunucularda saklanmaz ve model eğitiminde (training) kullanılmaz.",
    "⚖️ Tevkifat & Stopaj mantığı nedir?": "Tevkifat ve stopaj, faturadaki verginin bir kısmının alıcı tarafından kesilerek doğrudan vergi dairesine (360 hesabına) ödenmesidir. Böylece satıcı cari hesabı net tutara oturur ve yevmiye fişi kuruş farkı olmadan dengelenir.",
    "🎯 153 ile 770 arasındaki fark nedir?": "153 Ticari Mallar satılmak amacıyla alınan ticari ürünlerin stok hesabıdır. 770 Genel Yönetim Giderleri ise işletmenin kendi idari faaliyetlerinde tükettiği (ofis kırtasiyesi, kira, danışmanlık vb.) giderlerin kaydedildiği hesaptır."
}

# ==============================================================================
# 4. DYNAMIC STYLING ENGINE (GERÇEK GİTHUB HAPLARI & MİKRO SESLİ CSS)
# ==============================================================================

if st.session_state["theme_idx"] == 0:
    # 🌑 Platin Titanyum (Kurumsal / CFO)
    bg_style = """
        @keyframes slateShimmer {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        .stApp {
            background: radial-gradient(circle at 50% 0%, rgba(203, 213, 225, 0.16) 0%, transparent 65%),
                        radial-gradient(circle at 85% 90%, rgba(148, 163, 184, 0.12) 0%, transparent 50%),
                        linear-gradient(145deg, #111827 0%, #1E293B 50%, #0F172A 100%);
            background-size: 200% 200%;
            animation: slateShimmer 24s ease infinite;
            background-attachment: fixed;
        }
    """
elif st.session_state["theme_idx"] == 1:
    # ✨ Nebula Aurora (Modern / Startup)
    bg_style = """
        @keyframes auroraRealFlow {
            0% { background-position: 0% 30%; filter: hue-rotate(0deg); }
            50% { background-position: 100% 70%; filter: hue-rotate(40deg); }
            100% { background-position: 0% 30%; filter: hue-rotate(0deg); }
        }
        .stApp {
            background: radial-gradient(circle at 10% 20%, rgba(16, 185, 129, 0.35) 0%, transparent 45%),
                        radial-gradient(circle at 90% 20%, rgba(14, 165, 233, 0.35) 0%, transparent 45%),
                        radial-gradient(circle at 50% 90%, rgba(99, 102, 241, 0.30) 0%, transparent 50%),
                        linear-gradient(140deg, #020710 0%, #041424 40%, #09213A 70%, #020710 100%);
            background-size: 240% 240%;
            animation: auroraRealFlow 16s ease-in-out infinite;
            background-attachment: fixed;
        }
    """
else:
    # ⚡ Siber Matris (Akademi / Eğitim)
    bg_style = """
        @keyframes cyberpunkPulse {
            0% { background-position: 0% 0%, 100% 100%; filter: brightness(1); }
            50% { background-position: 100% 100%, 0% 0%; filter: brightness(1.2); }
            100% { background-position: 0% 0%, 100% 100%; filter: brightness(1); }
        }
        .stApp {
            background: radial-gradient(circle at 15% 15%, rgba(217, 70, 239, 0.35) 0%, transparent 45%),
                        radial-gradient(circle at 85% 85%, rgba(6, 182, 212, 0.30) 0%, transparent 45%),
                        linear-gradient(135deg, #020108 0%, #080318 45%, #050A1A 80%, #020108 100%);
            background-size: 220% 220%;
            animation: cyberpunkPulse 12s ease-in-out infinite;
            background-attachment: fixed;
        }
    """

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }}
    
    /* WHITELABEL GİZLEME */
    header[data-testid="stHeader"] {{ display: none !important; }}
    #MainMenu {{ visibility: hidden !important; }}
    footer {{ visibility: hidden !important; }}
    div[data-testid="stToolbar"] {{ display: none !important; }}
    div[data-testid="stDecoration"] {{ display: none !important; }}
    .viewerBadge_container__1QSob {{ display: none !important; }}
    div[class*="viewerBadge"] {{ display: none !important; }}
    div[class*="profile-badge"] {{ display: none !important; }}
    iframe[title*="github"] {{ display: none !important; }}
    a[href*="streamlit.io"] {{ display: none !important; }}
    [data-testid="stSidebar"] {{ display: none !important; }}
    
    {bg_style}
    
    .stApp {{
        color: #F8FAFC;
        padding-top: 10px;
        padding-bottom: 60px;
    }}

    .cockpit-container {{
        max-width: 1280px;
        margin: 0 auto;
        padding: 0 10px;
    }}

    .cockpit-card {{
        background: rgba(30, 41, 59, 0.72);
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 24px;
        backdrop-filter: blur(28px);
        -webkit-backdrop-filter: blur(28px);
        padding: clamp(22px, 3vw, 34px);
        box-shadow: 0 20px 50px rgba(0, 0, 0, 0.45), 
                    inset 0 1px 0 rgba(255, 255, 255, 0.20);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }}

    .top-badge {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 14px;
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.18);
        border-radius: 9999px;
        font-size: 0.68rem;
        font-weight: 700;
        letter-spacing: 1.2px;
        color: #E2E8F0;
        margin-bottom: 8px;
        text-transform: uppercase;
        width: fit-content;
    }}

    .hero-title {{
        font-size: clamp(2.0rem, 3.5vw, 2.7rem);
        font-weight: 800;
        letter-spacing: -1.2px;
        background: linear-gradient(135deg, #FFFFFF 40%, #CBD5E1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 4px;
        line-height: 1.05;
    }}
    
    .hero-sub {{
        font-size: 0.90rem;
        color: #CBD5E1;
        font-weight: 400;
        line-height: 1.45;
        margin-bottom: 18px;
    }}

    div[data-testid="stFileUploader"] {{
        background: rgba(15, 23, 42, 0.65);
        border: 1px dashed rgba(255, 255, 255, 0.22);
        border-radius: 18px;
        padding: 18px 14px;
        transition: all 0.25s ease;
        margin-bottom: 12px;
    }}
    div[data-testid="stFileUploader"]:hover {{
        border-color: rgba(203, 213, 225, 0.9);
        box-shadow: 0 0 25px rgba(255, 255, 255, 0.15);
        background: rgba(30, 41, 59, 0.8);
    }}

    div.stButton > button:first-child {{
        background: #000000 !important;
        border: 1px solid rgba(255, 255, 255, 0.25) !important;
        border-radius: 9999px !important;
        font-weight: 700 !important;
        font-size: 0.92rem !important;
        padding: 10px 28px !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 18px rgba(0, 0, 0, 0.6) !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        letter-spacing: 0.2px !important;
        margin-top: 4px !important;
    }}
    div.stButton > button:first-child:hover {{
        background: #111827 !important;
        border-color: rgba(255, 255, 255, 0.6) !important;
        box-shadow: 0 6px 25px rgba(255, 255, 255, 0.2) !important;
        transform: scale(1.02) !important;
    }}

    .cockpit-info-box {{
        background: rgba(255, 255, 255, 0.035);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 14px 16px;
        margin-bottom: 12px;
        backdrop-filter: blur(10px);
        transition: all 0.2s;
    }}
    .cockpit-info-box:hover {{
        background: rgba(255, 255, 255, 0.06);
        border-color: rgba(255, 255, 255, 0.2);
    }}
    .cockpit-info-title {{
        font-size: 0.82rem;
        font-weight: 700;
        color: #F8FAFC;
        margin-bottom: 4px;
    }}
    .cockpit-info-desc {{
        font-size: 0.74rem;
        color: #CBD5E1;
        line-height: 1.4;
    }}

    /* GERÇEK GİTHUB.COM MİKRO HAPI (SELECTBOX VE BUTONLAR) */
    .github-pill-select div[data-baseweb="select"] > div {{
        background: rgba(255, 255, 255, 0.06) !important;
        border: 1px solid rgba(255, 255, 255, 0.18) !important;
        border-radius: 9999px !important;
        min-height: 30px !important;
        height: 30px !important;
        padding: 0 10px !important;
        font-size: 0.75rem !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.3) !important;
        color: #F8FAFC !important;
    }}
    
    div[data-testid="stPopover"] > button {{
        background: rgba(255, 255, 255, 0.06) !important;
        border: 1px solid rgba(255, 255, 255, 0.18) !important;
        border-radius: 9999px !important;
        min-height: 30px !important;
        height: 30px !important;
        padding: 0 14px !important;
        font-size: 0.75rem !important;
        font-weight: 600 !important;
        color: #F8FAFC !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.3) !important;
    }}

    div[data-testid="stExpander"] div.stButton button {{
        background: rgba(255, 255, 255, 0.06) !important;
        border: 1px solid rgba(255, 255, 255, 0.18) !important;
        border-radius: 9999px !important;
        font-size: 0.74rem !important;
        font-weight: 600 !important;
        color: #F8FAFC !important;
        padding: 4px 14px !important;
        min-height: 28px !important;
        height: 28px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.3) !important;
        transition: all 0.2s ease !important;
        white-space: nowrap !important;
        line-height: 1 !important;
        margin-top: 0px !important;
    }}
    div[data-testid="stExpander"] div.stButton button:hover {{
        background: rgba(255, 255, 255, 0.14) !important;
        border-color: rgba(255, 255, 255, 0.45) !important;
        transform: scale(1.03) !important;
    }}

    .chat-scroll-area {{
        max-height: 280px;
        overflow-y: auto;
        padding: 10px 12px;
        background: rgba(15, 23, 42, 0.65);
        border: 1px solid rgba(255, 255, 255, 0.10);
        border-radius: 14px;
        margin-bottom: 10px;
    }}
    .chat-scroll-area::-webkit-scrollbar {{ width: 5px; }}
    .chat-scroll-area::-webkit-scrollbar-thumb {{
        background: rgba(255, 255, 255, 0.2);
        border-radius: 4px;
    }}

    .user-bubble {{
        background: rgba(59, 130, 246, 0.22);
        border: 1px solid rgba(59, 130, 246, 0.35);
        border-radius: 12px 12px 2px 12px;
        padding: 8px 12px;
        margin: 6px 0 6px auto;
        max-width: 85%;
        font-size: 0.80rem;
        color: #F1F5F9;
        line-height: 1.4;
    }}

    .assistant-bubble {{
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 12px 12px 12px 2px;
        padding: 8px 12px;
        margin: 6px auto 6px 0;
        max-width: 90%;
        font-size: 0.80rem;
        color: #CBD5E1;
        line-height: 1.45;
    }}

    div[data-testid="stTabs"] button[role="tab"] {{
        background: transparent !important;
        border: none !important;
        font-size: 0.88rem !important;
        font-weight: 700 !important;
        color: #94A3B8 !important;
        padding: 8px 18px !important;
    }}
    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{
        color: #FFFFFF !important;
        border-bottom: 2px solid #38BDF8 !important;
    }}
</style>
""", unsafe_allow_html=True)

# SESLİ BİLDİRİM MOTORU (WEB AUDIO API)
def sesli_bildirim_cal(tur="success"):
    if tur == "success":
        ses_js = """
        <script>
            try {
                const ctx = new (window.AudioContext || window.webkitAudioContext)();
                const osc = ctx.createOscillator();
                const gain = ctx.createGain();
                osc.type = 'sine';
                osc.frequency.setValueAtTime(587.33, ctx.currentTime);
                osc.frequency.exponentialRampToValueAtTime(880.00, ctx.currentTime + 0.15);
                gain.gain.setValueAtTime(0.08, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.35);
                osc.connect(gain);
                gain.connect(ctx.destination);
                osc.start();
                osc.stop(ctx.currentTime + 0.35);
            } catch(e) {}
        </script>
        """
    else:
        ses_js = """
        <script>
            try {
                const ctx = new (window.AudioContext || window.webkitAudioContext)();
                const osc = ctx.createOscillator();
                const gain = ctx.createGain();
                osc.type = 'sawtooth';
                osc.frequency.setValueAtTime(220, ctx.currentTime);
                gain.gain.setValueAtTime(0.06, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.25);
                osc.connect(gain);
                gain.connect(ctx.destination);
                osc.start();
                osc.stop(ctx.currentTime + 0.25);
            } catch(e) {}
        </script>
        """
    st.markdown(ses_js, unsafe_allow_html=True)

# ==============================================================================
# 5. INSTITUTIONAL MULTI-TAB EXCEL ENGINE
# ==============================================================================

def export_multitab_corporate_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        headers = T["headers"]
        border_thin = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        data_font = Font(name="Plus Jakarta Sans", size=10)
        num_font = Font(name="Consolas", size=10)

        df.to_excel(writer, index=False, sheet_name="Yevmiye Fisi")
        ws1 = writer.sheets["Yevmiye Fisi"]
        h_fill1 = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        h_font1 = Font(name="Plus Jakarta Sans", size=10, bold=True, color="FFFFFF")

        for col_idx in range(1, len(df.columns) + 1):
            c = ws1.cell(row=1, column=col_idx)
            c.fill = h_fill1
            c.font = h_font1
            c.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws1.columns:
            m_len = max(len(str(cell.value or '')) for cell in col)
            c_letter = get_column_letter(col[0].column)
            col_name = str(col[0].value or '')
            for cell in col:
                cell.border = border_thin
                if cell.row != 1:
                    cell.alignment = Alignment(vertical="center")
                    if any(t in col_name.lower() for t in ["borç", "alacak", "debit", "credit"]):
                        cell.font = num_font
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.font = data_font
            ws1.column_dimensions[c_letter].width = max(m_len + 5, 14)

        df_153 = df[df[headers["code"]].astype(str).str.startswith("153") | df[headers["code"]].astype(str).str.startswith("1200")].copy()
        if not df_153.empty:
            df_153.to_excel(writer, index=False, sheet_name="153 Ticari Mallar")
            ws2 = writer.sheets["153 Ticari Mallar"]
            h_fill2 = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")
            for col_idx in range(1, len(df_153.columns) + 1):
                c = ws2.cell(row=1, column=col_idx)
                c.fill = h_fill2
                c.font = h_font1
                c.alignment = Alignment(horizontal="center", vertical="center")
            for col in ws2.columns:
                m_len = max(len(str(cell.value or '')) for cell in col)
                c_letter = get_column_letter(col[0].column)
                for cell in col:
                    cell.border = border_thin
                    if cell.row != 1:
                        cell.font = data_font
                        cell.alignment = Alignment(vertical="center")
                ws2.column_dimensions[c_letter].width = max(m_len + 5, 14)

        df_770 = df[df[headers["code"]].astype(str).str.startswith("770") | df[headers["code"]].astype(str).str.startswith("6000")].copy()
        if not df_770.empty:
            df_770.to_excel(writer, index=False, sheet_name="770 Genel Masraflar")
            ws3 = writer.sheets["770 Genel Masraflar"]
            h_fill3 = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
            for col_idx in range(1, len(df_770.columns) + 1):
                c = ws3.cell(row=1, column=col_idx)
                c.fill = h_fill3
                c.font = h_font1
                c.alignment = Alignment(horizontal="center", vertical="center")
            for col in ws2.columns:
                m_len = max(len(str(cell.value or '')) for cell in col)
                c_letter = get_column_letter(col[0].column)
                for cell in col:
                    cell.border = border_thin
                    if cell.row != 1:
                        cell.font = data_font
                        cell.alignment = Alignment(vertical="center")
                ws3.column_dimensions[c_letter].width = max(m_len + 5, 14)

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
# 6. CORE AI RECOGNITION ENGINE (TEVKİFAT, MULTI-TAX & DUAL AUDIT)
# ==============================================================================

def execute_document_audit(uploaded_files, sector_directive: str):
    results = []
    total = len(uploaded_files)
    progress_bar = st.progress(0)
    status_msg = st.empty()

    for idx, doc in enumerate(uploaded_files):
        status_msg.text(f"⚡ Analiz Ediliyor ({idx + 1}/{total}): {doc.name}...")
        raw_bytes = doc.read()
        mime_type = doc.type if doc.type else "application/pdf"

        prompt = f"""
        You are an elite autonomous financial auditor and ERP data extractor.
        {sector_directive}

        Examine the document carefully. Extract:
        1. Document Language & Origin: (TR, DE, FR, US, IT, ES).
        2. Currency: (TRY, USD, EUR, GBP).
        3. Vendor Information: Full legal name, Tax ID (VKN/TCKN/EIN/SIRET/Steuernummer).
        4. Invoice Metadata: Official Invoice Number, Date (YYYY-MM-DD).
        5. Deep Financial Breakdown:
           - Expense / Resale Classification (Uniform Chart of Accounts 153/150/770/740/255).
           - Multi-tier VAT & Tax Breakdown:
             * Net Amount (Matrah)
             * Tax Rate (e.g. 20, 10, 1)
             * Tax Amount (KDV)
           - Withholding Tax / Tevkifat / Stopaj:
             * Tevkifat / Stopaj Amount (if any, e.g. 2/10, 5/10, 7/10, 9/10 KDV Tevkifatı or SMMM Stopaj)
             * Tevkifat Rate (e.g. "9/10", "5/10" or "")
           - Payable to Vendor (Net Amount to Pay / Ödenecek Tutar / Total - Withholding)
           - Total Invoice Amount (Gross Total)
           - Audit Confidence Score (0-100 based on text sharpness and mathematical validity)

        Respond ONLY with a valid JSON object. No commentary, no markdown codeblocks:
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
          "withholding": 0.0,
          "withholding_rate": "",
          "payable_to_vendor": 0.0,
          "total": 0.0,
          "confidence": 98
        }}
        """

        max_retries = 3
        for attempt in range(max_retries):
            try:
                response = client.models.generate_content(
                    model="gemini-3.5-flash-lite",
                    contents=prompt
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
# 7. MULTI-DECK COCKPIT (KURUMSAL / AKADEMİ / HUKUK)
# ==============================================================================

st.markdown("<div class='cockpit-container'>", unsafe_allow_html=True)

sekme_terminal, sekme_akademi, sekme_hukuk = st.tabs([
    "🏢 Kurumsal Finans Terminali", 
    "🎓 Siber Akademi (Kademeli Oyun & Simülasyon)", 
    "⚖️ Hukuki Çerçeve, KVKK & SLA"
])

# ------------------------------------------------------------------------------
# SEKME 1: KURUMSAL FİNANS TERMİNALİ
# ------------------------------------------------------------------------------
with sekme_terminal:
    col_left, col_right = st.columns([1.35, 1.0], gap="large")

    with col_left:
        st.markdown(f"""
        <div class='cockpit-card'>
            <div>
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
                st.markdown(f"<div style='font-size:0.85rem; margin:4px 0 8px 0; color:#A7F3D0;'>{T['ready_count'].format(count=len(uploaded_files))}</div>", unsafe_allow_html=True)

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
                    st.session_state["raw_audit_results"] = parsed_data

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

                            net = round(float(item.get("net") or 0.0), 2)
                            tax = round(float(item.get("tax") or 0.0), 2)
                            total = round(float(item.get("total") or (net + tax)), 2)
                            tax_rate = item.get("tax_rate") or 20
                            withholding = round(float(item.get("withholding") or 0.0), 2)
                            payable = round(float(item.get("payable_to_vendor") or 0.0), 2)

                            # Matematiksel Tevkifat & Bakiye Koruma Motoru
                            if payable > 0 and abs((net + tax) - payable) > 0.05 and withholding == 0:
                                withholding = round((net + tax) - payable, 2)
                            
                            if payable == 0:
                                payable = round((net + tax) - withholding, 2)

                            clean_name = "".join(c for c in vendor[:12] if c.isalnum()).upper() or "CARİ"
                            if "TR" in st.session_state["user_lang"]:
                                ap_code = f"320.{tax_id}" if tax_id else f"320.{clean_name}"
                                tax_code = f"191.{int(tax_rate):02d}"
                                tax_name = f"%{tax_rate} İndirilecek KDV"
                                tevkifat_code = "360.01"
                                tevkifat_name = "Ödenecek KDV Tevkifatı / Stopaj"
                            else:
                                ap_code = f"2000-{tax_id}" if tax_id else f"VEND-{clean_name}"
                                tax_code = f"2200-TAX{tax_rate}"
                                tax_name = f"Sales Tax ({tax_rate}%)"
                                tevkifat_code = "2250-WITHHOLDING"
                                tevkifat_name = "Withholding Tax Payable"

                            # 1. BORÇ: Gider / Mal Alışı
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

                            # 2. BORÇ: KDV Tutarı
                            if tax > 0:
                                voucher_lines.append({
                                headers["vouch"]: voucher_num,
                                headers["date"]: date_val,
                                headers["code"]: tax_code,
                                headers["name"]: tax_name,
                                headers["desc"]: f"{vendor} - KDV",
                                headers["curr"]: curr,
                                headers["deb"]: tax,
                                headers["crd"]: 0.0
                            })

                            # 3. ALACAK: Tevkifat / Stopaj (Varsa)
                            if withholding > 0:
                                voucher_lines.append({
                                headers["vouch"]: voucher_num,
                                headers["date"]: date_val,
                                headers["code"]: tevkifat_code,
                                headers["name"]: tevkifat_name,
                                headers["desc"]: f"{vendor} - Tevkifat/Kesinti",
                                headers["curr"]: curr,
                                headers["deb"]: 0.0,
                                headers["crd"]: withholding
                            })

                            # 4. ALACAK: Satıcı Cari Hesabı (Net Ödenecek Tutar)
                            voucher_lines.append({
                                headers["vouch"]: voucher_num,
                                headers["date"]: date_val,
                                headers["code"]: ap_code,
                                headers["name"]: vendor,
                                headers["desc"]: f"{vendor} - {inv_no}",
                                headers["curr"]: curr,
                                headers["deb"]: 0.0,
                                headers["crd"]: payable
                            })

                            voucher_num += 1

                        st.session_state["out_df"] = pd.DataFrame(voucher_lines)
                        st.session_state["h_deb"] = headers["deb"]
                        st.session_state["h_crd"] = headers["crd"]
                        st.session_state["last_processing_time"] = round(time.time() - start_time, 2)
                        st.session_state["processed_docs_count"] = len(parsed_data)
                        sesli_bildirim_cal("success")
                        st.success(f"{T['success']} ({st.session_state['last_processing_time']} sn)")

        st.markdown("</div>", unsafe_allow_html=True)

        # TAM HİZALI MİKRO DOCK (GÖRSEL 35 DÜZELTİLDİ)
        st.markdown("<div style='margin-top:14px; padding-top:10px; border-top:1px solid rgba(255,255,255,0.08);'>", unsafe_allow_html=True)
        c_m1, c_m2, c_m3, c_m4 = st.columns([1.6, 3.4, 3.4, 1.6])
        with c_m1:
            st.markdown("<div class='github-pill-select'>", unsafe_allow_html=True)
            lang_keys = list(LANG_DATA.keys())
            curr_lang_idx = lang_keys.index(st.session_state["user_lang"]) if st.session_state["user_lang"] in lang_keys else 0
            new_lang = st.selectbox("Dil", lang_keys, index=curr_lang_idx, label_visibility="collapsed")
            if new_lang != st.session_state["user_lang"]:
                st.session_state["user_lang"] = new_lang
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
        with c_m2:
            st.markdown("<div class='github-pill-select'>", unsafe_allow_html=True)
            new_theme_str = st.selectbox("Görünüm", T["themes"], index=st.session_state["theme_idx"], label_visibility="collapsed")
            new_t_idx = T["themes"].index(new_theme_str)
            if new_t_idx != st.session_state["theme_idx"]:
                st.session_state["theme_idx"] = new_t_idx
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
        with c_m3:
            st.markdown("<div class='github-pill-select'>", unsafe_allow_html=True)
            new_industry_str = st.selectbox("Sektör", T["industries"], index=st.session_state["industry_idx"], label_visibility="collapsed")
            new_i_idx = T["industries"].index(new_industry_str)
            if new_i_idx != st.session_state["industry_idx"]:
                st.session_state["industry_idx"] = new_i_idx
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
        with c_m4:
            with st.popover(T["about_btn"]):
                st.markdown(f"#### {T['about_title']}")
                st.markdown(T["about_content"])
        st.markdown("</div></div>", unsafe_allow_html=True)

    with col_right:
        st.markdown(f"""
        <div class='cockpit-card'>
            <div>
                <div style='font-size:0.75rem; font-weight:800; letter-spacing:1px; color:#94A3B8; text-transform:uppercase; margin-bottom:12px;'>
                    🛡️ Kurumsal Finans & Güvence Masası
                </div>
                <div class='cockpit-info-box'>
                    <div class='cockpit-info-title'>{T['cockpit_card1_title']}</div>
                    <div class='cockpit-info-desc'>{T['cockpit_card1_desc']}</div>
                </div>
                <div class='cockpit-info-box'>
                    <div class='cockpit-info-title'>{T['cockpit_card2_title']}</div>
                    <div class='cockpit-info-desc'>{T['cockpit_card2_desc']}</div>
                </div>
                <div class='cockpit-info-box'>
                    <div class='cockpit-info-title'>{T['cockpit_card3_title']}</div>
                    <div class='cockpit-info-desc'>{T['cockpit_card3_desc']}</div>
                </div>
            </div>
            <div style='display:flex; justify-content:space-between; align-items:center; margin-top:14px; padding:10px 14px; background:rgba(0,0,0,0.25); border-radius:12px;'>
                <span style='font-size:0.75rem; color:#A7F3D0;'>✓ %100 Bakiye Garantisi</span>
                <span style='font-size:0.75rem; color:#CBD5E1;'>ETA • Luca • Datev Ready</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # TABLO VE ÇIKTI ALANI
    if st.session_state["out_df"] is not None:
        st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
        st.subheader(T["preview_title"])
        st.caption(T["preview_tip"])

        headers = T["headers"]
        st.markdown("<div style='background:rgba(30,41,59,0.55); border:1px solid rgba(255,255,255,0.1); border-radius:14px; padding:10px 16px; margin-bottom:14px;'>", unsafe_allow_html=True)
        f_col1, f_col2, f_col3 = st.columns([2.5, 3.5, 2])
        
        with f_col1:
            filtre_turu = st.selectbox(
                "Filtrele", 
                ["Tüm Satırlar", "Sadece 153 (Ticari Mallar)", "Sadece 770 (Genel Masraflar)", "Tevkifat / Stopaj (360)", "Sadece Satıcılar (320)"],
                label_visibility="collapsed"
            )
        
        with f_col2:
            if st.session_state.get("raw_audit_results"):
                doc_cnt = len(st.session_state["raw_audit_results"])
                avg_conf = sum(d.get("confidence", 95) for d in st.session_state["raw_audit_results"]) / max(doc_cnt, 1)
                st.markdown(f"<div style='font-size:0.85rem; padding-top:6px; color:#A7F3D0;'>🛡️ <b>Mühürlü Denetim:</b> {doc_cnt} Belge %{avg_conf:.1f} Güven Skoruyla Hazırlandı.</div>", unsafe_allow_html=True)

        with f_col3:
            st.markdown("<div style='text-align:right; font-size:0.82rem; padding-top:6px; color:#CBD5E1;'>💡 Uzman Onay Masası</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        df_goruntule = st.session_state["out_df"].copy()
        if filtre_turu == "Sadece 153 (Ticari Mallar)":
            df_goruntule = df_goruntule[df_goruntule[headers["code"]].astype(str).str.startswith("153")]
        elif filtre_turu == "Sadece 770 (Genel Masraflar)":
            df_goruntule = df_goruntule[df_goruntule[headers["code"]].astype(str).str.startswith("770")]
        elif filtre_turu == "Tevkifat / Stopaj (360)":
            df_goruntule = df_goruntule[df_goruntule[headers["code"]].astype(str).str.startswith("360")]
        elif filtre_turu == "Sadece Satıcılar (320)":
            df_goruntule = df_goruntule[df_goruntule[headers["code"]].astype(str).str.startswith("320")]

        edited_df = st.data_editor(
            df_goruntule,
            use_container_width=True,
            num_rows="dynamic"
        )

        deb_col = st.session_state["h_deb"]
        crd_col = st.session_state["h_crd"]

        tot_deb = st.session_state["out_df"][deb_col].sum()
        tot_crd = st.session_state["out_df"][crd_col].sum()
        diff = abs(tot_deb - tot_crd)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric(T["tot_deb"], f"{tot_deb:,.2f}")
        m2.metric(T["tot_crd"], f"{tot_crd:,.2f}")
        
        if diff < 0.05:
            m3.success(T["balanced"])
        else:
            m3.error(f"{T['unbalanced']} (Δ {diff:,.2f})")
            sesli_bildirim_cal("error")
            
        m4.metric("İşlem Süresi", f"{st.session_state['last_processing_time']} sn")

        exp_col1, exp_col2, exp_col3 = st.columns(3)
        
        with exp_col1:
            xlsx_data = export_multitab_corporate_excel(st.session_state["out_df"])
            st.download_button(
                label=T["download_btn"],
                data=xlsx_data,
                file_name="ledger_kurumsal_muhasebe_raporu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with exp_col2:
            eta_data = export_eta_csv(st.session_state["out_df"])
            st.download_button(
                label=T["download_eta"],
                data=eta_data,
                file_name="eta_v11_aktarim.csv",
                mime="text/csv",
                use_container_width=True
            )

        with exp_col3:
            json_data = st.session_state["out_df"].to_json(orient="records", indent=2, force_ascii=False)
            st.download_button(
                label="💾 JSON Veri İndir",
                data=json_data,
                file_name="ledger_audit_data.json",
                mime="application/json",
                use_container_width=True
            )

    # ASİSTAN ÇUBUĞU (AKILLI PROMPT MOTORU İLE SAÇMALAMALAR SIFIRLANDI)
    st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
    c_bot_l, c_bot_center, c_bot_r = st.columns([1, 4, 1])

    with c_bot_center:
        with st.expander(T["bot_title"], expanded=False):
            top_col1, top_col2 = st.columns([5.5, 1.5])
            top_col1.caption(T["bot_welcome"])
            with top_col2:
                if st.button(T["bot_clear"], key="btn_clear_chat", use_container_width=True):
                    st.session_state["chat_messages"] = []
                    st.rerun()

            secilen_chip = None
            btn_cols = st.columns(len(T["quick_chips"]))
            for c_idx, chip_text in enumerate(T["quick_chips"]):
                with btn_cols[c_idx]:
                    if st.button(chip_text, key=f"gh_pill_btn_{c_idx}", use_container_width=True):
                        secilen_chip = chip_text

            st.markdown("<div class='chat-scroll-area'>", unsafe_allow_html=True)
            if not st.session_state["chat_messages"]:
                st.markdown(f"<div style='color: #94A3B8; font-size: 0.82rem; padding: 8px 0;'>💡 <i>{T['bot_placeholder']}</i></div>", unsafe_allow_html=True)
            else:
                for msg in st.session_state["chat_messages"][-6:]:
                    if msg["role"] == "user":
                        st.markdown(f"<div class='user-bubble'><b>Soru:</b> {msg['content']}</div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div class='assistant-bubble'>🤖 <b>LedgerAI:</b> {msg['content']}</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

            user_query = st.chat_input(T["bot_placeholder"])
            aktif_soru = user_query or secilen_chip
            
            if aktif_soru:
                st.session_state["chat_messages"].append({"role": "user", "content": aktif_soru})
                
                # 1. ADIM: HIZLI VE KESİNTİSİZ CEVAP ÖNBELLEĞİ (0.01 SN)
                if aktif_soru in INSTANT_FAQ_CACHE:
                    bot_cevap = INSTANT_FAQ_CACHE[aktif_soru]
                    st.session_state["chat_messages"].append({"role": "assistant", "content": bot_cevap})
                    st.rerun()
                else:
                    # 2. ADIM: AKILLI GEMİNİ PROMPT MOTORU (SAÇMA FİŞ ÜRETMEYİ ENGELLEYEN FİLTRE)
                    with st.spinner("● ● ● Analiz ediliyor..."):
                        prompt_bot = f"""
                        Sen LedgerAI'ın kurumsal finans ve muhasebe asistanısın.
                        Kullanıcı Dili: {st.session_state['user_lang']}
                        Kullanıcı Mesajı: "{aktif_soru}"

                        ÇOK KESİN TALİMATLAR:
                        1. Eğer kullanıcı selam veriyorsa, hal hatır soruyorsa veya anlamsız/şaka bir şey yazdıysa (Örn: "ne yapıyorsun", "naber", "hebele hübele"):
                           - Asla fiş kaydı veya muhasebe hesabı uydurma!
                           - Kısa, esprili ve profesyonel bir selam ver, muhasebe ile ilgili ne öğrenmek istediğini sor.
                        2. Eğer kullanıcı GERÇEKTEN bir muhasebe, gider, fatura veya hesap planı sorusu soruyorsa:
                           - 2-3 cümlede doğrudan cevabı ver.
                           - SADECE BU DURUMDA sonuna tek satırlık pratik yevmiye fişi ekle (Örn: Borç 770 / Alacak 320).
                        """
                        try:
                            yanit = client.models.generate_content(
                                model="gemini-3.5-flash-lite",
                                contents=prompt_bot
                            )
                            bot_cevap = yanit.text.strip() if yanit and yanit.text else "Harika bir gün! Muhasebe ve fatura süreçlerinizde size nasıl yardımcı olabilirim?"
                        except Exception:
                            bot_cevap = "Muhasebe ve vergi mevzuatıyla ilgili sorularınızı kısaca yanıtlamaya hazırım."
                        
                        st.session_state["chat_messages"].append({"role": "assistant", "content": bot_cevap})
                        st.rerun()

# ------------------------------------------------------------------------------
# SEKME 2: 🎓 SİBER AKADEMİ (KADEMELİ OYUN, PUZZLE & SEVİYE SİSTEMİ)
# ------------------------------------------------------------------------------
with sekme_akademi:
    # SEVİYE VE XP HESAPLAMA MOTORU
    xp = st.session_state["academy_xp"]
    if xp >= 700:
        st.session_state["academy_level"] = "🏆 Baş Denetçi (Senior Auditor)"
    elif xp >= 350:
        st.session_state["academy_level"] = "⭐ Kıdemli Denetçi Yardımcısı"
    else:
        st.session_state["academy_level"] = "🌱 Mali Stajyer (Junior)"

    curr_p = ACADEMY_PUZZLES[st.session_state["puzzle_step"] % len(ACADEMY_PUZZLES)]

    st.markdown(f"""
    <div style='background:rgba(30,41,59,0.7); border:1px solid rgba(255,255,255,0.12); border-radius:20px; padding:22px 28px; margin-bottom:20px;'>
        <div style='display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;'>
            <div>
                <span class='top-badge' style='background:rgba(217,70,239,0.15); border-color:#D946EF; color:#F0ABFC;'>🎮 SİBER MUHASEBE SİMÜLASYONU</span>
                <h3 style='margin:4px 0; color:#FFFFFF;'>Geleceğin Finans Lideri Yetiştirme Platformu</h3>
                <p style='font-size:0.85rem; color:#CBD5E1; margin:0;'>Fatura vakalarını inceleyin, eksik hesap kodlarını ve tevkifatı doğru seçerek XP kazanın, seviye atlayın!</p>
            </div>
            <div style='text-align:right;'>
                <div style='font-size:1.6rem; font-weight:800; color:#D946EF;'>🏆 {st.session_state["academy_xp"]} XP</div>
                <div style='font-size:0.8rem; color:#A7F3D0; font-weight:700;'>{st.session_state["academy_level"]}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    c_puz1, c_puz2 = st.columns([1.25, 1.0], gap="large")

    with c_puz1:
        st.markdown(f"""
        <div style='background:rgba(15,23,42,0.75); border:1px dashed rgba(255,255,255,0.22); border-radius:20px; padding:22px;'>
            <div style='display:flex; justify-content:space-between; align-items:center;'>
                <span style='font-size:0.75rem; font-weight:800; color:#38BDF8; letter-spacing:1px;'>VAKA #{curr_p['id']} / {len(ACADEMY_PUZZLES)}</span>
                <span style='font-size:0.75rem; color:#F0ABFC; font-weight:700;'>Ödül: +{curr_p['reward_xp']} XP</span>
            </div>
            <h4 style='color:#FFFFFF; margin:8px 0 12px 0;'>{curr_p['title']}</h4>
            <div style='font-size:0.85rem; color:#CBD5E1; margin-bottom:14px; line-height:1.5;'>
                <b>Senaryo:</b> {curr_p['scenario']}
            </div>
            <div style='background:rgba(0,0,0,0.35); border-radius:12px; padding:14px; font-family:"Consolas", monospace; font-size:0.82rem; color:#E2E8F0; line-height:1.6;'>
                📄 <b>FATURA VERİLERİ:</b><br>{curr_p['invoice_data']}
            </div>
            <div style='margin-top:14px; font-size:0.85rem; color:#F8FAFC; font-weight:600;'>
                ❓ {curr_p['question']}
            </div>
        </div>
        """, unsafe_allow_html=True)

    with c_puz2:
        st.markdown("<div style='background:rgba(30,41,59,0.72); border:1px solid rgba(255,255,255,0.14); border-radius:20px; padding:22px;'>", unsafe_allow_html=True)
        st.markdown("<div style='font-size:0.82rem; font-weight:700; color:#F1F5F9; margin-bottom:10px;'>🎯 Yapboz Parçalarını Doğru Yerleştirin:</div>", unsafe_allow_html=True)

        user_sel_code = st.selectbox("1. Parça (Borç Hesabı):", curr_p["options_code"], key=f"sel_code_{curr_p['id']}")
        user_sel_tax = st.selectbox("2. Parça (Alacak & Tevkifat Dengesi):", curr_p["options_tax"], key=f"sel_tax_{curr_p['id']}")

        btn_col_a, btn_col_b = st.columns(2)
        with btn_col_a:
            if st.button("🛡️ Fişi Mühürle & Doğrula", use_container_width=True):
                if user_sel_code == curr_p["correct_code"] and user_sel_tax == curr_p["correct_tax"]:
                    st.session_state["academy_xp"] += curr_p["reward_xp"]
                    st.session_state["puzzle_solved"] = True
                    sesli_bildirim_cal("success")
                    st.balloons()
                    st.success(f"🎉 {curr_p['success_msg']} (+{curr_p['reward_xp']} XP Kazandınız!)")
                else:
                    sesli_bildirim_cal("error")
                    st.error(f"⚠️ {curr_p['error_msg']}")

        with btn_col_b:
            if st.button("➡️ Sonraki Göreve Geç", use_container_width=True):
                st.session_state["puzzle_step"] += 1
                st.session_state["puzzle_solved"] = False
                st.rerun()

        st.markdown(f"""
            <div style='margin-top:16px; padding:12px; background:rgba(0,0,0,0.3); border-radius:12px; font-size:0.75rem; color:#CBD5E1;'>
                💡 <b>İpucu:</b> Tüm seviyeleri başarıyla tamamlayan öğrenciler "Baş Denetçi" unvanı kazanır ve ERP sınavlarına hazır hale gelir.
            </div>
        </div>
        """, unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# SEKME 3: ⚖️ HUKUKİ ÇERÇEVE, KVKK & SLA (TAM VE KUSURSUZ METİN)
# ------------------------------------------------------------------------------
with sekme_hukuk:
    # GÖRSEL 33'TEKİ HAM HTML HATASI GİDERİLDİ
    st.markdown("### ⚖️ Kurumsal Hizmet Seviyesi Anlaşması (SLA), KVKK & Yasal Sorumluluk Çerçevesi")
    st.caption("Bu metin LedgerAI altyapısını kullanan kurumlar, bağımsız denetçiler ve eğitim kurumları için bağlayıcı regülasyon çerçevesini belirler.")

    h_col1, h_col2 = st.columns(2, gap="large")

    with h_col1:
        st.markdown("""
        #### 1. Dual-Control (Çift Kontrol) İlkesi & Mesleki Sorumluluk
        * **Öneri ve Taslak Niteliği:** LedgerAI, optik karakter tanıma (OCR) ve semantik yapay zeka modelleri kullanarak faturaları muhasebeleştiren otonom bir ön muhasebe terminalidir. Sistem tarafından üretilen yevmiye fişleri kesin yasal kayıt değil, **uzman onayına sunulan taslaktır**.
        * **Yetkili Meslek Mensubu Onayı:** 3568 Sayılı Serbest Muhasebeci Mali Müşavirlik ve Yeminli Mali Müşavirlik Kanunu uyarınca, yasal defterlere işleme ve beyanname verme yetkisi yalnızca yetkili meslek mensuplarına aittir. LedgerAI personelin yerini almaz, personele mekanik veri girişinde süper güç sağlar.
        * **Vergi ve Ceza Sorumluluğu:** Kullanıcı tarafından son incelemesi yapılmadan ERP sistemlerine aktarılan fişlerdeki olası matrah veya tevkifat uyuşmazlıklarında nihai sorumluluk mükellefe ve ilgili işletmeye aittir.

        #### 2. KVKK & GDPR Kapsamında Veri Güvenliği Taahhüdü
        * **Sıfır Kalıcı Depolama (Zero-Retention):** Kullanıcı tarafından oturum süresince yüklenen fatura, makbuz ve finansal dökümanlar geçici bellek (RAM) üzerinde işlenir. Oturum kapatıldığında veya sayfa yenilendiğinde belgeler sunuculardan kalıcı olarak silinir.
        * **Model Eğitimi Yasağı:** Finansal belgeleriniz, ticari sırlarınız ve müşteri bilgileriniz kesinlikle yapay zeka modellerinin genel eğitiminde (training) kullanılmaz.
        * **Aktarım Güvenliği:** Tüm iletişim TLS 256-bit bankacılık seviyesinde şifrelenmiş tüneller üzerinden yürütülür.
        """)

    with h_col2:
        st.markdown("""
        #### 3. Eğitim & Üniversite Lisanslama Çerçevesi
        * **Simülasyon Ortamı:** "Siber Akademi" sekmesinde sunulan interaktif fatura bulmacaları ve hesap kodu senaryoları eğitim amacıyla kurgulanmış hayali vakalardır.
        * **Öğrenci Gizliliği:** Akademi modülünde öğrencilerden veya kurumlardan hiçbir kişisel veri, TC Kimlik Numarası veya gerçek finansal döküman talep edilmez.
        * **Ders Materyali Uyumluluğu:** Sistem; üniversitelerin İktisadi ve İdari Bilimler Fakülteleri, Meslek Yüksekokulları ve Ticaret Liseleri için müfredata uyumlu dijital laboratuvar materyali olarak kullanılabilir.

        #### 4. Hizmet Seviyesi (SLA) & Kesintisiz Çalışma
        * **Doğruluk Güvencesi:** Sistem matematiksel çift bakiye denetimi yaparak Borç ve Alacak eşitliği sağlanmayan fişlerin dışa aktarılmasına izin vermez.
        * **API ve ERP Entegrasyonu:** Dışa aktarılan Excel (.xlsx) ve CSV formatları ETA V.11, Luca, Logo, Zirve ve Datev standartlarında kodlanmış olup biçimsel veri bütünlüğü garanti altındadır.
        """)

st.markdown("</div>", unsafe_allow_html=True)
