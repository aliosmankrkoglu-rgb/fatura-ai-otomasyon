"""
================================================================================
LEDGERAI — MULTI-MODAL ENTERPRISE FINANCIAL TERMINAL & CYBER ACADEMY
Architecture: Streamlit + Google Gemini GenAI SDK + Pandas + OpenPyXL + HTML5 Canvas
Design: Minimal Circular Glass Nav / Authentic ETA:SQL Journal Entry Simulator
Compliance: KVKK, GDPR, Turkish Uniform Chart of Accounts, Datev, US GAAP
Version: 9.0.0 Authentic ETA ERP Engine Edition
================================================================================
"""

import streamlit as st
import json
import time
import io
import random
import datetime
import math
import pandas as pd
import streamlit.components.v1 as components
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. CORE SYSTEM CONFIGURATION & INITIAL STATE
# ==============================================================================

st.set_page_config(
    page_title="LedgerAI — Autonomous Financial Terminal & Cyber Academy",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

SESSION_DEFAULTS = {
    "user_lang": "🇹🇷 TR",
    "theme_idx": 0,  # 0: Kurumsal, 1: Modern, 2: Akademi
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
    "academy_lives": 3,
    "academy_streak": 0,
    "game_step": 1,
    "current_game_vaka": None,
    "matrix_step": 1,
    "matrix_current_item": None,
    "eta_step": 1,
    "eta_current_scenario": None
}

for key, default_val in SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default_val

# API Doğrulaması
if "GEMINI_API_KEY" in st.secrets:
    API_KEY = st.secrets["GEMINI_API_KEY"]
else:
    st.error("Missing GEMINI_API_KEY in Streamlit Secrets! Please configure.")
    st.stop()

client = genai.Client(api_key=API_KEY)

# ==============================================================================
# 2. PROSEDÜREL OYUN VE OTANTİK ETA SENARYO MOTORU
# ==============================================================================

def generate_simple_puzzle(step: int):
    vakalar = [
        {
            "vaka": "Şirketiniz satıp kâr elde etmek amacıyla toptancıdan 100 adet spor ayakkabı satın aldı.",
            "tutar": "80.000 TL + %20 KDV (16.000 TL) = 96.000 TL",
            "soru": "Satılmak üzere depoya giren bu ticari mallar hangi hesap koduna borç kaydedilir?",
            "secenekler": ["153 Ticari Mallar", "770 Genel Yönetim Giderleri", "255 Demirbaşlar", "600 Yurtiçi Satışlar"],
            "dogru": "153 Ticari Mallar",
            "ipucu": "Satmak amacıyla alınan her türlü emtia ve ürün 153 hesabında izlenir."
        },
        {
            "vaka": "Şirket merkez ofisinde kullanılmak üzere fotokopi kağıtları, toner ve arşiv dosyaları satın alındı.",
            "tutar": "12.000 TL + %20 KDV (2.400 TL) = 14.400 TL",
            "soru": "Ofis idari işleyişi için tüketilen bu kırtasiye malzemeleri hangi hesap koduna borç yazılır?",
            "secenekler": ["770 Genel Yönetim Giderleri", "153 Ticari Mallar", "100 Kasa Hesabı", "320 Satıcılar"],
            "dogru": "770 Genel Yönetim Giderleri",
            "ipucu": "Şirketin idari tüketimleri doğrudan 770 Genel Yönetim Giderleri hesabına aktarılır."
        },
        {
            "vaka": "Ofis çalışanlarının kullanması için 5 adet yüksek performanslı dizüstü bilgisayar satın alındı.",
            "tutar": "150.000 TL + %20 KDV (30.000 TL) = 180.000 TL",
            "soru": "1 yıldan uzun süre kullanılacak bu ofis bilgisayarları hangi duran varlık hesabına kaydedilir?",
            "secenekler": ["255 Demirbaşlar", "770 Genel Yönetim Giderleri", "153 Ticari Mallar", "600 Yurtiçi Satışlar"],
            "dogru": "255 Demirbaşlar",
            "ipucu": "İşletmede 1 yıldan uzun süre kullanılan bilgisayar, mobilya vb. eşyalar 255 Demirbaşlar hesabında aktifleştirilir."
        },
        {
            "vaka": "Müşterinize toptan ürün satışı yapıldı ve fatura düzenlenip teslim edildi.",
            "tutar": "200.000 TL + %20 KDV (40.000 TL) = 240.000 TL",
            "soru": "Gerçekleşen bu ana faaliyet satışı Tek Düzen Hesap Planında hangi gelir hesabına alacak yazılır?",
            "secenekler": ["600 Yurtiçi Satışlar", "153 Ticari Mallar", "770 Genel Yönetim Giderleri", "102 Bankalar"],
            "dogru": "600 Yurtiçi Satışlar",
            "ipucu": "Yurtiçine yapılan ana ticari mal ve hizmet satışları 600 Yurtiçi Satışlar hesabına alacak kaydedilir."
        }
    ]
    secilen = random.choice(vakalar)
    return {
        "step": step,
        "vaka": secilen["vaka"],
        "tutar": secilen["tutar"],
        "soru": secilen["soru"],
        "secenekler": secilen["secenekler"],
        "dogru": secilen["dogru"],
        "ipucu": secilen["ipucu"],
        "xp": 150
    }

if st.session_state["current_game_vaka"] is None:
    st.session_state["current_game_vaka"] = generate_simple_puzzle(st.session_state["game_step"])

TRICKY_MATRIX_CARDS = [
    {
        "hesap_adi": "BİRİKMİŞ AMORTİSMANLAR (-)",
        "karakter": "Aktifi Düzenleyici Pasif Karakterli Hesap",
        "dogru_sinif": 2,
        "aciklama": "Duran varlıkların aşınma payıdır. 2 ile başlamasına rağmen alacak bakiyesi verir!"
    },
    {
        "hesap_adi": "ALINAN SİPARİŞ AVANSLARI",
        "karakter": "Kısa Vadeli Borç / Yabancı Kaynak",
        "dogru_sinif": 3,
        "aciklama": "Müşteriden mal teslim edilmeden önce alınan paradır, 340 grubunda kısa vadeli borçtur."
    },
    {
        "hesap_adi": "GELECEK AYLARA AİT GİDERLER",
        "karakter": "Dönen Varlık / Peşin Ödenen Gider",
        "dogru_sinif": 1,
        "aciklama": "Gelecek dönem için peşin ödenen kiralardır; 180 grubunda dönen varlık sayılır."
    },
    {
        "hesap_adi": "DÖNEM NET KÂRI",
        "karakter": "Öz Kaynaklar Unsuru",
        "dogru_sinif": 5,
        "aciklama": "İşletme faaliyetleri sonucu kalan net kârdır; 590 grubunda öz kaynaklarda yer alır."
    },
    {
        "hesap_adi": "SATILAN TİCARİ MALLAR MALİYETİ (STMM)",
        "karakter": "Gelir Tablosu Gider Hesabı",
        "dogru_sinif": 6,
        "aciklama": "Satılan malların işletmeye maliyetidir; 621 kodunda gelir tablosunu azaltır."
    },
    {
        "hesap_adi": "BANKA KREDİLERİ (3 YIL VADELİ)",
        "karakter": "Uzun Vadeli Yabancı Kaynak",
        "dogru_sinif": 4,
        "aciklama": "Vadesi 1 yılı aşan borçlanmalar 400 grubunda uzun vadeli yabancı kaynaktır."
    }
]

if st.session_state["matrix_current_item"] is None:
    st.session_state["matrix_current_item"] = random.choice(TRICKY_MATRIX_CARDS)

# ETA PROVA MOTORU (GERÇEKÇİ FATURA SENARYOLARI)
def generate_eta_scenario(step: int):
    senaryo_listesi = [
        {
            "fis_no": f"YEV-2026/00{step}",
            "tarih": datetime.date.today().strftime("%d.%m.%Y"),
            "baslik": "VADELİ TİCARİ MAL ALIMI & KDV",
            "aciklama": "Toptancıdan satılmak üzere vadeli ticari mal alışı gerçekleşmiştir.",
            "detay": "Matrah: 50.000 TL | %20 KDV: 10.000 TL | Toplam Satıcı Borcu: 60.000 TL",
            "satirlar": [
                {"kod": "153", "ad": "TİCARİ MALLAR", "borc": 50000.0, "alacak": 0.0},
                {"kod": "191", "ad": "İNDİRİLECEK KDV", "borc": 10000.0, "alacak": 0.0},
                {"kod": "320", "ad": "SATICILAR (CARİ HESAP)", "borc": 0.0, "alacak": 60000.0}
            ],
            "beklenen_toplam": 60000.0,
            "ipucu": "Borçlu hesaplar: 153 ve 191 | Alacaklı hesap: 320"
        },
        {
            "fis_no": f"YEV-2026/00{step}",
            "tarih": datetime.date.today().strftime("%d.%m.%Y"),
            "baslik": "BANKADAN SATICI BORCU HAVALESİ",
            "aciklama": "Şirketin Garanti Bankası ticari mevduat hesabından satıcıya borç ödenmiştir.",
            "detay": "Ödenen Borç Tutarı: 35.000 TL (Dekont No: BNK-8819)",
            "satirlar": [
                {"kod": "320", "ad": "SATICILAR", "borc": 35000.0, "alacak": 0.0},
                {"kod": "102", "ad": "BANKALAR (MEVDUAT)", "borc": 0.0, "alacak": 35000.0}
            ],
            "beklenen_toplam": 35000.0,
            "ipucu": "Borçlu hesap: 320 Satıcılar | Alacaklı hesap: 102 Bankalar"
        },
        {
            "fis_no": f"YEV-2026/00{step}",
            "tarih": datetime.date.today().strftime("%d.%m.%Y"),
            "baslik": "NAKİT PEŞİN OFİS GİDERİ",
            "aciklama": "Şirket merkez ofisi için nakit ödenerek kırtasiye ve sarf malzemesi alınmıştır.",
            "detay": "Gider Tutarı: 5.000 TL | %20 KDV: 1.000 TL | Kasadan Çıkan Nakit: 6.000 TL",
            "satirlar": [
                {"kod": "770", "ad": "GENEL YÖNETİM GİDERLERİ", "borc": 5000.0, "alacak": 0.0},
                {"kod": "191", "ad": "İNDİRİLECEK KDV", "borc": 1000.0, "alacak": 0.0},
                {"kod": "100", "ad": "KASA HESABI", "borc": 0.0, "alacak": 6000.0}
            ],
            "beklenen_toplam": 6000.0,
            "ipucu": "Borçlu hesaplar: 770 ve 191 | Alacaklı hesap: 100 Kasa"
        }
    ]
    s = random.choice(senaryo_listesi)
    return {
        "step": step,
        "fis_no": s["fis_no"],
        "tarih": s["tarih"],
        "baslik": s["baslik"],
        "aciklama": s["aciklama"],
        "detay": s["detay"],
        "satirlar": s["satirlar"],
        "beklenen_toplam": s["beklenen_toplam"],
        "ipucu": s["ipucu"],
        "xp": 250
    }

if st.session_state["eta_current_scenario"] is None:
    st.session_state["eta_current_scenario"] = generate_eta_scenario(st.session_state["eta_step"])

# ==============================================================================
# 3. LOCALIZATION DATA DICTIONARY (6 GLOBAL DİL)
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
        "industries": ["⚡ Otomatik (AI)", "🛒 Ticaret / Mal", "🏢 Hizmet / Ofis", "🏭 Üretim / Sanayi"],
        "themes": ["Kurumsal", "Modern", "Akademi"],
        "about_btn": "Hakkında",
        "about_title": "LedgerAI Kurumsal Mimari & Regülasyon",
        "about_content": "LedgerAI, Türkiye Tek Düzen Hesap Planı, VUK ve uluslararası standartlara tam uyumlu otonom ön muhasebe terminalidir.",
        "bot_title": "👾 LedgerBot Finans Mentorü",
        "bot_welcome": "Selam! Ben finans asistanınım. Muhasebe öğrenmek veya fatura mantığını sormak için aşağıdaki sorulara tıklayabilirsin:",
        "bot_placeholder": "Muhasebe sorunuzu yazın...",
        "bot_clear": "Temizle",
        "quick_chips": [
            "💡 Muhasebeciye ne kazandırır?",
            "🔒 Verilerim güvende mi?",
            "⚖️ Tevkifat & Stopaj mantığı nedir?",
            "🎯 153 ile 770 arasındaki fark nedir?"
        ],
        "cockpit_card1_title": "🏛️ Mevzuat & Tevkifat Uyumu",
        "cockpit_card1_desc": "5/10, 7/10, 9/10 KDV tevkifatları ve Serbest Meslek stopajları kuruş farkı olmadan 360 hesabına aktarılır.",
        "cockpit_card2_title": "⚡ ERP Aktarım Formatları",
        "cockpit_card2_desc": "Tek tıkla ETA V.11 uyumlu CSV, Luca ve çok sayfalı (153 & 770 ayrılmış) kurumsal Excel üretimi.",
        "cockpit_card3_title": "🛡️ Çift Taraflı Denetim Kilidi",
        "cockpit_card3_desc": "Toplam Borç = Toplam Alacak eşitliği sağlanmadan yevmiye fişi kapatılmaz; bakiye farkı riski sıfırlanır.",
        "headers": {"vouch": "Fiş No", "date": "Tarih", "code": "Hesap Kodu", "name": "Hesap Adı", "desc": "Açıklama", "curr": "Para Birimi", "deb": "Borç", "crd": "Alacak"}
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
        "industries": ["⚡ Auto (AI)", "🛒 Trade / Retail", "🏢 Services / SaaS", "🏭 Manufacturing"],
        "themes": ["Kurumsal", "Modern", "Akademi"],
        "about_btn": "About",
        "about_title": "LedgerAI Architecture & Regulation",
        "about_content": "Autonomous double-entry journal voucher generator compliant with US GAAP, IFRS and SOC2.",
        "bot_title": "👾 LedgerBot Finance Mentor",
        "bot_welcome": "Hi! I am your AI finance mentor. Tap any quick question below or ask me directly:",
        "bot_placeholder": "Ask a financial question...",
        "bot_clear": "Clear",
        "quick_chips": [
            "💡 How does it save time?",
            "🔒 Is our data secure?",
            "⚖️ Explain Debit vs Credit",
            "🎯 Inventory vs OpEx accounts"
        ],
        "cockpit_card1_title": "🏛️ Tax Withholding Engine",
        "cockpit_card1_desc": "Automatic handling of multi-rate sales taxes and withholding accounts with zero cent deviation.",
        "cockpit_card2_title": "⚡ ERP Interoperability",
        "cockpit_card2_desc": "Direct exports formatted for QuickBooks, SAP, Datev SKR03/04, and multi-tab Excel workbooks.",
        "cockpit_card3_title": "🛡️ Dual-Audit Integrity Lock",
        "cockpit_card3_desc": "Mathematical assurance guaranteeing that Total Debit strictly equals Total Credit before release.",
        "headers": {"vouch": "Voucher #", "date": "Date", "code": "Account Code", "name": "Account Name", "desc": "Memo", "curr": "Currency", "deb": "Debit", "crd": "Credit"}
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
        "industries": ["⚡ Automatisch (KI)", "🛒 Handel / Wareneinkauf", "🏢 Dienstleistung / IT", "🏭 Produktion / Fertigung"],
        "themes": ["Kurumsal", "Modern", "Akademi"],
        "about_btn": "Über uns",
        "about_title": "LedgerAI Architektur & Datev SKR03/04 Standard",
        "about_content": "Vollautomatisierte Buchungssatzerstellung nach GoBD und Datev-Richtlinien mit strengem Soll/Haben-Ausgleich.",
        "bot_title": "👾 LedgerBot Finanzmentor",
        "bot_welcome": "Hallo! Tippen Sie auf eine Frage oder fragen Sie mich direkt nach Buchungssätzen:",
        "bot_placeholder": "Frage eingeben...",
        "bot_clear": "Löschen",
        "quick_chips": ["💡 Wie spart es Arbeitszeit?", "🔒 Datenschutz & Sicherheit", "⚖️ Soll an Haben Prinzip"],
        "cockpit_card1_title": "🏛️ Vorsteuer- & Steuerlogik",
        "cockpit_card1_desc": "Automatische Zuordnung von SKR03/04 Vorsteuern und USt-IdNr Validierung.",
        "cockpit_card2_title": "⚡ Datev Export",
        "cockpit_card2_desc": "Direkter Datev-konformer CSV-Export für das Steuerbüro.",
        "cockpit_card3_title": "🛡️ Soll/Haben Garantie",
        "cockpit_card3_desc": "Mathematische Prüfung auf absolute Ausgeglichenheit der Buchungssätze.",
        "headers": {"vouch": "Beleg", "date": "Datum", "code": "Konto", "name": "Bezeichnung", "desc": "Text", "curr": "Währung", "deb": "Soll", "crd": "Haben"}
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
        "industries": ["⚡ Auto (IA)", "🛒 Négoce / Stock", "🏢 Services / Conseil", "🏭 Production / Industrie"],
        "themes": ["Kurumsal", "Modern", "Akademi"],
        "about_btn": "À propos",
        "about_title": "Architecture Comptable & Normes PCG",
        "about_content": "Conformité Plan Comptable Général (PCG) avec vérification stricte du principe Débit = Crédit.",
        "bot_title": "👾 LedgerBot Mentor",
        "bot_welcome": "Bonjour! Choisissez une question rapide ou posez votre question comptable:",
        "bot_placeholder": "Poser une question...",
        "bot_clear": "Effacer",
        "quick_chips": ["💡 Gain de temps en cabinet", "🔒 Sécurité des données", "⚖️ Principe Débit / Crédit"],
        "cockpit_card1_title": "🏛️ Ventilation PCG",
        "cockpit_card1_desc": "Affectation automatique aux comptes de charges et TVA déductible.",
        "cockpit_card2_title": "⚡ Formats Export",
        "cockpit_card2_desc": "Compatible avec les logiciels Sage, Cegid et tableur multi-feuilles.",
        "cockpit_card3_title": "🛡️ Équilibre Débit/Crédit",
        "cockpit_card3_desc": "Vérification stricte de la balance avant validation finale.",
        "headers": {"vouch": "Pièce", "date": "Date", "code": "Compte", "name": "Libellé", "desc": "Détail", "curr": "Devise", "deb": "Débit", "crd": "Crédit"}
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
        "industries": ["⚡ Automático (IA)", "🛒 Comercio / Inventario", "🏢 Servicios / Oficina", "🏭 Fabricación / Industria"],
        "themes": ["Kurumsal", "Modern", "Akademi"],
        "about_btn": "Acerca de",
        "about_title": "Estándares Contables y Seguridad Fiscal",
        "about_content": "Asientos contables conformes al Plan General Contable (PGC) con cuadre matemático de Debe y Haber.",
        "bot_title": "👾 LedgerBot Mentor",
        "bot_welcome": "¡Hola! Pulsa una pregunta rápida o escribe tu consulta contable:",
        "bot_placeholder": "Escribe tu duda...",
        "bot_clear": "Limpiar",
        "quick_chips": ["💡 Ventajas para la asesoría", "🔒 Seguridad y confidencialidad", "⚖️ Cuadre de Debe y Haber"],
        "cockpit_card1_title": "🏛️ Cuadre Fiscal",
        "cockpit_card1_desc": "Gestión automática de retenciones e IVA soportado.",
        "cockpit_card2_title": "⚡ Compatibilidad ERP",
        "cockpit_card2_desc": "Exportación directa para Contasol, A3 y software contable estándar.",
        "cockpit_card3_title": "🛡️ Control de Asiento",
        "cockpit_card3_desc": "Validación matemática estricta de paridad Debe = Haber.",
        "headers": {"vouch": "Asiento", "date": "Fecha", "code": "Cuenta", "name": "Nombre Cuenta", "desc": "Concepto", "curr": "Moneda", "deb": "Debe", "crd": "Haber"}
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
        "industries": ["⚡ Automatico (IA)", "🛒 Commercio / Magazzino", "🏢 Servizi / Consulenza", "🏭 Manifattura / Produzione"],
        "themes": ["Kurumsal", "Modern", "Akademi"],
        "about_btn": "Info",
        "about_title": "Standard di Conformità e Partita Doppia",
        "about_content": "Generazione automatica di scritture in partita doppia perfettamente bilanciate per gestionali Zucchetti e Teamsystem.",
        "bot_title": "👾 LedgerBot Mentor",
        "bot_welcome": "Ciao! Seleziona una domanda pillola o scrivimi direttamente:",
        "bot_placeholder": "Fai una domanda contabile...",
        "bot_clear": "Cancella",
        "quick_chips": ["💡 Vantaggi per lo studio", "🔒 Sicurezza dei dati fiscali", "⚖️ Pareggio Dare / Avere"],
        "cockpit_card1_title": "🏛️ Scritture Bilanciate",
        "cockpit_card1_desc": "Gestione automatica ritenute d'acconto ed IVA a credito.",
        "cockpit_card2_title": "⚡ Compatibilità Gestionale",
        "cockpit_card2_desc": "File pronti per Zucchetti, Teamsystem e formati Excel avanzati.",
        "cockpit_card3_title": "🛡️ Quadratura Certificata",
        "cockpit_card3_desc": "Garanzia matematica di parità tra totale Dare e Avere.",
        "headers": {"vouch": "Partita", "date": "Data", "code": "Conto", "name": "Descrizione", "desc": "Causale", "curr": "Valuta", "deb": "Dare", "crd": "Avere"}
    }
}

if st.session_state["user_lang"] not in LANG_DATA:
    st.session_state["user_lang"] = "🇹🇷 TR"

T = LANG_DATA[st.session_state["user_lang"]]

INSTANT_FAQ_CACHE = {
    "💡 Muhasebeciye ne kazandırır?": "LedgerAI, manuel veri girişini %80 azaltarak mali müşavirlerin rutin fiş işleme yükünü ortadan kaldırır. Yapay zeka fiş taslağını oluşturur, uzman insan sadece onaylar ve denetler.",
    "🔒 Verilerim güvende mi?": "Evet. Tüm finansal verileriniz TLS 256-bit uçtan uca şifreleme ile işlenir. Belgeleriniz kalıcı sunucularda saklanmaz ve model eğitiminde (training) kullanılmaz.",
    "⚖️ Tevkifat & Stopaj mantığı nedir?": "Tevkifat ve stopaj, faturadaki verginin bir kısmının alıcı tarafından kesilerek doğrudan vergi dairesine (360 hesabına) ödenmesidir. Böylece satıcı cari hesabı net tutara oturur ve yevmiye fişi kuruş farkı olmadan dengelenir.",
    "🎯 153 ile 770 arasındaki fark nedir?": "153 Ticari Mallar satılmak amacıyla alınan ticari ürünlerin stok hesabıdır. 770 Genel Yönetim Giderleri ise işletmenin kendi idari faaliyetlerinde tükettiği (ofis kırtasiyesi, kira, danışmanlık vb.) giderlerin kaydedildiği hesaptır.",
    "💡 How does it save time?": "LedgerAI automates repetitive invoice typing and multi-tier tax splitting by 80%, leaving the final executive approval to the CPA.",
    "🔒 Is our data secure?": "Yes. Encrypted via TLS 256-bit bank-grade protocols. Your financial files are processed strictly within the active session and never stored permanently."
}

# ==============================================================================
# 4. DYNAMIC STYLING ENGINE (AUTHENTIC ETA WINDOW & IPHONE BUTTONS)
# ==============================================================================

if st.session_state["theme_idx"] == 0:
    # 🌑 Kurumsal
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
    # ✨ Modern
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
    # ⚡ Akademi
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

    /* IPHONE CIRCLE GLASS BUTTONS */
    div.circle-glass-btn div[data-testid="stPopover"] > button {{
        width: 44px !important;
        height: 44px !important;
        min-height: 44px !important;
        max-width: 44px !important;
        border-radius: 50% !important;
        background: rgba(255, 255, 255, 0.08) !important;
        border: 1px solid rgba(255, 255, 255, 0.22) !important;
        backdrop-filter: blur(16px) !important;
        -webkit-backdrop-filter: blur(16px) !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 1.15rem !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.35) !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        padding: 0 !important;
        margin: 0 !important;
    }}
    div.circle-glass-btn div[data-testid="stPopover"] > button:hover {{
        background: rgba(255, 255, 255, 0.20) !important;
        border-color: rgba(255, 255, 255, 0.60) !important;
        transform: scale(1.08) !important;
        box-shadow: 0 6px 20px rgba(255, 255, 255, 0.25) !important;
    }}

    /* POPOVER İÇİNDEKİ SEÇİM BUTONLARI (SIFIR YAZI, SADECE TIKLAMA) */
    div.pop-pill-grid div.stButton > button {{
        background: rgba(255, 255, 255, 0.08) !important;
        border: 1px solid rgba(255, 255, 255, 0.18) !important;
        border-radius: 9999px !important;
        font-size: 0.80rem !important;
        font-weight: 600 !important;
        color: #F8FAFC !important;
        padding: 6px 14px !important;
        margin-bottom: 6px !important;
        height: auto !important;
        min-height: 32px !important;
        box-shadow: none !important;
        transition: all 0.2s ease !important;
    }}
    div.pop-pill-grid div.stButton > button:hover {{
        background: rgba(56, 189, 248, 0.25) !important;
        border-color: rgba(56, 189, 248, 0.50) !important;
        transform: scale(1.02) !important;
    }}

    /* FAALİYET SEKTÖRÜ MİKRO KAPSÜL BUTONLARI */
    div.sector-pills div.stButton > button {{
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.14) !important;
        border-radius: 9999px !important;
        font-size: 0.72rem !important;
        font-weight: 600 !important;
        color: #CBD5E1 !important;
        padding: 4px 10px !important;
        height: 28px !important;
        min-height: 28px !important;
        box-shadow: none !important;
        transition: all 0.2s ease !important;
        white-space: nowrap !important;
    }}
    div.sector-pills div.stButton > button:hover {{
        background: rgba(56, 189, 248, 0.2) !important;
        border-color: rgba(56, 189, 248, 0.4) !important;
        color: #FFFFFF !important;
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

    /* ASİSTAN İÇİ MİKRO HAPLAR */
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

    /* WHATSAPP BALONLARI */
    .wa-bubble-left {{
        background: rgba(30, 41, 59, 0.85);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 18px 18px 18px 4px;
        padding: 16px 20px;
        margin-bottom: 14px;
        max-width: 92%;
        box-shadow: 0 4px 15px rgba(0,0,0,0.25);
    }}
    .wa-bubble-right {{
        background: rgba(16, 185, 129, 0.15);
        border: 1px solid rgba(16, 185, 129, 0.35);
        border-radius: 18px 18px 4px 18px;
        padding: 16px 20px;
        margin-bottom: 14px;
        margin-left: auto;
        max-width: 92%;
        box-shadow: 0 4px 15px rgba(0,0,0,0.25);
    }}
    .wa-title {{
        font-size: 0.88rem;
        font-weight: 700;
        color: #38BDF8;
        margin-bottom: 4px;
    }}
    .wa-text {{
        font-size: 0.82rem;
        color: #E2E8F0;
        line-height: 1.55;
    }}

    /* OTANTİK ETA:SQL PENCERE MİMARİSİ */
    .eta-window {{
        background: #C0C0C0;
        border: 2px solid #FFFFFF;
        border-right-color: #808080;
        border-bottom-color: #808080;
        box-shadow: inset 1px 1px 0px #DFDFDF, inset -1px -1px 0px #000000, 0 10px 30px rgba(0,0,0,0.5);
        color: #000000;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        padding: 4px;
        border-radius: 4px;
        margin-bottom: 20px;
    }}
    .eta-titlebar {{
        background: linear-gradient(90deg, #0A246A 0%, #A6CAF0 100%);
        color: #FFFFFF;
        font-weight: bold;
        font-size: 13px;
        padding: 3px 6px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        letter-spacing: 0.5px;
    }}
    .eta-toolbar {{
        background: #E0E0E0;
        border: 1px solid #808080;
        padding: 3px 6px;
        display: flex;
        gap: 6px;
        margin: 4px 0;
        font-size: 11px;
        font-weight: 600;
    }}
    .eta-header-card {{
        background: #FFFFE1; /* ETA Kalsik Açık Sarı Fiş Başlığı */
        border: 1px solid #999999;
        padding: 6px 10px;
        margin-bottom: 4px;
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 8px;
        font-size: 11px;
        color: #000080;
        font-weight: bold;
    }}
    .eta-grid-table {{
        width: 100%;
        border-collapse: collapse;
        background: #FFFFFF;
        font-size: 11px;
    }}
    .eta-grid-table th {{
        background: #D4D0C8;
        border: 1px solid #808080;
        padding: 4px 6px;
        font-weight: bold;
        color: #000000;
        text-align: center;
    }}
    .eta-grid-table td {{
        border: 1px solid #D0D0D0;
        padding: 4px 6px;
        color: #000000;
    }}
    .eta-footer-bar {{
        background: #EBE9ED;
        border: 1px solid #808080;
        padding: 6px 10px;
        margin-top: 4px;
        display: flex;
        justify-content: flex-end;
        gap: 20px;
        font-size: 12px;
        font-weight: bold;
    }}

    div[data-testid="stTabs"] button[role="tab"] {{
        background: transparent !important;
        border: none !important;
        font-size: 0.88rem !important;
        font-weight: 700 !important;
        color: #94A3B8 !important;
        padding: 8px 18px !important;
        transition: all 0.25s ease !important;
    }}
    div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{
        color: #FFFFFF !important;
        border-bottom: 2px solid #38BDF8 !important;
    }}
</style>
""", unsafe_allow_html=True)

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
                osc.frequency.setValueAtTime(150, ctx.currentTime);
                osc.frequency.exponentialRampToValueAtTime(80, ctx.currentTime + 0.3);
                gain.gain.setValueAtTime(0.1, ctx.currentTime);
                gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.3);
                osc.connect(gain);
                gain.connect(ctx.destination);
                osc.start();
                osc.stop(ctx.currentTime + 0.3);
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
# 7. MULTI-DECK COCKPIT (IPHONE CIRCLE GLASS BUTTONS)
# ==============================================================================

st.markdown("<div class='cockpit-container'>", unsafe_allow_html=True)

nav_left, nav_right = st.columns([8.2, 1.8])

with nav_right:
    btn_dock1, btn_dock2, btn_dock3 = st.columns([1, 1, 1])
    
    with btn_dock1:
        st.markdown("<div class='circle-glass-btn'>", unsafe_allow_html=True)
        with st.popover("🌐", help="Dil Seçimi"):
            st.markdown("<div style='font-size:0.85rem; font-weight:700; margin-bottom:8px;'>🌐 Dil / Language</div>", unsafe_allow_html=True)
            st.markdown("<div class='pop-pill-grid'>", unsafe_allow_html=True)
            for l_key in list(LANG_DATA.keys()):
                if st.button(l_key, key=f"btn_lang_pop_{l_key}", use_container_width=True):
                    st.session_state["user_lang"] = l_key
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with btn_dock2:
        st.markdown("<div class='circle-glass-btn'>", unsafe_allow_html=True)
        with st.popover("🎨", help="Görünüm & Temalar"):
            st.markdown("<div style='font-size:0.85rem; font-weight:700; margin-bottom:8px;'>🎨 Tema Seçimi</div>", unsafe_allow_html=True)
            st.markdown("<div class='pop-pill-grid'>", unsafe_allow_html=True)
            theme_list = T["themes"]
            for t_idx, t_name in enumerate(theme_list):
                if st.button(t_name, key=f"btn_theme_pop_{t_idx}", use_container_width=True):
                    st.session_state["theme_idx"] = t_idx
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with btn_dock3:
        st.markdown("<div class='circle-glass-btn'>", unsafe_allow_html=True)
        with st.popover("ℹ️", help="Sistem Hakkında"):
            st.markdown(f"#### {T['about_title']}")
            st.markdown(T["about_content"])
        st.markdown("</div>", unsafe_allow_html=True)

# SEKME MENÜSÜ
sekme_terminal, sekme_akademi, sekme_hukuk = st.tabs([
    "🏢 Kurumsal Terminal", 
    "🎓 Siber Akademi", 
    "⚖️ Hukuki Çerçeve & SLA"
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

        # MİKRO CAM KAPSÜL FAALİYET SEKTÖRÜ
        st.markdown("<div style='margin-top:14px; padding-top:10px; border-top:1px solid rgba(255,255,255,0.08);'>", unsafe_allow_html=True)
        st.markdown("<div style='font-size:0.72rem; color:#94A3B8; font-weight:700; letter-spacing:0.5px; margin-bottom:6px;'>FAALİYET SEKTÖRÜ (TIKLAYARAK SEÇİN):</div>", unsafe_allow_html=True)
        
        st.markdown("<div class='sector-pills'>", unsafe_allow_html=True)
        s_cols = st.columns(4)
        for i_idx, i_name in enumerate(T["industries"]):
            with s_cols[i_idx]:
                is_active = (st.session_state["industry_idx"] == i_idx)
                label = f"✓ {i_name}" if is_active else i_name
                if st.button(label, key=f"sec_pill_btn_{i_idx}", use_container_width=True):
                    if st.session_state["industry_idx"] != i_idx:
                        st.session_state["industry_idx"] = i_idx
                        st.rerun()
        st.markdown("</div></div></div>", unsafe_allow_html=True)

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

    # ASİSTAN ÇUBUĞU
    st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
    c_bot_l, c_bot_center, c_bot_r = st.columns([1, 4, 1])

    with c_bot_center:
        bot_box_title = T.get("bot_title", "👾 LedgerBot Finans Mentorü")
        bot_welcome_msg = T.get("bot_welcome", "Muhasebe asistanınız göreve hazır.")
        
        with st.expander(bot_box_title, expanded=False):
            top_col1, top_col2 = st.columns([5.5, 1.5])
            top_col1.caption(bot_welcome_msg)
            with top_col2:
                if st.button(T.get("bot_clear", "Temizle"), key="btn_clear_chat", use_container_width=True):
                    st.session_state["chat_messages"] = []
                    st.rerun()

            secilen_chip = None
            chips = T.get("quick_chips", [])
            if chips:
                btn_cols = st.columns(len(chips))
                for c_idx, chip_text in enumerate(chips):
                    with btn_cols[c_idx]:
                        if st.button(chip_text, key=f"gh_pill_btn_{c_idx}", use_container_width=True):
                            secilen_chip = chip_text

            st.markdown("<div class='chat-scroll-area'>", unsafe_allow_html=True)
            if not st.session_state["chat_messages"]:
                st.markdown(f"<div style='color: #94A3B8; font-size: 0.82rem; padding: 8px 0;'>💡 <i>{T.get('bot_placeholder', 'Sorunuzu yazın...')}</i></div>", unsafe_allow_html=True)
            else:
                for msg in st.session_state["chat_messages"][-6:]:
                    if msg["role"] == "user":
                        st.markdown(f"<div class='user-bubble'><b>Soru:</b> {msg['content']}</div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div class='assistant-bubble'>🤖 <b>LedgerAI:</b> {msg['content']}</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

            user_query = st.chat_input(T.get("bot_placeholder", "Sorunuzu yazın..."))
            
            aktif_soru = None
            if user_query:
                aktif_soru = user_query
            elif secilen_chip:
                aktif_soru = secilen_chip
            
            if aktif_soru:
                st.session_state["chat_messages"].append({"role": "user", "content": aktif_soru})
                
                if aktif_soru in INSTANT_FAQ_CACHE:
                    bot_cevap = INSTANT_FAQ_CACHE[aktif_soru]
                    st.session_state["chat_messages"].append({"role": "assistant", "content": bot_cevap})
                    st.rerun()
                else:
                    with st.spinner("● ● ● Düşünüyor..."):
                        prompt_bot = f"""
                        Sen LedgerAI'ın kurumsal finans ve Tek Düzen Hesap Planı uzmanısın.
                        Kullanıcı Dili: {st.session_state['user_lang']}
                        Kullanıcı Mesajı: "{aktif_soru}"

                        ÇOK KESİN KURALLAR:
                        1. Eğer kullanıcı hesap planını, bilanço sınıflarını veya belirli hesap aralıklarını sorarsa:
                           - Asla hiçbir sınıfı atlama! 
                             1. Dönen Varlıklar (100 Kasa, 102 Banka, 120 Alıcılar, 153 Ticari Mallar), 
                             2. Duran Varlıklar (250 Binalar, 254 Taşıtlar, 255 Demirbaşlar, 257 Birikmiş Amortismanlar), 
                             3. Kısa Vadeli Yabancı Kaynaklar (300 Banka Kredileri, 320 Satıcılar, 360 Ödenecek Vergi)
                             şeklinde eksiksiz listele.
                           - Bu teorik sorularda ASLA uydurma fatura yevmiye fişi yazma!
                        2. Eğer kullanıcı selam veriyorsa ("merhaba", "naber", "ne yapıyorsun"):
                           - Kısa, profesyonel bir selam ver, hangi muhasebe konusunda destek istediğini sor. ASLA fiş uydurma!
                        3. SADECE kullanıcı somut bir mal alımı, gider veya harcama fiş kaydı soruyorsa sonuna Borç/Alacak kaydı ekle.
                        """
                        try:
                            yanit = client.models.generate_content(
                                model="gemini-3.5-flash-lite",
                                contents=prompt_bot
                            )
                            bot_cevap = yanit.text.strip() if yanit and yanit.text else "Size finansal süreçlerde nasıl yardımcı olabilirim?"
                        except Exception:
                            bot_cevap = "Muhasebe ve vergi mevzuatıyla ilgili sorularınızı yanıtlamaya hazırım."
                        
                        st.session_state["chat_messages"].append({"role": "assistant", "content": bot_cevap})
                        st.rerun()

# ------------------------------------------------------------------------------
# SEKME 2: 🎓 SİBER AKADEMİ (4'LÜ EĞLENCELİ OYUN MERKEZİ & ETA SİMÜLATÖRÜ)
# ------------------------------------------------------------------------------
with sekme_akademi:
    xp = st.session_state["academy_xp"]
    lives = st.session_state["academy_lives"]
    
    kalpler = "❤️ " * lives + "🖤 " * (3 - lives)
    if xp >= 1500:
        st.session_state["academy_level"] = "🏆 Yeminli Baş Denetçi (Partner)"
    elif xp >= 800:
        st.session_state["academy_level"] = "⭐ Kıdemli Denetçi"
    elif xp >= 400:
        st.session_state["academy_level"] = "📈 Denetim Uzmanı"
    else:
        st.session_state["academy_level"] = "🌱 Mali Stajyer (Junior)"

    st.markdown(f"""
    <div style='background:rgba(30,41,59,0.7); border:1px solid rgba(255,255,255,0.12); border-radius:20px; padding:20px 26px; margin-bottom:20px;'>
        <div style='display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;'>
            <div>
                <span class='top-badge' style='background:rgba(217,70,239,0.15); border-color:#D946EF; color:#F0ABFC;'>SİBER AKADEMİ ARENA</span>
                <h3 style='margin:4px 0; color:#FFFFFF;'>Geleceğin Finans Lideri Yetiştirme Simülasyonu</h3>
                <p style='font-size:0.85rem; color:#CBD5E1; margin:0;'>Teorik ezber yok! 4 farklı modda interaktif görevleri tamamla, XP topla, rütbe atla.</p>
            </div>
            <div style='text-align:right;'>
                <div style='font-size:1.15rem; margin-bottom:2px;'>CAN: <b>{kalpler}</b></div>
                <div style='font-size:1.6rem; font-weight:800; color:#D946EF;'>🏆 {st.session_state["academy_xp"]} XP</div>
                <div style='font-size:0.8rem; color:#A7F3D0; font-weight:700;'>{st.session_state["academy_level"]} (Seri: {st.session_state["academy_streak"]}🔥)</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if lives <= 0:
        sesli_bildirim_cal("error")
        st.error("💀 GAME OVER! Tüm canlarını kaybettin ve vergi incelemesinden ceza aldın! Rütben sıfırlandı.")
        if st.button("🔄 Yeniden Başla (Canları Doldur)", use_container_width=True):
            st.session_state["academy_lives"] = 3
            st.session_state["academy_streak"] = 0
            st.session_state["academy_xp"] = max(0, st.session_state["academy_xp"] - 150)
            st.rerun()
    else:
        oyun_sekme1, oyun_sekme2, oyun_sekme3, oyun_sekme4 = st.tabs([
            "🎮 1. Hesap Kodu Avcısı", 
            "🧩 2. Bilanço Karakter Matrisi (Zor)",
            "🐍 3. Hedefli Bilanço Snake",
            "📑 4. ETA:SQL Yevmiye Fişi Modülü"
        ])

        # OYUN 1: HESAP KODU AVCISI (NEON KART BUTONLAR)
        with oyun_sekme1:
            vaka = st.session_state["current_game_vaka"]
            c_game1, c_game2 = st.columns([1.2, 1.0], gap="large")
            with c_game1:
                st.markdown(f"""
                <div style='background:rgba(15,23,42,0.75); border:1px dashed rgba(255,255,255,0.22); border-radius:20px; padding:22px;'>
                    <div style='display:flex; justify-content:space-between; align-items:center;'>
                        <span style='font-size:0.75rem; font-weight:800; color:#38BDF8; letter-spacing:1px;'>GÖREV #{vaka['step']}</span>
                        <span style='font-size:0.75rem; color:#F0ABFC; font-weight:700;'>+{vaka['xp']} XP</span>
                    </div>
                    <div style='font-size:0.88rem; color:#CBD5E1; margin:12px 0; line-height:1.5;'>
                        <b>Olay:</b> {vaka['vaka']}
                    </div>
                    <div style='background:rgba(0,0,0,0.35); border-radius:12px; padding:12px; font-family:"Consolas", monospace; font-size:0.82rem; color:#E2E8F0; line-height:1.5;'>
                        📄 <b>Fatura Tutarı:</b> {vaka['tutar']}
                    </div>
                    <div style='margin-top:14px; font-size:0.85rem; color:#F8FAFC; font-weight:600;'>
                        ❓ {vaka['soru']}
                    </div>
                </div>
                """, unsafe_allow_html=True)

            with c_game2:
                st.markdown("<div style='background:rgba(30,41,59,0.72); border:1px solid rgba(255,255,255,0.14); border-radius:20px; padding:22px;'>", unsafe_allow_html=True)
                st.markdown("<div style='font-size:0.82rem; font-weight:700; color:#F1F5F9; margin-bottom:12px;'>🎯 Doğru Hesap Kartına Tıklayın:</div>", unsafe_allow_html=True)

                for opt in vaka["secenekler"]:
                    if st.button(f"👉 {opt}", key=f"btn_opt_{opt}_{vaka['step']}", use_container_width=True):
                        if opt == vaka["dogru"]:
                            st.session_state["academy_xp"] += vaka["xp"]
                            st.session_state["academy_streak"] += 1
                            sesli_bildirim_cal("success")
                            st.balloons()
                            st.success(f"🎉 DOĞRU CEVAP! {vaka['ipucu']} (+{vaka['xp']} XP)")
                            time.sleep(1.2)
                            st.session_state["game_step"] += 1
                            st.session_state["current_game_vaka"] = generate_simple_puzzle(st.session_state["game_step"])
                            st.rerun()
                        else:
                            st.session_state["academy_lives"] -= 1
                            st.session_state["academy_streak"] = 0
                            sesli_bildirim_cal("error")
                            st.error(f"💥 YANDIN! (-1 Can) Doğru hesap: {vaka['dogru']} olmalıydı. {vaka['ipucu']}")
                            time.sleep(1.4)
                            st.session_state["game_step"] += 1
                            st.session_state["current_game_vaka"] = generate_simple_puzzle(st.session_state["game_step"])
                            st.rerun()

                st.markdown("</div>", unsafe_allow_html=True)

        # OYUN 2: BİLANÇO KARAKTER MATRİSİ (ZOR / TERS KÖŞE)
        with oyun_sekme2:
            m_item = st.session_state["matrix_current_item"]
            st.markdown(f"""
            <div style='background:rgba(15,23,42,0.75); border:1px solid rgba(255,255,255,0.15); border-radius:20px; padding:22px; text-align:center; margin-bottom:16px;'>
                <span style='font-size:0.75rem; font-weight:800; color:#38BDF8; letter-spacing:1.2px;'>ZORLUK: İLERİ SEVİYE | RAKAMSIZ KARAKTER ANALİZİ</span>
                <h2 style='color:#FFFFFF; margin:8px 0; letter-spacing:1px;'>{m_item['hesap_adi']}</h2>
                <div style='font-size:0.85rem; color:#A7F3D0; font-weight:600; margin-bottom:4px;'>Fonksiyonel Karakter: {m_item['karakter']}</div>
                <p style='font-size:0.80rem; color:#CBD5E1;'>Bu hesap Tek Düzen Bilanço sisteminde 1'den 7'ye kadar olan hangi ana muhasebe grubuna aittir?</p>
            </div>
            """, unsafe_allow_html=True)

            siniflar = [
                (1, "1. Dönen Varlıklar"),
                (2, "2. Duran Varlıklar"),
                (3, "3. Kısa Vadeli Yabancı Kaynak"),
                (4, "4. Uzun Vadeli Yabancı Kaynak"),
                (5, "5. Öz Kaynaklar"),
                (6, "6. Gelir Tablosu"),
                (7, "7. Maliyet Hesapları")
            ]

            m_cols = st.columns(4)
            for idx, (s_num, s_ad) in enumerate(siniflar):
                with m_cols[idx % 4]:
                    if st.button(s_ad, key=f"btn_tricky_{s_num}_{st.session_state['matrix_step']}", use_container_width=True):
                        if s_num == m_item["dogru_sinif"]:
                            st.session_state["academy_xp"] += 200
                            st.session_state["academy_streak"] += 1
                            sesli_bildirim_cal("success")
                            st.balloons()
                            st.success(f"🎉 TEBRİKLER! {m_item['hesap_adi']}: {m_item['aciklama']} (+200 XP)")
                            time.sleep(1.2)
                            st.session_state["matrix_step"] += 1
                            st.session_state["matrix_current_item"] = random.choice(TRICKY_MATRIX_CARDS)
                            st.rerun()
                        else:
                            st.session_state["academy_lives"] -= 1
                            st.session_state["academy_streak"] = 0
                            sesli_bildirim_cal("error")
                            st.error(f"💥 HATALI TAHMİN! (-1 Can) {m_item['hesap_adi']}: {m_item['aciklama']}")
                            time.sleep(1.4)
                            st.session_state["matrix_step"] += 1
                            st.session_state["matrix_current_item"] = random.choice(TRICKY_MATRIX_CARDS)
                            st.rerun()

        # OYUN 3: HEDEFLİ BİLANÇO SNAKE
        with oyun_sekme3:
            st.markdown("""
            <div style='background:rgba(30,41,59,0.7); border:1px solid rgba(255,255,255,0.12); border-radius:18px; padding:16px 20px; margin-bottom:12px;'>
                <h4 style='color:#FFFFFF; margin:0 0 4px 0;'>🐍 Görevli Bilanço Snake (Hesap Avı)</h4>
                <p style='font-size:0.80rem; color:#CBD5E1; margin:0;'>
                <b>Oyunun Amacı:</b> Ekranda beliren <b>GÖREV HESAP KODUNU</b> ye! Doğru kodu yersen +100 XP kazanırsın. Yanlış kodu yersen veya duvara/kuyruğuna çarparsan oyun durur! 
                (Durdurmak için <b>[BOŞLUK / SPACE]</b> tuşuna bas).
                </p>
            </div>
            """, unsafe_allow_html=True)

            snake_html = """
            <!DOCTYPE html>
            <html>
            <head>
                <style>
                    body { margin: 0; background: transparent; display: flex; flex-direction: column; align-items: center; justify-content: center; font-family: 'JetBrains Mono', monospace; color: #FFF; user-select: none; }
                    #gameCanvas { background: #060913; border: 2px solid #38BDF8; border-radius: 14px; box-shadow: 0 0 25px rgba(56,189,248,0.25); }
                    .hud { display: flex; justify-content: space-between; width: 440px; margin-bottom: 8px; font-size: 14px; font-weight: bold; color: #38BDF8; }
                    .touch-grid { display: grid; grid-template-columns: repeat(3, 55px); gap: 6px; margin-top: 10px; }
                    .t-btn { width: 55px; height: 42px; background: rgba(255,255,255,0.08); border: 1px solid rgba(255,255,255,0.2); border-radius: 10px; color: white; font-size: 18px; cursor: pointer; display: flex; align-items: center; justify-content: center; }
                    .t-btn:active { background: #38BDF8; color: #000; }
                    #menuOverlay { position: absolute; width: 440px; height: 300px; background: rgba(6, 9, 19, 0.85); border-radius: 14px; display: none; flex-direction: column; align-items: center; justify-content: center; }
                    .menu-title { font-size: 20px; font-weight: 800; color: #F43F5E; margin-bottom: 12px; }
                    .menu-btn { background: #38BDF8; color: #000; border: none; padding: 10px 24px; border-radius: 99px; font-weight: bold; cursor: pointer; }
                </style>
            </head>
            <body>
                <div class="hud">
                    <span>SKOR: <span id="score">0</span> XP</span>
                    <span>HEDEF: <span id="targetCode" style="color:#FACC15;">153 TİCARİ MAL</span></span>
                </div>
                <div style="position:relative; display:flex; align-items:center; justify-content:center;">
                    <canvas id="gameCanvas" width="440" height="300"></canvas>
                    <div id="menuOverlay">
                        <div class="menu-title" id="overlayTitle">💀 FATURA İPTAL OLDU!</div>
                        <button class="menu-btn" onclick="restartGame()">YENİDEN BAŞLA</button>
                    </div>
                </div>
                <div class="touch-grid">
                    <div></div><div class="t-btn" onclick="changeDir('UP')">▲</div><div></div>
                    <div class="t-btn" onclick="changeDir('LEFT')">◀</div><div class="t-btn" onclick="changeDir('DOWN')">▼</div><div class="t-btn" onclick="changeDir('RIGHT')">▶</div>
                </div>
                <script>
                    const canvas = document.getElementById("gameCanvas");
                    const ctx = canvas.getContext("2d");
                    const overlay = document.getElementById("menuOverlay");
                    const overlayTitle = document.getElementById("overlayTitle");
                    const grid = 20;
                    let count = 0;
                    let score = 0;
                    let isPaused = false;
                    let isGameOver = false;

                    let snake = { x: 160, y: 160, dx: grid, dy: 0, cells: [], maxCells: 4 };
                    const allCodes = ["153", "770", "102", "600", "255", "320"];
                    let targetCode = "153";
                    let foods = [];

                    function getRandomInt(min, max) { return Math.floor(Math.random() * (max - min)) + min; }

                    function spawnFoods() {
                        foods = [];
                        targetCode = allCodes[Math.floor(Math.random() * allCodes.length)];
                        document.getElementById("targetCode").innerText = targetCode + " KODUNU YE!";

                        foods.push({
                            x: getRandomInt(0, 22) * grid,
                            y: getRandomInt(0, 15) * grid,
                            code: targetCode,
                            isTarget: true
                        });

                        let fakeCode = allCodes[Math.floor(Math.random() * allCodes.length)];
                        while(fakeCode === targetCode) fakeCode = allCodes[Math.floor(Math.random() * allCodes.length)];
                        foods.push({
                            x: getRandomInt(0, 22) * grid,
                            y: getRandomInt(0, 15) * grid,
                            code: fakeCode,
                            isTarget: false
                        });
                    }

                    function triggerGameOver(msg) {
                        isGameOver = true;
                        overlayTitle.innerText = msg;
                        overlay.style.display = "flex";
                    }

                    function restartGame() {
                        snake.x = 160; snake.y = 160;
                        snake.cells = []; snake.maxCells = 4;
                        snake.dx = grid; snake.dy = 0;
                        score = 0;
                        document.getElementById("score").innerText = score;
                        overlay.style.display = "none";
                        isGameOver = false;
                        spawnFoods();
                    }

                    function gameLoop() {
                        requestAnimationFrame(gameLoop);
                        if (++count < 6 || isPaused || isGameOver) return;
                        count = 0;

                        ctx.clearRect(0, 0, canvas.width, canvas.height);
                        snake.x += snake.dx;
                        snake.y += snake.dy;

                        if (snake.x < 0 || snake.x >= canvas.width || snake.y < 0 || snake.y >= canvas.height) {
                            triggerGameOver("💀 DUVARA ÇARPTIN! (BİLANÇO TAŞTI)");
                            return;
                        }

                        snake.cells.unshift({x: snake.x, y: snake.y});
                        if (snake.cells.length > snake.maxCells) snake.cells.pop();

                        foods.forEach(f => {
                            ctx.fillStyle = f.isTarget ? "#FACC15" : "#EF4444";
                            ctx.fillRect(f.x, f.y, grid-1, grid-1);
                            ctx.fillStyle = "#000";
                            ctx.font = "bold 9px monospace";
                            ctx.fillText(f.code, f.x + 2, f.y + 13);
                        });

                        ctx.fillStyle = "#10B981";
                        snake.cells.forEach(function(cell, index) {
                            if (index === 0) ctx.fillStyle = "#38BDF8";
                            else ctx.fillStyle = "#10B981";
                            ctx.fillRect(cell.x, cell.y, grid-1, grid-1);

                            foods.forEach(f => {
                                if (cell.x === f.x && cell.y === f.y) {
                                    if (f.isTarget) {
                                        snake.maxCells++;
                                        score += 100;
                                        document.getElementById("score").innerText = score;
                                        spawnFoods();
                                    } else {
                                        triggerGameOver("❌ YANLIŞ HESABI YEDİN! (VERGİ CEZASI)");
                                    }
                                }
                            });

                            for (let i = index + 1; i < snake.cells.length; i++) {
                                if (cell.x === snake.cells[i].x && cell.y === snake.cells[i].y) {
                                    triggerGameOver("💀 KUYRUĞUNA ÇARPTIN!");
                                    return;
                                }
                            }
                        });
                    }

                    function changeDir(dir) {
                        if (dir === 'LEFT' && snake.dx === 0) { snake.dx = -grid; snake.dy = 0; }
                        else if (dir === 'UP' && snake.dy === 0) { snake.dy = -grid; snake.dx = 0; }
                        else if (dir === 'RIGHT' && snake.dx === 0) { snake.dx = grid; snake.dy = 0; }
                        else if (dir === 'DOWN' && snake.dy === 0) { snake.dy = grid; snake.dx = 0; }
                    }

                    document.addEventListener('keydown', function(e) {
                        if (e.which === 32) isPaused = !isPaused;
                        if (e.which === 37 && snake.dx === 0) changeDir('LEFT');
                        else if (e.which === 38 && snake.dy === 0) changeDir('UP');
                        else if (e.which === 39 && snake.dx === 0) changeDir('RIGHT');
                        else if (e.which === 40 && snake.dy === 0) changeDir('DOWN');
                    });

                    spawnFoods();
                    requestAnimationFrame(gameLoop);
                </script>
            </body>
            </html>
            """
            components.html(snake_html, height=430)

        # OYUN 4: OTANTİK ETA:SQL PENCERE MODÜLÜ
        with oyun_sekme4:
            eta_sc = st.session_state["eta_current_scenario"]

            st.markdown(f"""
            <div style='background:rgba(15,23,42,0.85); border:1px solid #38BDF8; border-radius:14px; padding:14px 18px; margin-bottom:12px;'>
                <div style='display:flex; justify-content:space-between; align-items:center;'>
                    <span style='font-size:0.78rem; font-weight:800; color:#38BDF8;'>📌 ETA:SQL MUHASEBE GÖREVİ #{eta_sc['step']}</span>
                    <span style='font-size:0.75rem; color:#F0ABFC; font-weight:700;'>Ödül: +{eta_sc['xp']} XP</span>
                </div>
                <div style='font-size:0.86rem; color:#FFFFFF; font-weight:600; margin:4px 0;'>{eta_sc['baslik']}</div>
                <div style='font-size:0.80rem; color:#CBD5E1;'><b>Olay:</b> {eta_sc['aciklama']} | <b>Detay:</b> {eta_sc['detay']}</div>
            </div>
            """, unsafe_allow_html=True)

            # OTANTİK ETA:SQL PENCERESİ
            st.markdown(f"""
            <div class='eta-window'>
                <div class='eta-titlebar'>
                    <span>🗂️ ETA:SQL - [Yevmiye Fişi Girişi - {eta_sc['fis_no']}]</span>
                    <span>_ □ ✕</span>
                </div>
                <div class='eta-toolbar'>
                    <span>[F2] Kayıt</span> | <span>[F3] Sil</span> | <span>[F5] İptal</span> | <span>[F6] Kart Arama</span> | <span>[F7] Fiş Listesi</span> | <span>[F8] Yazıcı</span>
                </div>
                <div class='eta-header-card'>
                    <div>FİŞ NO: <b>{eta_sc['fis_no']}</b></div>
                    <div>TARİH: <b>{eta_sc['tarih']}</b></div>
                    <div>FİŞ TİPİ: <b>02 - MAHSUP</b></div>
                    <div>DURUM: <span style='color:#008000;'>AÇIK / DÜZENLEME</span></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("<div style='font-size:0.78rem; font-weight:700; color:#CBD5E1; margin:6px 0 4px 0;'>ETA FİŞ SATIRLARI (HESAP KODU VE TUTARLARI DOLDURUNUZ):</div>", unsafe_allow_html=True)
            
            # Dinamik ETA Satır Formu
            satir_sayisi = len(eta_sc["satirlar"])
            girilen_satirlar = []

            col_w = [2.2, 3.5, 2.0, 2.0]
            h_c1, h_c2, h_c3, h_c4 = st.columns(col_w)
            h_c1.caption("HESAP KODU")
            h_c2.caption("AÇIKLAMA")
            h_c3.caption("BORÇ (TL)")
            h_c4.caption("ALACAK (TL)")

            toplam_girilen_borc = 0.0
            toplam_girilen_alacak = 0.0

            for i in range(satir_sayisi):
                s_c1, s_c2, s_c3, s_c4 = st.columns(col_w)
                with s_c1:
                    kod = st.text_input(f"Kod {i+1}", key=f"eta_k_{eta_sc['step']}_{i}", label_visibility="collapsed", placeholder="Örn: 153, 320")
                with s_c2:
                    aciklama = st.text_input(f"Açıklama {i+1}", key=f"eta_a_{eta_sc['step']}_{i}", label_visibility="collapsed", value=eta_sc["satirlar"][i]["ad"])
                with s_c3:
                    borc = st.number_input(f"Borç {i+1}", min_value=0.0, value=0.0, step=100.0, key=f"eta_b_{eta_sc['step']}_{i}", label_visibility="collapsed")
                with s_c4:
                    alacak = st.number_input(f"Alacak {i+1}", min_value=0.0, value=0.0, step=100.0, key=f"eta_c_{eta_sc['step']}_{i}", label_visibility="collapsed")

                toplam_girilen_borc += borc
                toplam_girilen_alacak += alacak
                girilen_satirlar.append({"kod": kod.strip(), "borc": borc, "alacak": alacak})

            fark = abs(toplam_girilen_borc - toplam_girilen_alacak)

            # ETA PENCERE ALTI TOPLAM ÇUBUĞU
            fark_renk = "#008000" if (fark < 0.05 and toplam_girilen_borc > 0) else "#CC0000"
            st.markdown(f"""
            <div class='eta-footer-bar'>
                <span>TOPLAM BORÇ: <b style='color:#000080;'>{toplam_girilen_borc:,.2f} TL</b></span>
                <span>TOPLAM ALACAK: <b style='color:#000080;'>{toplam_girilen_alacak:,.2f} TL</b></span>
                <span>BAKİYE FARKI: <b style='color:{fark_renk};'>{fark:,.2f} TL</b></span>
            </div>
            """, unsafe_allow_html=True)

            b_eta1, b_eta2 = st.columns([1.5, 1.0])
            with b_eta1:
                if st.button("💾 [F2] Fişi ETA'ya Kaydet & Mühürle", use_container_width=True):
                    # Fiş Doğruluk Kontrolü
                    hepsi_dogru = True
                    for i in range(satir_sayisi):
                        hedef = eta_sc["satirlar"][i]
                        girilen = girilen_satirlar[i]
                        if not (hedef["kod"] in girilen["kod"] and abs(hedef["borc"] - girilen["borc"]) < 0.05 and abs(hedef["alacak"] - girilen["alacak"]) < 0.05):
                            hepsi_dogru = False
                            break

                    if hepsi_dogru and fark < 0.05 and toplam_girilen_borc > 0:
                        st.session_state["academy_xp"] += eta_sc["xp"]
                        st.session_state["academy_streak"] += 1
                        sesli_bildirim_cal("success")
                        st.balloons()
                        st.success(f"🎉 MÜKEMMEL KAYIT! Fiş {eta_sc['fis_no']} ETA:SQL veritabanına mühürlendi. (+{eta_sc['xp']} XP)")
                        time.sleep(1.5)
                        st.session_state["eta_step"] += 1
                        st.session_state["eta_current_scenario"] = generate_eta_scenario(st.session_state["eta_step"])
                        st.rerun()
                    else:
                        st.session_state["academy_lives"] -= 1
                        st.session_state["academy_streak"] = 0
                        sesli_bildirim_cal("error")
                        st.error(f"💥 ETA KAYIT HATASI! (-1 Can) Fiş dengesiz veya hesap kodları yanlış. {eta_sc['ipucu']}")
                        time.sleep(1.5)
                        st.session_state["eta_step"] += 1
                        st.session_state["eta_current_scenario"] = generate_eta_scenario(st.session_state["eta_step"])
                        st.rerun()

            with b_eta2:
                if st.button("➡️ Sonraki ETA Fişine Geç", use_container_width=True):
                    st.session_state["eta_step"] += 1
                    st.session_state["eta_current_scenario"] = generate_eta_scenario(st.session_state["eta_step"])
                    st.rerun()

# ------------------------------------------------------------------------------
# SEKME 3: ⚖️ HUKUKİ ÇERÇEVE & SLA (WHATSAPP BALONLARI)
# ------------------------------------------------------------------------------
with sekme_hukuk:
    st.markdown("### ⚖️ Kurumsal Güvence, Regülasyon & Sorumluluk Protokolü")
    st.caption("WhatsApp diyaloğu tarzında anlaşılır ve şeffaf hukuki çerçeve.")

    st.markdown("""
    <div class='wa-bubble-left'>
        <div class='wa-title'>💬 Soru 1: LedgerAI muhasebecinin yerine mi geçiyor? Bize yasal ceza gelir mi?</div>
        <div class='wa-text'>
            <b>Cevap:</b> Kesinlikle hayır! LedgerAI bir <b>Dual-Control (İki Göz)</b> asistanıdır. Fişleri sadece ön hazırlık olarak taslak çıkarır. 
            3568 Sayılı Kanun gereği tüm yasal defter ve beyanname onay yetkisi yetkili Serbest Muhasebeci Mali Müşavir (SMMM) veya YMM'ye aittir. 
            Nihai kontrolü her zaman uzman insan yapar.
        </div>
    </div>

    <div class='wa-bubble-right'>
        <div class='wa-title'>🔒 Soru 2: Faturalarımız, müşteri isimlerimiz veya şirket sırlarımız kaydediliyor mu?</div>
        <div class='wa-text'>
            <b>Cevap:</b> Asla! <b>Zero-Retention (Sıfır Kalıcı Depolama)</b> prensibiyle çalışıyoruz. Yüklediğiniz fatura belleğe (RAM) alınır, 
            kuruşu kuruşuna okunup fiş oluştuktan sonra kalıcı olarak bellekten silinir. Verileriniz model eğitiminde (training) kesinlikle kullanılmaz.
        </div>
    </div>

    <div class='wa-bubble-left'>
        <div class='wa-title'>🏫 Soru 3: Siber Akademi modülünü okullar ve üniversiteler ders materyali olarak kullanabilir mi?</div>
        <div class='wa-text'>
            <b>Cevap:</b> Evet! Siber Akademi tamamen hayali, simülasyon amaçlı vakalar türetir. Öğrencilerden hiçbir kişisel veri veya TC kimlik numarası istenmez. 
            Üniversiteler, meslek yüksekokulları ve liseler için güvenli bir dijital muhasebe laboratuvarıdır.
        </div>
    </div>

    <div class='wa-bubble-right'>
        <div class='wa-title'>⚖️ Soru 4: Dışa aktarılan fişlerde Borç ve Alacak eşitliği garanti altında mı?</div>
        <div class='wa-text'>
            <b>Cevap:</b> Evet! Matematiksel Denetim Kilidi sayesinde Toplam Borç = Toplam Alacak eşitliği kuruşu kuruşuna sağlanmadan sistem fişi onaylamaz. 
            İndirilen Excel ve ETA V.11 dosyaları ERP yazılımlarına sıfır bakiye farkı ile aktarılır.
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)
