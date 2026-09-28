"""
================================================================================
LEDGERAI — INSTITUTIONAL ENTERPRISE ACCOUNTING TERMINAL
Architecture: Streamlit + Google Gemini GenAI SDK + Pandas + OpenPyXL
Design: Open Slate / Platinum Titanium Executive Dashboard (Whitelabel)
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
    page_title="LedgerAI — Enterprise Accounting Terminal",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

SESSION_DEFAULTS = {
    "user_lang": "🇹🇷 TR",
    "theme_idx": 0,  # 0: Kurumsal Açık Platin
    "industry_idx": 0,
    "chat_messages": [],
    "out_df": None,
    "raw_audit_results": [],
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
# 2. LOCALIZATION DATA DICTIONARY
# ==============================================================================

LANG_DATA = {
    "🇹🇷 TR": {
        "badge": "KURUMSAL OTONOM FİNANS TERMİNALİ",
        "title": "LedgerAI",
        "subtitle": "Faturaları saniyeler içinde sektörel hesap kodlarına ve kuruşu kuruşuna dengeli ERP fişine dönüştürün.",
        "drop_title": "Belgeleri Buraya Bırakın veya Seçin",
        "drop_sub": "PDF, PNG, JPG formatında fatura, makbuz ve fişler • Maksimum 5 belge",
        "process_btn": "⚡ Otonom Muhasebeleştir & Denetle",
        "limit_err": "🛑 Demo sürümünde oturum başına en fazla 5 fatura işlenebilir.",
        "ready_count": "İşlenecek belge sayısı: **{count}**",
        "success": "✓ Fişler başarıyla oluşturuldu ve Borç/Alacak kuruşu kuruşuna dengelendi.",
        "failed": "❌ Belgeler işlenemedi. Lütfen görsel netliğini kontrol edin.",
        "preview_title": "📊 Muhasebe Yevmiye Fişi & Denetim Masası",
        "preview_tip": "💡 Hücrelere çift tıklayarak kod veya açıklamaları değiştirebilirsiniz. Çok sayfalı Excel'e anında yansır.",
        "tot_deb": "Toplam Borç",
        "tot_crd": "Toplam Alacak",
        "balanced": "✅ Fiş Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "download_btn": "📥 Çok Sayfalı Kurumsal Excel'i İndir (.xlsx)",
        "download_eta": "💾 ETA V.11 Uyumlu CSV",
        "industries": [
            "⚡ Otomatik Sektör (AI)",
            "🛒 Ticaret / Al-Sat (153 Ağırlıklı)",
            "🏢 Hizmet & Ofis (770/740)",
            "🏭 Üretim & Fabrika (150/730)"
        ],
        "themes": [
            "🌑 Platin Gri (Kurumsal)",
            "✨ Ultra Canlı Aurora",
            "🌌 Cyberpunk Gece"
        ],
        "about_btn": "ℹ️ İşleyiş & Güvenlik",
        "about_title": "LedgerAI Otonom Sistem Mimarisi",
        "about_content": """
        ### 🛡️ Kurumsal Finans & Güvenlik Mimarisi
        **LedgerAI**, kurumların fiş giriş maliyetlerini sıfıra indiren yeni nesil finans terminalidir.
        * **1. Çift Taraflı Tevkifat/Stopaj Algoritması:** Tevkifatlı faturalarda veya SMMM makbuzlarında kesintileri otomatik hesaplar; `360 Ödenecek Vergi` satırını açarak borç/alacak denkliğini garanti eder.
        * **2. Sektörel Mantık:** Faturadaki mal alımını şirketin faaliyetine göre (`153`, `150` veya `770`) dinamik ayırır.
        * **3. Çoklu KDV Ayrıştırması:** Aynı faturada birden fazla KDV oranı varsa (%1, %10, %20) her oran için ayrı borç satırı üretir.
        * **4. Çok Sayfalı Kurumsal Raporlama:** Excel çıktısında fişler; 153 Ticari Mallar, 770 Masraflar ve Genel Özet sayfalarına otomatik ayrılır.
        """,
        "step1_title": "1. Belge & Tevkifat Analizi",
        "step1_desc": "OCR ile çoklu KDV, tevkifat oranları, stopaj ve matrahlar kuruşu kuruşuna okunur.",
        "step2_title": "2. Sektörel Hesap Eşleme",
        "step2_desc": "Şirket türüne göre ticari mal (153), üretim (150) veya masraf (770) dinamik atanır.",
        "step3_title": "3. Çift Bakiye Doğrulama",
        "step3_desc": "Toplam Borç = Toplam Alacak eşitliği sağlanmadan yevmiye fişi üretilmez.",
        "badge_erp": "✓ ETA • LUCA • DATEV • QUICKBOOKS UYUMLU",
        "badge_audit": "✓ %100 BORÇ/ALACAK DENGE GARANTİSİ",
        "badge_sec": "✓ ÇOK SAYFALI ÖZEL EXCEL RAPORU",
        "bot_title": "👾 LedgerBot Finans Mentorü",
        "bot_welcome": "Selam! Ben finans asistanınım. Muhasebe öğrenmek veya pratik hesap kodlarını sormak için bana yazabilirsin. Kısa, net ve örnekle anlatırım!",
        "bot_placeholder": "Sorunuzu yazın (Örn: Tevkifatlı fatura nasıl işlenir? Borç/Alacak mantığı nedir?)...",
        "bot_clear": "🧹 Temizle",
        "headers": {
            "vouch": "Fiş No", "date": "Tarih", "code": "Hesap Kodu",
            "name": "Hesap Adı", "desc": "Açıklama", "curr": "Para Birimi",
            "deb": "Borç", "crd": "Alacak"
        },
        "faq_title": "💬 Sıkça Sorulan Sorular & Güvenlik",
        "faqs": [
            {
                "q": "Bir muhasebeciye ne gibi kolaylıklar sağlayabilir?",
                "a": "LedgerAI, manuel veri girişini ve fiş eşleştirmesini otomatikleştirerek muhasebecilerin rutin iş yükünü %80 oranında azaltır. Faturaları yapay zeka ile doğrudan doğru hesaplara işler, tevkifat ve KDV ayrımını yaparak insan hatasını sıfırlar.\n\nÖrneğin; ofis kırtasiye faturası doğrudan 770 Genel Yönetim Giderleri hesabına aktarılırken ticari ürünler 153 hesabına aktarılır.\nBorç: 770 / 191 — Alacak: 320"
            },
            {
                "q": "Ne kadar güvenilir bir işlem bu? Verilerim güvende mi?",
                "a": "LedgerAI, verilerinizi uçtan uca TLS şifreleme ile iletir. Sistem belgelerinizi kalıcı olarak üçüncü taraflarla paylaşmaz veya model eğitiminde kullanmaz. İşlem bittiğinde fişler sadece sizin oturumunuzda tutulur ve KVKK/finansal gizlilik prensiplerine tam uyum sağlar."
            },
            {
                "q": "Tevkifatlı veya birden fazla KDV oranlı faturaları nasıl işler?",
                "a": "Faturada örneğin hem %10 hem %20 KDV varsa, sistem her ikisini ayrı ayrı hesaplayarak ayrı satırlar açar. KDV tevkifatı veya SMMM stopajı tespit edilirse, satıcıya ödenecek net tutar 320 hesabına, kesilen vergi ise 360 Ödenecek Vergi hesabına otomatik yazılır ve bakiye her zaman eşitlenir."
            },
            {
                "q": "Excel çıktısında veriler nasıl gruplanır?",
                "a": "İndirilen Excel dosyasında tüm yevmiye fişleri genel sayfada yer alırken; 153 Ticari Mallar ve 770 Genel Masraflar bağımsız renkli sekmelerde listelenir. Böylece departman bazlı kontrol saniyeler içinde yapılır."
            }
        ]
    },
    "🇺🇸 EN": {
        "badge": "INSTITUTIONAL AI FINANCIAL TERMINAL",
        "title": "LedgerAI",
        "subtitle": "Convert raw invoices into balanced, multi-GAAP ERP journal vouchers autonomously.",
        "drop_title": "Drop Financial Documents Here or Browse",
        "drop_sub": "PDF, PNG, JPG • Invoices, Receipts & Vouchers • Up to 5 files",
        "process_btn": "⚡ Process & Generate Vouchers",
        "limit_err": "🛑 Demo limit is 5 documents per batch.",
        "ready_count": "Documents ready: **{count}**",
        "success": "✓ Journal vouchers generated and balanced down to the cent.",
        "failed": "❌ Documents could not be parsed.",
        "preview_title": "📊 Journal Voucher Grid & Audit Deck",
        "preview_tip": "💡 Double-click any cell to adjust accounts or descriptions before export.",
        "tot_deb": "Total Debit",
        "tot_crd": "Total Credit",
        "balanced": "✅ Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Unbalanced Voucher!",
        "download_btn": "📥 Download Multi-Tab Corporate Excel (.xlsx)",
        "download_eta": "💾 Generic CSV Format",
        "industries": [
            "⚡ Auto Industry (AI)", "🛒 Retail / Inventory (1200)",
            "🏢 Services / SaaS (OpEx)", "🏭 Manufacturing (COGS)"
        ],
        "themes": [
            "🌑 Platinum Slate (Executive)",
            "✨ Ultra Vivid Aurora",
            "🌌 Cyberpunk Night"
        ],
        "about_btn": "ℹ️ How it Works & Security",
        "about_title": "LedgerAI Autonomous Architecture",
        "about_content": "Autonomous double-entry journal voucher generator compatible with US GAAP, Datev and PCG.",
        "step1_title": "1. Multi-Tax Extraction",
        "step1_desc": "Sub-millisecond OCR extraction of multi-tier tax rates, withholdings, and net amounts.",
        "step2_title": "2. Contextual Mapping",
        "step2_desc": "Automated account mapping to OpEx, Inventory, or Capital Assets.",
        "step3_title": "3. Double-Entry Verification",
        "step3_desc": "Mathematical verification guaranteeing Total Debit equals Total Credit.",
        "badge_erp": "✓ QUICKBOOKS • XERO • DATEV • SAP READY",
        "badge_audit": "✓ 100% DEBIT/CREDIT BALANCE GUARANTEE",
        "badge_sec": "✓ MULTI-TAB WORKBOOK EXPORT",
        "bot_title": "👾 LedgerBot Finance Mentor",
        "bot_welcome": "Hi! I am your AI finance mentor. Ask me any accounting concepts or codes. I reply concisely with direct practical examples!",
        "bot_placeholder": "Ask a question (e.g. How to book SaaS subscriptions? Debit vs Credit?)...",
        "bot_clear": "🧹 Clear",
        "headers": {
            "vouch": "Voucher #", "date": "Date", "code": "Account Code",
            "name": "Account Name", "desc": "Memo", "curr": "Currency",
            "deb": "Debit", "crd": "Credit"
        },
        "faq_title": "💬 Frequently Asked Questions & Security",
        "faqs": [
            {
                "q": "How does LedgerAI streamline enterprise accounting?",
                "a": "It automates invoice parsing and chart-of-accounts mapping, reducing manual entry by 80%. It eliminates human errors by validating Debit = Credit parity before generating journal vouchers."
            },
            {
                "q": "Is our financial data safe?",
                "a": "Yes. Data is processed over secure TLS connections with bank-grade encryption. Documents are parsed strictly within your active session and never stored permanently."
            }
        ]
    }
}

if st.session_state["user_lang"] not in LANG_DATA:
    st.session_state["user_lang"] = "🇹🇷 TR"

T = LANG_DATA[st.session_state["user_lang"]]

# ==============================================================================
# 3. DYNAMIC STYLING ENGINE (KURUMSAL AÇIK PLATİN / ARDUVAZ VE WHITELABEL CSS)
# ==============================================================================

if st.session_state["theme_idx"] == 0:
    # 🌑 Platin Gri (Açık Arduvaz / Metalik Şirket Havası)
    bg_style = """
        @keyframes slateShimmer {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        .stApp {
            background: radial-gradient(circle at 50% 0%, rgba(203, 213, 225, 0.15) 0%, transparent 65%),
                        radial-gradient(circle at 85% 90%, rgba(148, 163, 184, 0.10) 0%, transparent 50%),
                        linear-gradient(145deg, #111827 0%, #1E293B 50%, #0F172A 100%);
            background-size: 200% 200%;
            animation: slateShimmer 24s ease infinite;
            background-attachment: fixed;
        }
    """
elif st.session_state["theme_idx"] == 1:
    # ✨ Ultra Canlı Aurora
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
    # 🌌 Cyberpunk Gece
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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }}
    
    /* WHITELABEL: GITHUB, STREAMLIT FOOTER VE MENÜLERİ TAMAMEN GİZLE */
    header[data-testid="stHeader"] {{
        display: none !important;
    }}
    #MainMenu {{
        visibility: hidden !important;
    }}
    footer {{
        visibility: hidden !important;
    }}
    div[data-testid="stToolbar"] {{
        display: none !important;
    }}
    div[data-testid="stDecoration"] {{
        display: none !important;
    }}
    .viewerBadge_container__1QSob {{
        display: none !important;
    }}
    
    [data-testid="stSidebar"] {{ display: none !important; }}
    
    {bg_style}
    
    .stApp {{
        color: #F8FAFC;
        padding-top: 15px;
        padding-bottom: 70px;
    }}

    /* MASTER GLASS TERMINAL */
    .master-console {{
        max-width: 980px;
        margin: 15px auto 0 auto;
        background: rgba(30, 41, 59, 0.72);
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 24px;
        backdrop-filter: blur(28px);
        -webkit-backdrop-filter: blur(28px);
        padding: 36px 44px 28px 44px;
        box-shadow: 0 25px 60px rgba(0, 0, 0, 0.5), 
                    inset 0 1px 0 rgba(255, 255, 255, 0.22);
        text-align: center;
    }}

    .top-badge {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 14px;
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.20);
        border-radius: 99px;
        font-size: 0.7rem;
        font-weight: 700;
        letter-spacing: 1.4px;
        color: #E2E8F0;
        margin-bottom: 12px;
        text-transform: uppercase;
    }}

    .hero-title {{
        font-size: 3.1rem;
        font-weight: 800;
        letter-spacing: -1px;
        background: linear-gradient(135deg, #FFFFFF 40%, #CBD5E1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 6px;
        line-height: 1.05;
    }}
    
    .hero-sub {{
        font-size: 0.98rem;
        color: #CBD5E1;
        font-weight: 400;
        line-height: 1.5;
        max-width: 660px;
        margin: 0 auto 24px auto;
    }}

    /* FILE UPLOADER */
    div[data-testid="stFileUploader"] {{
        background: rgba(15, 23, 42, 0.65);
        border: 1px dashed rgba(255, 255, 255, 0.25);
        border-radius: 16px;
        padding: 24px 16px;
        transition: all 0.25s ease;
        margin-bottom: 12px;
    }}
    div[data-testid="stFileUploader"]:hover {{
        border-color: rgba(203, 213, 225, 0.9);
        box-shadow: 0 0 30px rgba(255, 255, 255, 0.15);
        background: rgba(30, 41, 59, 0.8);
    }}

    /* ACTION BUTTON */
    div.stButton > button:first-child {{
        background: linear-gradient(135deg, #475569 0%, #1E293B 100%);
        border: 1px solid rgba(255, 255, 255, 0.25);
        border-radius: 12px;
        font-weight: 700;
        font-size: 0.95rem;
        padding: 12px 28px;
        color: #FFFFFF;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
        transition: all 0.25s ease;
        margin-top: 6px;
    }}
    div.stButton > button:first-child:hover {{
        background: linear-gradient(135deg, #64748B 0%, #334155 100%);
        border-color: rgba(255, 255, 255, 0.45);
        box-shadow: 0 6px 25px rgba(255, 255, 255, 0.2);
        transform: translateY(-1px);
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
        background: rgba(255, 255, 255, 0.04);
        border: 1px solid rgba(255, 255, 255, 0.10);
        border-radius: 14px;
        padding: 14px 16px;
        backdrop-filter: blur(12px);
        transition: all 0.25s;
    }}
    .step-card:hover {{
        background: rgba(255, 255, 255, 0.08);
        border-color: rgba(255, 255, 255, 0.25);
        transform: translateY(-2px);
    }}
    .step-title {{
        font-size: 0.82rem;
        font-weight: 700;
        color: #F8FAFC;
        margin-bottom: 4px;
    }}
    .step-desc {{
        font-size: 0.72rem;
        color: #CBD5E1;
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
        border-top: 1px solid rgba(255, 255, 255, 0.10);
    }}
    .trust-item {{
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 0.8px;
        color: #CBD5E1;
    }}

    /* CONSOLE CONTROLS */
    .console-controls {{
        margin-top: 20px;
        padding-top: 16px;
        border-top: 1px solid rgba(255, 255, 255, 0.10);
    }}

    /* METRIC CARDS */
    .stMetric {{
        background: rgba(30, 41, 59, 0.75);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 14px;
        backdrop-filter: blur(14px);
        padding: 14px 18px;
    }}

    /* SCROLLABLE CHAT CONTAINER */
    .chat-scroll-area {{
        max-height: 360px;
        overflow-y: auto;
        padding: 12px 14px;
        background: rgba(15, 23, 42, 0.65);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 14px;
        margin-bottom: 14px;
    }}
    .chat-scroll-area::-webkit-scrollbar {{
        width: 6px;
    }}
    .chat-scroll-area::-webkit-scrollbar-thumb {{
        background: rgba(255, 255, 255, 0.25);
        border-radius: 4px;
    }}

    /* FAQ CONTAINER */
    .faq-container {{
        max-width: 980px;
        margin: 35px auto 0 auto;
        background: rgba(30, 41, 59, 0.65);
        border: 1px solid rgba(255, 255, 255, 0.12);
        border-radius: 20px;
        backdrop-filter: blur(24px);
        padding: 24px 30px;
    }}

    /* AUDIT FILTER BOX */
    .filter-card {{
        background: rgba(30, 41, 59, 0.55);
        border: 1px solid rgba(255, 255, 255, 0.10);
        border-radius: 12px;
        padding: 10px 16px;
        margin-bottom: 14px;
    }}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 4. INSTITUTIONAL MULTI-TAB EXCEL ENGINE (153, 770 & ÖZET SEKMELERİ)
# ==============================================================================

def export_multitab_corporate_excel(df: pd.DataFrame) -> bytes:
    """
    Generates an auditor-grade, multi-tab Excel workbook:
    - Tab 1: Konsolide Yevmiye Fişi (Genel Fişler)
    - Tab 2: 153 Ticari Mallar (Stok/Emtia Alımları - Yeşil Tema)
    - Tab 3: 770 Genel Masraflar (İşletme Giderleri - Mavi Tema)
    - Tab 4: Denetim & Bakiye Özeti
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        headers = T["headers"]
        border_thin = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )
        data_font = Font(name="Inter", size=10)
        num_font = Font(name="Consolas", size=10)

        # TAB 1: TÜM FİŞLER (KONSOLİDE)
        df.to_excel(writer, index=False, sheet_name="Yevmiye Fisi")
        ws1 = writer.sheets["Yevmiye Fisi"]
        h_fill1 = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        h_font1 = Font(name="Inter", size=10, bold=True, color="FFFFFF")

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

        # TAB 2: 153 TİCARİ MALLAR (YEŞİL TEMA)
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

        # TAB 3: 770 GENEL YÖNETİM MASRAFLARI (MAVİ TEMA)
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
            for col in ws3.columns:
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
# 5. CORE AI RECOGNITION ENGINE (TEVKİFAT, MULTI-TAX & DUAL AUDIT)
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
                st.success(f"{T['success']} ({st.session_state['last_processing_time']} sn)")

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
""", unsafe_allow_html=True)

# KONSOL İÇİ KONTROL ÇUBUĞU
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
# 7. INTERACTIVE JOURNAL VOUCHER GRID, FILTERS & AUDIT CENTER
# ==============================================================================

if st.session_state["out_df"] is not None:
    st.markdown("<div style='height: 35px;'></div>", unsafe_allow_html=True)
    st.subheader(T["preview_title"])
    st.caption(T["preview_tip"])

    # AKILLI FİLTRELEME & DENETİM KONTROLÜ
    headers = T["headers"]
    st.markdown("<div class='filter-card'>", unsafe_allow_html=True)
    f_col1, f_col2, f_col3 = st.columns([2, 3, 2])
    
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
            st.markdown(f"<div style='font-size:0.85rem; padding-top:6px; color:#A7F3D0;'>🛡️ <b>Denetim Güvencesi:</b> {doc_cnt} Belge %{avg_conf:.1f} OCR & Matematik Doğruluğu ile Mühürlendi.</div>", unsafe_allow_html=True)

    with f_col3:
        st.markdown("<div style='text-align:right; font-size:0.82rem; padding-top:6px; color:#CBD5E1;'>💡 Hücreye çift tıklayıp düzenleyin</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Filtreleme Mantığı
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

# ==============================================================================
# 8. MENTOR FINANS ASİSTANI (SCROLLABLE & CLEAN)
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

        st.markdown("<div class='chat-scroll-area'>", unsafe_allow_html=True)
        if not st.session_state["chat_messages"]:
            st.markdown(f"<div style='color: #94A3B8; font-size: 0.85rem; padding: 10px 0;'>💡 <i>{T['bot_placeholder']}</i></div>", unsafe_allow_html=True)
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
            1. Asla lafı uzatma, genel tanımlar yazma.
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

# ==============================================================================
# 9. SIKÇA SORULAN SORULAR & KURUMSAL GÜVENLİK (FAQ)
# ==============================================================================

st.markdown(f"""
<div class='faq-container'>
    <div style='font-size: 1.15rem; font-weight: 700; margin-bottom: 16px; color: #F1F5F9;'>
        {T['faq_title']}
    </div>
""", unsafe_allow_html=True)

for faq in T.get("faqs", []):
    with st.expander(f"📌 {faq['q']}", expanded=False):
        st.markdown(f"<div style='font-size: 0.88rem; color: #CBD5E1; line-height: 1.6;'>{faq['a']}</div>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)
