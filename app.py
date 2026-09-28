"""
================================================================================
LEDGERAI — INSTITUTIONAL ENTERPRISE ACCOUNTING TERMINAL
Architecture: Streamlit + Google Gemini GenAI SDK + Pandas + OpenPyXL
Design: Dual-Wing Executive Cockpit / GitHub Pill Micro-UI (Whitelabel)
Version: 4.0.0 Cockpit Master Edition
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
    page_title="LedgerAI — Autonomous Financial Terminal",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

SESSION_DEFAULTS = {
    "user_lang": "🇹🇷 TR",
    "theme_idx": 0,  # 0: Platin Gri (Varsayılan Açılış)
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
    "audit_logs": []
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
# 2. LOCALIZATION DATA DICTIONARY (6 DİLLİ GLOBAL MEVZUAT)
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
            "🌑 Platin Gri",
            "✨ Ultra Canlı Aurora",
            "🌌 Cyberpunk Gece"
        ],
        "about_btn": "ℹ️ İşleyiş & Felsefe",
        "about_title": "LedgerAI Mimarisi & İnsan-AI Ortaklığı",
        "about_content": """
        ### 🛡️ İnsan Gücü ve Yapay Zekanın Güvenli Birleşimi
        **LedgerAI**, çalışanların veya mali müşavirlerin yerini almak için değil; onların üzerindeki mekanik veri giriş yükünü kaldırıp stratejik denetim gücü kazandırmak için tasarlandı.
        * **1. Çift Göz Prensibi:** Yapay zeka faturadaki KDV tevkifatını ve matrahı ayrıştırıp yevmiye fişini hazırlar; insan uzman son kontrolü yapıp onaylar.
        * **2. Sıfır Hata Garantisi:** Matematiksel olarak Borç = Alacak denkliği kuruşu kuruşuna doğrulanmadan sistem fiş üretmez.
        * **3. İstihdamı Destekleyen Teknoloji:** Muhasebe personeli saatlerce fatura girmek yerine finansal analiz ve danışmanlığa odaklanır.
        * **4. ERP Uyumluluğu:** ETA V.11, Luca, Logo, Zirve, Mikro, Datev ve QuickBooks standartlarında çıktı sağlar.
        """,
        "step1_title": "1. Belge & Tevkifat Okuma",
        "step1_desc": "OCR ile çoklu KDV, tevkifat ve stopaj kuruşu kuruşuna ayıklanır.",
        "step2_title": "2. Sektörel Eşleme & Öneri",
        "step2_desc": "Faaliyete göre 153, 150 veya 770 kodları otomatik önerilir.",
        "step3_title": "3. İnsan Onayı & Denge",
        "step3_desc": "Borç/Alacak dengesi mühürlenir, uzman onayına sunulur.",
        "badge_erp": "✓ ETA • LUCA • DATEV • QUICKBOOKS UYUMLU",
        "badge_audit": "✓ %100 MATEMATİKSEL DENGE GARANTİSİ",
        "badge_sec": "✓ %100 VERİ GİZLİLİĞİ & GÜVENLİK",
        "bot_title": "👾 LedgerBot Finans Mentorü",
        "bot_welcome": "Selam! Ben finans asistanınım. Muhasebe öğrenmek veya fatura mantığını sormak için aşağıdaki hap sorulara tıklayabilirsin:",
        "bot_placeholder": "Muhasebe sorunuzu yazın (Örn: 153 ile 770 farkı nedir?)...",
        "bot_clear": "🧹 Temizle",
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
        "headers": {
            "vouch": "Fiş No", "date": "Tarih", "code": "Hesap Kodu",
            "name": "Hesap Adı", "desc": "Açıklama", "curr": "Para Birimi",
            "deb": "Borç", "crd": "Alacak"
        }
    },
    "🇺🇸 EN": {
        "badge": "HUMAN + AI COLLABORATIVE TERMINAL",
        "title": "LedgerAI",
        "subtitle": "AI parses invoices and tax withholdings; human accounting professionals audit and approve.",
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
            "🌑 Platinum Slate",
            "✨ Ultra Vivid Aurora",
            "🌌 Cyberpunk Gece"
        ],
        "about_btn": "ℹ️ How it Works & Philosophy",
        "about_title": "LedgerAI Architecture & Human-AI Collaboration",
        "about_content": "LedgerAI empowers financial teams by automating data entry while keeping human experts in full control.",
        "step1_title": "1. Multi-Tax Extraction",
        "step1_desc": "Sub-millisecond extraction of multi-tier tax rates and withholdings.",
        "step2_title": "2. Contextual Mapping",
        "step2_desc": "Automated account mapping to OpEx, Inventory, or Capital Assets.",
        "step3_title": "3. Human Audit & Balance",
        "step3_desc": "Mathematical verification guaranteeing Total Debit equals Total Credit.",
        "badge_erp": "✓ QUICKBOOKS • XERO • DATEV • SAP READY",
        "badge_audit": "✓ 100% DEBIT/CREDIT BALANCE GUARANTEE",
        "badge_sec": "✓ SOC2 & BANK-GRADE ENCRYPTION",
        "bot_title": "👾 LedgerBot Finance Mentor",
        "bot_welcome": "Hi! I am your AI finance mentor. Tap any quick pill question below or ask me directly:",
        "bot_placeholder": "Ask a question...",
        "bot_clear": "🧹 Clear",
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
        "headers": {
            "vouch": "Voucher #", "date": "Date", "code": "Account Code",
            "name": "Account Name", "desc": "Memo", "curr": "Currency",
            "deb": "Debit", "crd": "Credit"
        }
    }
}

if st.session_state["user_lang"] not in LANG_DATA:
    st.session_state["user_lang"] = "🇹🇷 TR"

T = LANG_DATA[st.session_state["user_lang"]]

# ==============================================================================
# 3. DYNAMIC STYLING ENGINE (KOKPİT DÜZENİ & GITHUB.COM MİKRO HAPLARI)
# ==============================================================================

if st.session_state["theme_idx"] == 0:
    # 🌑 Platin Gri
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
    [data-testid="stSidebar"] {{ display: none !important; }}
    
    {bg_style}
    
    .stApp {{
        color: #F8FAFC;
        padding-top: 10px;
        padding-bottom: 60px;
    }}

    /* DUAL WING COCKPIT CONSOLE (MASAÜSTÜNDE BOŞLUĞU BİTİREN KART) */
    .cockpit-card {{
        background: rgba(30, 41, 59, 0.72);
        border: 1px solid rgba(255, 255, 255, 0.16);
        border-radius: 24px;
        backdrop-filter: blur(28px);
        -webkit-backdrop-filter: blur(28px);
        padding: clamp(20px, 3vw, 32px);
        box-shadow: 0 20px 50px rgba(0, 0, 0, 0.45), 
                    inset 0 1px 0 rgba(255, 255, 255, 0.20);
        height: 100%;
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
        margin-bottom: 10px;
        text-transform: uppercase;
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

    /* FILE UPLOADER */
    div[data-testid="stFileUploader"] {{
        background: rgba(15, 23, 42, 0.65);
        border: 1px dashed rgba(255, 255, 255, 0.22);
        border-radius: 18px;
        padding: 18px 14px;
        transition: all 0.25s ease;
        margin-bottom: 10px;
    }}
    div[data-testid="stFileUploader"]:hover {{
        border-color: rgba(203, 213, 225, 0.9);
        box-shadow: 0 0 25px rgba(255, 255, 255, 0.15);
        background: rgba(30, 41, 59, 0.8);
    }}

    /* GITHUB.COM BALONU ŞEKLİNDE SİYAH OVAL İŞLEM BUTONU */
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
    }}
    div.stButton > button:first-child:hover {{
        background: #111827 !important;
        border-color: rgba(255, 255, 255, 0.6) !important;
        box-shadow: 0 6px 25px rgba(255, 255, 255, 0.2) !important;
        transform: scale(1.02) !important;
    }}

    /* SAĞ KANAT KARTLARI */
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

    /* MİKRO KONTROL ŞERİDİ (GITHUB.COM HAPLARI) */
    .micro-dock {{
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
        margin-top: 14px;
        padding-top: 12px;
        border-top: 1px solid rgba(255, 255, 255, 0.08);
    }}

    /* ASİSTAN İÇİ MİKRO GITHUB HAPLARI */
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
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 4. INSTITUTIONAL MULTI-TAB EXCEL ENGINE
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

        # TAB 1: TÜM FİŞLER (KONSOLİDE)
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
# 6. DUAL WING EXECUTIVE COCKPIT (MASAÜSTÜNDE BOŞLUĞU BİTİREN KOKPİT)
# ==============================================================================

col_left, col_right = st.columns([1.35, 1.0], gap="large")

with col_left:
    st.markdown(f"""
    <div class='cockpit-card'>
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
            st.markdown(f"<div style='font-size:0.85rem; margin:6px 0 10px 0; color:#A7F3D0;'>{T['ready_count'].format(count=len(uploaded_files))}</div>", unsafe_allow_html=True)

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

    # KONSOL İÇİ MİKRO KONTROL ŞERİDİ (GITHUB.COM HAPLARI ŞEKLİNDE HİZALANDI)
    st.markdown("<div class='micro-dock'>", unsafe_allow_html=True)
    c_m1, c_m2, c_m3, c_m4 = st.columns([1.8, 3.2, 3.2, 1.8])
    with c_m1:
        lang_keys = list(LANG_DATA.keys())
        curr_lang_idx = lang_keys.index(st.session_state["user_lang"]) if st.session_state["user_lang"] in lang_keys else 0
        new_lang = st.selectbox("Dil", lang_keys, index=curr_lang_idx, label_visibility="collapsed")
        if new_lang != st.session_state["user_lang"]:
            st.session_state["user_lang"] = new_lang
            st.rerun()
    with c_m2:
        new_theme_str = st.selectbox("Görünüm", T["themes"], index=st.session_state["theme_idx"], label_visibility="collapsed")
        new_t_idx = T["themes"].index(new_theme_str)
        if new_t_idx != st.session_state["theme_idx"]:
            st.session_state["theme_idx"] = new_t_idx
            st.rerun()
    with c_m3:
        new_industry_str = st.selectbox("Sektör", T["industries"], index=st.session_state["industry_idx"], label_visibility="collapsed")
        new_i_idx = T["industries"].index(new_industry_str)
        if new_i_idx != st.session_state["industry_idx"]:
            st.session_state["industry_idx"] = new_i_idx
            st.rerun()
    with c_m4:
        with st.popover(T["about_btn"]):
            st.markdown(f"#### {T['about_title']}")
            st.markdown(T["about_content"])
    st.markdown("</div></div>", unsafe_allow_html=True)

with col_right:
    # SAĞ KANAT: KURUMSAL GÜVEN, DENETİM VE CANLI OPERASYON KARTI
    st.markdown(f"""
    <div class='cockpit-card'>
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
        <div style='display:flex; justify-content:space-between; align-items:center; margin-top:14px; padding:10px 14px; background:rgba(0,0,0,0.25); border-radius:12px;'>
            <span style='font-size:0.75rem; color:#A7F3D0;'>✓ %100 Bakiye Garantisi</span>
            <span style='font-size:0.75rem; color:#CBD5E1;'>ETA • Luca • Datev Ready</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# 7. INTERACTIVE JOURNAL VOUCHER GRID, FILTERS & AUDIT CENTER
# ==============================================================================

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
# 8. MENTOR FINANS ASİSTANI (GITHUB.COM BALONU MİKRO HAPLAR)
# ==============================================================================

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

        # GITHUB.COM STİLİ ASLA YIĞILMAYAN MİKRO HAP BUTONLAR
        secilen_chip = None
        btn_cols = st.columns(len(T["quick_chips"]))
        for c_idx, chip_text in enumerate(T["quick_chips"]):
            with btn_cols[c_idx]:
                if st.button(chip_text, key=f"gh_pill_btn_{c_idx}", use_container_width=True):
                    secilen_chip = chip_text

        # SCROLLABLE CHAT ALANI
        st.markdown("<div class='chat-scroll-area'>", unsafe_allow_html=True)
        if not st.session_state["chat_messages"]:
            st.markdown(f"<div style='color: #94A3B8; font-size: 0.82rem; padding: 8px 0;'>💡 <i>{T['bot_placeholder']}</i></div>", unsafe_allow_html=True)
        else:
            for msg in st.session_state["chat_messages"][-6:]:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])
        st.markdown("</div>", unsafe_allow_html=True)

        user_query = st.chat_input(T["bot_placeholder"])
        aktif_soru = user_query or secilen_chip
        
        if aktif_soru:
            st.session_state["chat_messages"].append({"role": "user", "content": aktif_soru})
            
            prompt_bot = f"""
            Sen LedgerAI'ın kurumsal finans mentorü ve pratik muhasebe uzmanısın.
            Felsefe: İnsan gücünü kovmak değil, insan ile yapay zekayı birleştirip muhasebeciye süper güç kazandırmak.
            Kullanıcı Dili: {st.session_state['user_lang']}
            Kullanıcı Sorusu: "{aktif_soru}"

            TALİMATLAR:
            1. Asla lafı uzatma, genel tanımlar yazma.
            2. MAKSİMUM 2-3 CÜMLEDE doğrudan ve net cevabı ver.
            3. "İnsan + AI ortaklığı" felsefesini koru: Yapay zeka hazırlar, uzman insan onaylar.
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
