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
    page_title="LedgerAI - Global Accounting Automation", 
    page_icon="💼", 
    layout="wide"
)

# API İstemcisi
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# --- DİL VE LOKALİZASYON METİNLERİ ---
TEXTS = {
    "EN": {
        "title": "💼 LedgerAI - Intelligent Invoice & Receipt Engine",
        "subtitle": "Upload invoices or receipts; AI automatically categorizes expenses, balances debits & credits, and generates formatted ERP/accounting imports.",
        "uploader_label": "Upload Invoices or Receipts (PDF, PNG, JPG)",
        "limit_warning": "🛑 Free demo limit is 2 files per batch. Contact us for custom ERP integrations and enterprise access.",
        "files_count": "Documents ready to process: **{count}**",
        "btn_process": "🚀 Process with AI Accounting Engine",
        "processing": "Processing ({current}/{total}): {filename}...",
        "retrying": "High demand, retrying ({wait}s)...",
        "completed": "Accounting vouchers generated successfully!",
        "preview_header": "📊 Generated Accounting Journal Voucher (Editable)",
        "preview_tip": "💡 Double-click any cell to modify accounts or amounts before downloading. Changes reflect instantly.",
        "total_debit": "Total Debit",
        "total_credit": "Total Credit",
        "balanced": "✅ Voucher Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Balance Difference Detected!",
        "btn_download": "📥 Download Formatted Journal Excel (.xlsx)",
        "sidebar_config": "⚙️ Accounting & Region",
        "lang_select": "Interface Language / Dil",
        "accounting_standard": "Accounting Standard / System",
        "std_us": "US / International (QuickBooks, Xero, NetSuite)",
        "std_tr": "Turkey (Uniform Chart of Accounts / Tek Düzen)",
        "custom_plan_header": "📁 Custom Chart of Accounts (Optional)",
        "custom_plan_help": "Upload your company's Chart of Accounts (.xlsx/.csv) with Account Code & Name.",
        "custom_plan_loaded": "✅ Loaded {count} accounts from your chart!",
        "rules_header": "📐 Default Account Rules",
        "vendor_rule_label": "Accounts Payable / Vendor Code Format",
        "rule_auto_tax": "Tax ID / EIN Based (e.g. 2000-EIN / 320.VKN)",
        "rule_auto_name": "Vendor Name Based (e.g. VEND-AMAZON / 320.SHELL)",
        "rule_standard": "Standard Sequential (e.g. 2000-01 / 320.01.001)",
        "rule_custom": "✏️ Custom Template...",
        "custom_format_help": "Use {NAME} for vendor name, {ID} for tax ID. Example: AP-{NAME} or 320-{ID}",
        "custom_input_label": "Enter Custom Account Template:",
    },
    "TR": {
        "title": "💼 LedgerAI - Akıllı Muhasebe Fiş Motoru",
        "subtitle": "Faturaları ve fişleri yükleyin; yapay zeka harcamaları otomatik sınıflandırsın, Borç/Alacak dengesini kursun ve resmi aktarım tablosu üretsin.",
        "uploader_label": "Fatura veya Fiş Yükleyin (PDF, PNG, JPG)",
        "limit_warning": "🛑 Ücretsiz demo sürümünde aynı anda en fazla 2 fatura işleyebilirsiniz. Kurumsal entegrasyon için iletişime geçin.",
        "files_count": "İşlenecek belge sayısı: **{count}**",
        "btn_process": "🚀 Akıllı Muhasebe Fişini Oluştur",
        "processing": "İşleniyor ({current}/{total}): {filename}...",
        "retrying": "Yoğunluk sebebiyle yeniden deneniyor ({wait} sn)...",
        "completed": "Tüm fişler başarıyla hazırlandı!",
        "preview_header": "📊 Oluşturulan Muhasebe Yevmiye Fişi (Düzenlenebilir)",
        "preview_tip": "💡 Tablodaki herhangi bir hücreye çift tıklayarak kod veya açıklamaları manuel değiştirebilirsiniz. Değişiklikler doğrudan Excel'e aktarılır.",
        "total_debit": "Toplam Borç",
        "total_credit": "Toplam Alacak",
        "balanced": "✅ Fiş Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "btn_download": "📥 Kurumsal Muhasebe Excel'ini İndir (.xlsx)",
        "sidebar_config": "⚙️ Muhasebe & Bölge Ayarları",
        "lang_select": "Dil Seçimi / Language",
        "accounting_standard": "Muhasebe Standardı / Program Türü",
        "std_us": "Uluslararası (QuickBooks, Xero, NetSuite, SAP)",
        "std_tr": "Türkiye (Tek Düzen - ETA, Luca, Zirve, Logo)",
        "custom_plan_header": "📁 Özel Hesap Planı Yükleme (Opsiyonel)",
        "custom_plan_help": "Programınızdan dışa aktardığınız Hesap Kodu ve Hesap Adı içeren Excel/CSV dosyasını yükleyin.",
        "custom_plan_loaded": "✅ {count} adet hesap başarıyla yüklendi!",
        "rules_header": "📐 Varsayılan Kodlama Kuralları",
        "vendor_rule_label": "Cari / Satıcı Kodlama Formatı",
        "rule_auto_tax": "Vergi No / TCKN Bazlı (Örn: 320.VKN)",
        "rule_auto_name": "Firma Adı Bazlı (Örn: 320.SHELL)",
        "rule_standard": "Standart Sıralı (Örn: 320.01.001)",
        "rule_custom": "✏️ Kendi Şablonumu Yazacağım...",
        "custom_format_help": "Firma adı için {AD}, vergi no için {VKN} yazın. Örnek: CARI-{AD} veya 320-{VKN}",
        "custom_input_label": "Özel Kod Şablonunuzu Yazın:",
    }
}

# --- YAN PANEL YAPILANDIRMASI ---
with st.sidebar:
    dil_secimi = st.selectbox("🌐 Language / Dil", ["English", "Türkçe"])
    lang = "EN" if dil_secimi == "English" else "TR"
    T = TEXTS[lang]
    
    st.title(T["sidebar_config"])
    muhasebe_sistemi = st.selectbox(
        T["accounting_standard"],
        [T["std_us"], T["std_tr"]]
    )
    is_global = (muhasebe_sistemi == T["std_us"])
    
    st.markdown("---")
    st.markdown(f"#### {T['custom_plan_header']}")
    hesap_plani_dosyasi = st.file_uploader(
        "Upload Accounts (.xlsx/.csv)", 
        type=["xlsx", "xls", "csv"],
        help=T["custom_plan_help"]
    )
    
    firma_hesap_ozeti = ""
    if hesap_plani_dosyasi:
        try:
            if hesap_plani_dosyasi.name.endswith(".csv"):
                df_plan = pd.read_csv(hesap_plani_dosyasi)
            else:
                df_plan = pd.read_excel(hesap_plani_dosyasi)
            
            cols = df_plan.columns[:2]
            ornek_kodlar = df_plan[cols].dropna().head(120).to_dict(orient="records")
            firma_hesap_ozeti = json.dumps(ornek_kodlar, ensure_ascii=False)
            st.success(T["custom_plan_loaded"].format(count=len(df_plan)))
        except Exception as e:
            st.error(f"Error: {str(e)[:50]}")
            
    st.markdown("---")
    st.markdown(f"#### {T['rules_header']}")
    
    secilen_kural = st.selectbox(
        T["vendor_rule_label"],
        [T["rule_auto_tax"], T["rule_auto_name"], T["rule_standard"], T["rule_custom"]]
    )
    
    ozel_sablon = ""
    if secilen_kural == T["rule_custom"]:
        ozel_sablon = st.text_input(
            T["custom_input_label"], 
            value="VEND-{NAME}" if is_global else "320.{AD}",
            help=T["custom_format_help"]
        )
        st.caption(f"💡 {T['custom_format_help']}")

# --- EXCEL OLUŞTURUCU (KURUMSAL VE EVRENSEL) ---
def excel_tablosu_olustur(df, sheet_name="Journal_Voucher"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]

        baslik_dolgu = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        baslik_yazi = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        govde_yazi = Font(name="Calibri", size=10)
        
        ince_kenarlik = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for col_idx in range(1, len(df.columns) + 1):
            hucre = worksheet.cell(row=1, column=col_idx)
            hucre.fill = baslik_dolgu
            hucre.font = baslik_yazi
            hucre.alignment = Alignment(horizontal="center", vertical="center")

        for col in worksheet.columns:
            maksimum_uzunluk = 0
            sutun_harfi = get_column_letter(col[0].column)
            
            for hucre in col:
                hucre.border = ince_kenarlik
                if hucre.row != 1:
                    hucre.font = govde_yazi
                    hucre.alignment = Alignment(vertical="center")
                
                val_str = str(hucre.value or '')
                if len(val_str) > maksimum_uzunluk:
                    maksimum_uzunluk = len(val_str)
            
            worksheet.column_dimensions[sutun_harfi].width = max(maksimum_uzunluk + 5, 14)

    return output.getvalue()

# --- ANA EKRAN ---
st.title(T["title"])
st.markdown(T["subtitle"])

yuklenen_dosyalar = st.file_uploader(
    T["uploader_label"], 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    if len(yuklenen_dosyalar) > 2:
        st.error(T["limit_warning"])
    else:
        st.info(T["files_count"].format(count=len(yuklenen_dosyalar)))
        
        if st.button(T["btn_process"], type="primary"):
            ham_veriler = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            toplam_dosya = len(yuklenen_dosyalar)
            
            for index, dosya in enumerate(yuklenen_dosyalar):
                status_text.text(T["processing"].format(current=index+1, total=toplam_dosya, filename=dosya.name))
                dosya_baytlari = dosya.read()
                mime_tipi = dosya.type if dosya.type else "application/pdf"
                
                # Hedef muhasebe sistemine göre özel Prompt kurgusu
                if is_global:
                    muhasebe_talimati = """
                    Classify expenses into US GAAP / Standard International Chart of Accounts:
                    - Office Supplies: 6100 Office Expense
                    - Fuel/Gas/Vehicle: 6200 Travel & Vehicle
                    - Meals & Entertainment: 6300 Meals & Ent.
                    - Inventory / Products: 1200 Inventory
                    - Software & SaaS: 6400 Software Subscription
                    - General Expense: 6000 Operating Expense
                    - Accounts Payable: 2000 Accounts Payable
                    Tax Account: 2200 Sales Tax / VAT
                    Currency: Extract the currency from the document (USD, EUR, GBP, TRY etc.)
                    """
                else:
                    muhasebe_talimati = """
                    Türkiye Tek Düzen Hesap Planı'na göre sınıflandır:
                    - Akaryakıt: 770.01
                    - Yemek & Ağırlama: 770.02
                    - Kırtasiye / Ofis: 770.03
                    - Ticari Mal: 153.01
                    - Kargo: 770.04
                    - Demirbaş: 255.01
                    - Genel: 770.99
                    KDV Hesabı: 191.01, 191.10 veya 191.20
                    Cari Hesabı: 320 grubu
                    """

                if firma_hesap_ozeti:
                    muhasebe_talimati += f"\nCRITICAL: Match against this company's custom chart of accounts: {firma_hesap_ozeti}"

                prompt = f"""
                You are a senior accountant and ERP data extraction specialist.
                {muhasebe_talimati}

                Extract and return ONLY a valid JSON object:
                {{
                  "invoice_no": "...",
                  "date": "YYYY-MM-DD",
                  "vendor": "...",
                  "tax_id": "...",
                  "currency": "USD",
                  "expense_category": "...",
                  "account_code": "...",
                  "account_name": "...",
                  "net_amount": 0.0,
                  "tax_rate": 0,
                  "tax_amount": 0.0,
                  "total_amount": 0.0
                }}
                Numeric values must be float. No markdown blocks, return pure JSON.
                """
                
                maksimum_deneme = 3
                for deneme in range(maksimum_deneme):
                    try:
                        yanit = client.models.generate_content(
                            model="gemini-3.5-flash-lite",
                            contents=[
                                types.Part.from_bytes(data=dosya_baytlari, mime_type=mime_tipi),
                                prompt
                            ]
                        )
                        
                        temiz_metin = yanit.text.replace("```json", "").replace("```", "").strip()
                        veri = json.loads(temiz_metin)
                        veri["filename"] = dosya.name
                        ham_veriler.append(veri)
                        break
                    except Exception as e:
                        hata_metni = str(e)
                        if ("503" in hata_metni or "429" in hata_metni) and deneme < maksimum_deneme - 1:
                            bekleme = 3 * (deneme + 1)
                            time.sleep(bekleme)
                            continue
                        else:
                            st.error(f"{dosya.name}: {hata_metni[:120]}")
                            break
                
                progress_bar.progress((index + 1) / toplam_dosya)
                time.sleep(1)
            
            status_text.text(T["completed"])
            
            if ham_veriler:
                fis_satirlari = []
                fis_sira_no = 1
                
                for item in ham_veriler:
                    fatura_no = str(item.get("invoice_no") or "").strip()
                    tarih = str(item.get("date") or "").strip()
                    satici = str(item.get("vendor") or "Bilinmeyen Satıcı").strip()
                    vkn = str(item.get("tax_id") or "").strip()
                    gider_kodu = str(item.get("account_code") or ("6000" if is_global else "770.01")).strip()
                    hesap_adi = str(item.get("account_name") or ("Operating Expense" if is_global else "Genel Gider")).strip()
                    para_birimi = str(item.get("currency") or ("USD" if is_global else "TL")).strip()
                    
                    try:
                        net = float(item.get("net_amount") or 0.0)
                    except (ValueError, TypeError):
                        net = 0.0

                    try:
                        tax = float(item.get("tax_amount") or 0.0)
                    except (ValueError, TypeError):
                        tax = 0.0

                    try:
                        total = float(item.get("total_amount") or (net + tax))
                    except (ValueError, TypeError):
                        total = net + tax

                    tax_rate = item.get("tax_rate") or 0

                    # Güvenli Temiz İsim Çıkarımı
                    temiz_ad = "".join(c for c in satici[:12] if c.isalnum()).upper()
                    if not temiz_ad:
                        temiz_ad = "VENDOR"

                    # Cari Kod Belirleme Mantığı
                    if secilen_kural == T["rule_custom"] and ozel_sablon:
                        cari_kod = ozel_sablon.replace("{NAME}", temiz_ad).replace("{AD}", temiz_ad).replace("{ID}", vkn).replace("{VKN}", vkn)
                    elif secilen_kural == T["rule_auto_tax"] and vkn:
                        cari_kod = f"2000-{vkn}" if is_global else f"320.{vkn}"
                    elif secilen_kural == T["rule_auto_name"] and satici:
                        cari_kod = f"VEND-{temiz_ad}" if is_global else f"320.{temiz_ad}"
                    else:
                        cari_kod = "2000-01" if is_global else "320.01.001"

                    # Başlık Etiketleri
                    col_voucher = "Voucher #" if is_global else "Fiş No"
                    col_date = "Date" if is_global else "Tarih"
                    col_code = "Account Code" if is_global else "Hesap Kodu"
                    col_name = "Account Description" if is_global else "Hesap Adı"
                    col_desc = "Memo / Line Description" if is_global else "Açıklama"
                    col_curr = "Currency" if is_global else "Para Birimi"
                    col_debit = "Debit" if is_global else "Borç"
                    col_credit = "Credit" if is_global else "Alacak"

                    # 1. Gider Satırı (Debit / Borç)
                    fis_satirlari.append({
                        col_voucher: fis_sira_no,
                        col_date: tarih,
                        col_code: gider_kodu,
                        col_name: hesap_adi,
                        col_desc: f"{satici} - Inv: {fatura_no}",
                        col_curr: para_birimi,
                        col_debit: net,
                        col_credit: 0.0
                    })
                    
                    # 2. Vergi / KDV Satırı (Debit / Borç)
                    if tax > 0:
                        tax_code = f"2200-TAX{tax_rate}" if is_global else f"191.{int(tax_rate):02d}"
                        fis_satirlari.append({
                            col_voucher: fis_sira_no,
                            col_date: tarih,
                            col_code: tax_code,
                            col_name: f"Tax ({tax_rate}%)" if is_global else f"%{tax_rate} İndirilecek KDV",
                            col_desc: f"{satici} - Tax",
                            col_curr: para_birimi,
                            col_debit: tax,
                            col_credit: 0.0
                        })
                    
                    # 3. Satıcı / AP Satırı (Credit / Alacak)
                    fis_satirlari.append({
                        col_voucher: fis_sira_no,
                        col_date: tarih,
                        col_code: cari_kod,
                        col_name: satici,
                        col_desc: f"{satici} - Inv: {fatura_no}",
                        col_curr: para_birimi,
                        col_debit: 0.0,
                        col_credit: total
                    })
                    
                    fis_sira_no += 1
                
                df_sonuc = pd.DataFrame(fis_satirlari)
                st.session_state["sonuc_tablosu"] = df_sonuc
                st.session_state["aktif_para_birimi"] = para_birimi
                st.session_state["is_global"] = is_global

if "sonuc_tablosu" in st.session_state:
    st.divider()
    st.subheader(T["preview_header"])
    st.info(T["preview_tip"])
    
    guncel_df = st.data_editor(
        st.session_state["sonuc_tablosu"], 
        use_container_width=True, 
        num_rows="dynamic"
    )
    
    col_debit = "Debit" if st.session_state["is_global"] else "Borç"
    col_credit = "Credit" if st.session_state["is_global"] else "Alacak"
    curr_symbol = "$" if st.session_state["is_global"] else "TL"

    toplam_debit = guncel_df[col_debit].sum()
    toplam_credit = guncel_df[col_credit].sum()
    
    col1, col2, col3 = st.columns(3)
    col1.metric(T["total_debit"], f"{curr_symbol} {toplam_debit:,.2f}")
    col2.metric(T["total_credit"], f"{curr_symbol} {toplam_credit:,.2f}")
    if abs(toplam_debit - toplam_credit) < 0.05:
        col3.success(T["balanced"])
    else:
        col3.warning(T["unbalanced"])

    excel_cikti = excel_tablosu_olustur(guncel_df, sheet_name="Journal_Import")
    st.download_button(
        label=T["btn_download"],
        data=excel_cikti,
        file_name="accounting_journal_voucher.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
