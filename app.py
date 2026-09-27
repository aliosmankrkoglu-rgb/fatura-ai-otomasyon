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
    page_title="LedgerAI - Autonomous Accounting & ERP Engine", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- MODERN VE GÖZ YORMAYAN ÖZEL CSS TEMASI ---
st.markdown("""
<style>
    .main { background-color: #0F172A; color: #F8FAFC; }
    .stMetric { background: rgba(30, 41, 59, 0.7); border: 1px solid #334155; border-radius: 10px; padding: 12px 16px; }
    .badge-info { background-color: #1E293B; border-left: 4px solid #3B82F6; padding: 12px; border-radius: 6px; margin-bottom: 16px; }
    .rule-card { background-color: #1E293B; border: 1px solid #334155; border-radius: 8px; padding: 14px; margin-top: 10px; font-size: 0.88rem; }
</style>
""", unsafe_allow_html=True)

# API İstemcisi
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# --- ÇOK DİLLİ VE ÇOK BÖLGELİ VERİ TABANI ---
LANG_DICT = {
    "English 🇺🇸": {
        "title": "⚡ LedgerAI — Autonomous Accounting Engine",
        "subtitle": "Transform raw invoices and receipts into structured, balanced multi-currency journal vouchers ready for ERP import.",
        "uploader_label": "Drag and drop receipts or invoices (PDF, PNG, JPG)",
        "limit_warning": "🛑 Free demo allows up to 5 documents per batch. Contact sales for unlimited volume.",
        "files_count": "Ready to process: **{count}** files",
        "btn_process": "⚡ Process & Generate Journal Vouchers",
        "processing": "Extracting & categorizing ({current}/{total}): {filename}...",
        "completed": "✅ All documents processed successfully!",
        "partial_success": "⚠️ Completed with warnings: {success} succeeded, {failed} failed.",
        "all_failed": "❌ None of the uploaded documents could be parsed. Please check file clarity.",
        "preview_header": "📊 Generated Accounting Journal Voucher (Editable)",
        "preview_tip": "💡 You can double-click any cell to adjust codes or values before downloading.",
        "total_debit": "Total Debit",
        "total_credit": "Total Credit",
        "balanced": "✅ Voucher Balanced (Debit = Credit)",
        "unbalanced": "⚠️ Unbalanced Voucher!",
        "btn_download": "📥 Download Clean ERP Excel (.xlsx)",
        "sidebar_title": "⚙️ Configuration Center",
        "lang_select": "Interface Language",
        "standard_select": "Target Accounting Standard",
        "std_us": "Global / US GAAP (QuickBooks, Xero, NetSuite)",
        "std_de": "Germany / DACH (SKR03 / SKR04 Datev)",
        "std_fr": "France (Plan Comptable Général - PCG)",
        "std_tr": "Turkey (Uniform Chart / Tek Düzen - ETA, Luca)",
        "guide_step1": "1️⃣ Custom Chart of Accounts (Optional)",
        "guide_step1_desc": "Upload your ERP's account list (.xlsx or .csv) so AI maps expenses to your exact internal codes.",
        "guide_step2": "2️⃣ Accounts Payable (AP) Rule",
        "rule_tax": "Tax ID / VAT Based (e.g., 2000-EIN)",
        "rule_name": "Vendor Name Based (e.g., VEND-AMAZON)",
        "rule_std": "Sequential Number (e.g., 2000-01)",
        "rule_custom": "Custom Pattern / Şablon...",
        "custom_pattern_label": "Enter Custom Format:",
        "custom_pattern_help": "Keywords: {NAME} = Vendor Name, {ID} = Tax ID/VAT. Example: AP-{NAME} or 320-{ID}",
        "step_badge": "ℹ️ <b>Quick Setup:</b> Invoices are verified for debit/credit balance, currency, and multi-tier tax classification."
    },
    "Türkçe 🇹🇷": {
        "title": "⚡ LedgerAI — Otonom Muhasebe Fiş Motoru",
        "subtitle": "Faturaları ve fişleri yükleyin; sistem harcamaları sınıflandırsın, Borç/Alacak dengesini kursun ve ERP aktarım dosyasını çıkarsın.",
        "uploader_label": "Fatura veya fişleri sürükleyip bırakın (PDF, PNG, JPG)",
        "limit_warning": "🛑 Ücretsiz demo sürümünde aynı anda en fazla 5 fatura işleyebilirsiniz.",
        "files_count": "İşlenecek dosya sayısı: **{count}**",
        "btn_process": "⚡ Akıllı Muhasebe Fişini Oluştur",
        "processing": "Ayrıştırılıyor ({current}/{total}): {filename}...",
        "completed": "✅ Tüm belgeler başarıyla işlendi ve yevmiye fişine dönüştürüldü!",
        "partial_success": "⚠️ İşlem tamamlandı: {success} başarılı, {failed} başarısız.",
        "all_failed": "❌ Belgelerin hiçbiri okunamadı. Lütfen dosya kalitesini kontrol edin.",
        "preview_header": "📊 Oluşturulan Muhasebe Yevmiye Fişi (Düzenlenebilir)",
        "preview_tip": "💡 Tablodaki herhangi bir hücreye çift tıklayarak kod veya tutarları değiştirebilirsiniz.",
        "total_debit": "Toplam Borç",
        "total_credit": "Toplam Alacak",
        "balanced": "✅ Fiş Bakiyesi Dengeli (Borç = Alacak)",
        "unbalanced": "⚠️ Bakiye Farkı Var!",
        "btn_download": "📥 Muhasebe Aktarım Excel'ini İndir (.xlsx)",
        "sidebar_title": "⚙️ Muhasebe & Kural Merkezi",
        "lang_select": "Arayüz Dili",
        "standard_select": "Muhasebe Standardı / Program",
        "std_us": "Global / US GAAP (QuickBooks, Xero, NetSuite)",
        "std_de": "Almanya / DACH (SKR03 / SKR04 Datev)",
        "std_fr": "Fransa (Plan Comptable Général - PCG)",
        "std_tr": "Türkiye (Tek Düzen - ETA, Luca, Zirve, Logo)",
        "guide_step1": "1️⃣ Özel Hesap Planı Yükleme (Opsiyonel)",
        "guide_step1_desc": "Muhasebe programınızdaki hesap listesini (Excel/CSV) yükleyin; yapay zeka harcamaları birebir kendi kodlarınıza eşlesin.",
        "guide_step2": "2️⃣ Cari / Satıcı (320) Kodlama Kuralı",
        "rule_tax": "Vergi No / TCKN Bazlı (Örn: 320.VKN)",
        "rule_name": "Firma Adı Bazlı (Örn: 320.SHELL)",
        "rule_std": "Standart Sıralı (Örn: 320.01.001)",
        "rule_custom": "Özel Kural Şablonu...",
        "custom_pattern_label": "Özel Şablonunuzu Yazın:",
        "custom_pattern_help": "Anahtar kelimeler: {AD} = Firma Adı, {VKN} = Vergi No. Örnek: CARI-{AD} veya 320-{VKN}",
        "step_badge": "ℹ️ <b>Hızlı Bilgi:</b> Belgeler çoklu KDV oranına, harcama kalemine ve kuruşu kuruşuna borç-alacak eşitliğine göre denetlenir."
    },
    "Deutsch 🇩🇪": {
        "title": "⚡ LedgerAI — Autonome Buchhaltungs-Engine",
        "subtitle": "Eingangsrechnungen und Belege automatisch erfassen, vorkontieren und Datev/ERP-konforme Buchungssätze generieren.",
        "uploader_label": "Rechnungen oder Belege hochladen (PDF, PNG, JPG)",
        "limit_warning": "🛑 Demo-Limit: Maximal 5 Dokumente pro Durchgang.",
        "files_count": "Bereit zur Verarbeitung: **{count}** Dokumente",
        "btn_process": "⚡ Buchungssätze automatisch erstellen",
        "processing": "Wird verarbeitet ({current}/{total}): {filename}...",
        "completed": "✅ Buchungssätze erfolgreich erstellt!",
        "partial_success": "⚠️ Teilweise abgeschlossen: {success} erfolgreich, {failed} fehlgeschlagen.",
        "all_failed": "❌ Keine Dokumente konnten analysiert werden.",
        "preview_header": "📊 Erstellte Buchungszeilen (Bearbeitbar)",
        "preview_tip": "💡 Doppelklicken Sie auf ein Feld, um Konten oder Beträge direkt zu editieren.",
        "total_debit": "Soll Gesamt",
        "total_credit": "Haben Gesamt",
        "balanced": "✅ Ausgeglichen (Soll = Haben)",
        "unbalanced": "⚠️ Differenz festgestellt!",
        "btn_download": "📥 Datev/Excel Buchungsdatei herunterladen (.xlsx)",
        "sidebar_title": "⚙️ Kontierungs-Einstellungen",
        "lang_select": "Sprache",
        "standard_select": "Buchhaltungsstandard",
        "std_us": "Global / US GAAP (QuickBooks, Xero)",
        "std_de": "Deutschland (Datev SKR03 / SKR04)",
        "std_fr": "Frankreich (PCG Standard)",
        "std_tr": "Türkei (Tek Düzen)",
        "guide_step1": "1️⃣ Kontenrahmen hochladen (Optional)",
        "guide_step1_desc": "Laden Sie Ihren SKR03/04 Kontenplan hoch für automatische Kontenzuordnung.",
        "guide_step2": "2️⃣ Kreditoren-Nummernlogik",
        "rule_tax": "Steuernummer / USt-IdNr",
        "rule_name": "Lieferantenname (z.B. KRED-SHELL)",
        "rule_std": "Standard Kreditor (z.B. 70000)",
        "rule_custom": "Benutzerdefiniertes Muster...",
        "custom_pattern_label": "Muster eingeben:",
        "custom_pattern_help": "Verwenden Sie {NAME} für Name, {ID} für Steuernummer. Beispiel: KRED-{NAME}",
        "step_badge": "ℹ️ <b>Status:</b> Automatische Vorsteueraufteilung und Soll/Haben-Validierung aktiv."
    },
    "Français 🇫🇷": {
        "title": "⚡ LedgerAI — Moteur Comptable Autonome",
        "subtitle": "Numérisez et ventilez vos factures fournisseurs en écritures comptables conformes pour votre logiciel ERP.",
        "uploader_label": "Déposer des factures ou reçus (PDF, PNG, JPG)",
        "limit_warning": "🛑 Démo limitée à 5 documents par session.",
        "files_count": "Prêt à traiter: **{count}** documents",
        "btn_process": "⚡ Générer les Écritures Comptables",
        "processing": "Analyse en cours ({current}/{total}): {filename}...",
        "completed": "✅ Écritures comptables générées avec succès!",
        "partial_success": "⚠️ Traitement partiel: {success} réussies, {failed} échouées.",
        "all_failed": "❌ Aucun document n'a pu être traité.",
        "preview_header": "📊 Journal Comptable Généré (Modifiable)",
        "preview_tip": "💡 Double-cliquez sur une cellule pour ajuster les comptes ou montants.",
        "total_debit": "Total Débit",
        "total_credit": "Total Crédit",
        "balanced": "✅ Équilibré (Débit = Crédit)",
        "unbalanced": "⚠️ Déséquilibre Détecté!",
        "btn_download": "📥 Télécharger le Journal Excel (.xlsx)",
        "sidebar_title": "⚙️ Configuration Comptable",
        "lang_select": "Langue",
        "standard_select": "Norme Comptable",
        "std_us": "International / US GAAP",
        "std_de": "Allemagne (Datev SKR03/04)",
        "std_fr": "France (Plan Comptable Général - PCG)",
        "std_tr": "Turquie (Tek Düzen)",
        "guide_step1": "1️⃣ Plan Comptable Personnalisé (Optionnel)",
        "guide_step1_desc": "Importez votre plan de comptes pour une imputation automatique.",
        "guide_step2": "2️⃣ Règle Compte Fournisseur (401)",
        "rule_tax": "Numéro SIREN / TVA Intra",
        "rule_name": "Nom du Fournisseur (ex: 401SHELL)",
        "rule_std": "Standard Séquentiel (ex: 401000)",
        "rule_custom": "Modèle Personnalisé...",
        "custom_pattern_label": "Format personnalisé:",
        "custom_pattern_help": "Utilisez {NAME} pour le nom, {ID} pour le numéro. Exemple: 401-{NAME}",
        "step_badge": "ℹ️ <b>Contrôle:</b> Ventilation TVA et équilibre Débit/Crédit vérifiés en continu."
    }
}

# --- YAN PANEL YAPILANDIRMASI ---
with st.sidebar:
    st.markdown("### 🌐 Localization")
    dil = st.selectbox("", list(LANG_DICT.keys()), label_visibility="collapsed")
    L = LANG_DICT[dil]
    
    st.title(L["sidebar_title"])
    st.markdown(f"<div class='badge-info'>{L['step_badge']}</div>", unsafe_allow_html=True)
    
    muhasebe_secimi = st.selectbox(
        L["standard_select"],
        [L["std_us"], L["std_de"], L["std_fr"], L["std_tr"]]
    )
    
    st.markdown(f"**{L['guide_step1']}**")
    st.caption(L["guide_step1_desc"])
    hesap_plani = st.file_uploader(
        "Upload Plan (.xlsx/.csv)", 
        type=["xlsx", "xls", "csv"], 
        label_visibility="collapsed"
    )
    
    hesap_ozeti = ""
    if hesap_plani:
        try:
            if hesap_plani.name.endswith(".csv"):
                df_p = pd.read_csv(hesap_plani)
            else:
                df_p = pd.read_excel(hesap_plani)
            c = df_p.columns[:2]
            hesap_ozeti = json.dumps(df_p[c].dropna().head(100).to_dict(orient="records"), ensure_ascii=False)
            st.success(f"✓ {len(df_p)} accounts loaded")
        except Exception as e:
            st.error(f"Error reading file: {str(e)[:40]}")

    st.markdown(f"**{L['guide_step2']}**")
    kural = st.selectbox(
        "AP Rule",
        [L["rule_tax"], L["rule_name"], L["rule_std"], L["rule_custom"]],
        label_visibility="collapsed"
    )
    
    ozel_kod_sablonu = ""
    if kural == L["rule_custom"]:
        ozel_kod_sablonu = st.text_input(
            L["custom_pattern_label"],
            value="AP-{NAME}" if L["std_us"] in muhasebe_secimi else "320.{AD}"
        )
        st.caption(L["custom_pattern_help"])

# --- FORMATLI EXCEL OLUŞTURUCU FONKSİYON ---
def excel_olustur(df, sheet_name="Journal"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
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
            m_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                cell.border = border
                if cell.row != 1:
                    cell.font = body_font
                    cell.alignment = Alignment(vertical="center")
                val_len = len(str(cell.value or ''))
                if val_len > m_len:
                    m_len = val_len
            ws.column_dimensions[col_letter].width = max(m_len + 5, 14)

    return output.getvalue()

# --- ANA EKRAN ---
st.title(L["title"])
st.markdown(L["subtitle"])

yuklenen_dosyalar = st.file_uploader(
    L["uploader_label"], 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    # 5 DOSYA LİMİT KONTROLÜ
    if len(yuklenen_dosyalar) > 5:
        st.error(L["limit_warning"])
    else:
        st.info(L["files_count"].format(count=len(yuklenen_dosyalar)))
        
        if st.button(L["btn_process"], type="primary"):
            ham_veriler = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            toplam_dosya = len(yuklenen_dosyalar)
            
            # Seçilen muhasebe standardına göre prompt direktifi
            if muhasebe_secimi == L["std_de"]:
                standard_prompt = "Format according to German SKR03/SKR04 Datev standards (e.g. 4900/6800 for general expenses, 1576 for Vorsteuer 19%, 70000 for Kreditoren)."
            elif muhasebe_secimi == L["std_fr"]:
                standard_prompt = "Format according to French PCG standards (e.g. 606/618 for charges, 44566 for TVA déductible, 401 for Fournisseurs)."
            elif muhasebe_secimi == L["std_tr"]:
                standard_prompt = "Türkiye Tek Düzen Hesap Planına göre kodla: Akaryakıt 770.01, Yemek 770.02, Kırtasiye 770.03, Mal 153.01, KDV 191 grubu, Satıcı 320 grubu."
            else:
                standard_prompt = "Format according to US GAAP / International standards: 6100 Office, 6200 Travel/Fuel, 6300 Meals, 1200 Inventory, 2200 Sales Tax, 2000 Accounts Payable."

            if hesap_ozeti:
                standard_prompt += f"\nSTRICT: Map against customer's chart of accounts: {hesap_ozeti}"

            for index, dosya in enumerate(yuklenen_dosyalar):
                status_text.text(L["processing"].format(current=index+1, total=toplam_dosya, filename=dosya.name))
                dosya_baytlari = dosya.read()
                mime_tipi = dosya.type if dosya.type else "application/pdf"
                
                prompt = f"""
                You are a senior automated accounting auditor.
                {standard_prompt}

                Examine this document and extract ONLY a valid JSON object:
                {{
                  "invoice_no": "...",
                  "date": "YYYY-MM-DD",
                  "vendor": "...",
                  "tax_id": "...",
                  "currency": "USD",
                  "account_code": "...",
                  "account_name": "...",
                  "net_amount": 0.0,
                  "tax_rate": 0,
                  "tax_amount": 0.0,
                  "total_amount": 0.0
                }}
                No markdown, output pure JSON only.
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
                        temiz = yanit.text.replace("```json", "").replace("```", "").strip()
                        veri = json.loads(temiz)
                        veri["filename"] = dosya.name
                        ham_veriler.append(veri)
                        break
                    except Exception as e:
                        hata_msg = str(e)
                        if ("503" in hata_msg or "429" in hata_msg) and deneme < maksimum_deneme - 1:
                            time.sleep(3 * (deneme + 1))
                            continue
                        else:
                            st.warning(f"⚠️ {dosya.name}: {hata_msg[:90]}")
                            break
                
                progress_bar.progress((index + 1) / toplam_dosya)
            
            # --- MANTIKSAL MESAJ YÖNETİMİ (HATA ALINDIĞINDA BAŞARILI YAZMAZ) ---
            if len(ham_veriler) == toplam_dosya:
                status_text.text(L["completed"])
            elif len(ham_veriler) > 0:
                status_text.text(L["partial_success"].format(success=len(ham_veriler), failed=toplam_dosya - len(ham_veriler)))
            else:
                status_text.text(L["all_failed"])

            if ham_veriler:
                fis_satirlari = []
                voucher_num = 1
                
                is_tr = (muhasebe_secimi == L["std_tr"])
                
                for item in ham_veriler:
                    inv_no = str(item.get("invoice_no") or "").strip()
                    date_val = str(item.get("date") or "").strip()
                    vendor = str(item.get("vendor") or "Vendor").strip()
                    tax_id = str(item.get("tax_id") or "").strip()
                    acc_code = str(item.get("account_code") or ("770.01" if is_tr else "6000")).strip()
                    acc_name = str(item.get("account_name") or ("Expense" if not is_tr else "Gider")).strip()
                    curr = str(item.get("currency") or ("TL" if is_tr else "USD")).strip()
                    
                    try:
                        net = float(item.get("net_amount") or 0.0)
                    except:
                        net = 0.0
                    try:
                        tax = float(item.get("tax_amount") or 0.0)
                    except:
                        tax = 0.0
                    try:
                        total = float(item.get("total_amount") or (net + tax))
                    except:
                        total = net + tax
                    
                    tax_rate = item.get("tax_rate") or 0
                    clean_name = "".join(c for c in vendor[:12] if c.isalnum()).upper() or "VENDOR"

                    # AP / Cari Hesap Kodlama
                    if kural == L["rule_custom"] and ozel_kod_sablonu:
                        ap_code = ozel_kod_sablonu.replace("{NAME}", clean_name).replace("{AD}", clean_name).replace("{ID}", tax_id).replace("{VKN}", tax_id)
                    elif kural == L["rule_tax"] and tax_id:
                        ap_code = f"320.{tax_id}" if is_tr else f"2000-{tax_id}"
                    elif kural == L["rule_name"]:
                        ap_code = f"320.{clean_name}" if is_tr else f"VEND-{clean_name}"
                    else:
                        ap_code = "320.01.001" if is_tr else "2000-01"

                    h_vouch = "Fiş No" if is_tr else "Voucher #"
                    h_date = "Tarih" if is_tr else "Date"
                    h_code = "Hesap Kodu" if is_tr else "Account Code"
                    h_name = "Hesap Adı" if is_tr else "Account Name"
                    h_memo = "Açıklama" if is_tr else "Memo"
                    h_curr = "Para Birimi" if is_tr else "Currency"
                    h_deb = "Borç" if is_tr else "Debit"
                    h_crd = "Alacak" if is_tr else "Credit"

                    # 1. Gider Satırı (Debit)
                    fis_satirlari.append({
                        h_vouch: voucher_num, h_date: date_val, h_code: acc_code,
                        h_name: acc_name, h_memo: f"{vendor} - {inv_no}",
                        h_curr: curr, h_deb: net, h_crd: 0.0
                    })
                    
                    # 2. Vergi / KDV Satırı (Debit)
                    if tax > 0:
                        t_code = f"191.{int(tax_rate):02d}" if is_tr else f"2200-TAX{tax_rate}"
                        t_name = f"%{tax_rate} İndirilecek KDV" if is_tr else f"Tax ({tax_rate}%)"
                        fis_satirlari.append({
                            h_vouch: voucher_num, h_date: date_val, h_code: t_code,
                            h_name: t_name, h_memo: f"{vendor} - Tax",
                            h_curr: curr, h_deb: tax, h_crd: 0.0
                        })
                    
                    # 3. Satıcı / AP Satırı (Credit)
                    fis_satirlari.append({
                        h_vouch: voucher_num, h_date: date_val, h_code: ap_code,
                        h_name: vendor, h_memo: f"{vendor} - {inv_no}",
                        h_curr: curr, h_deb: 0.0, h_crd: total
                    })
                    
                    voucher_num += 1

                df_out = pd.DataFrame(fis_satirlari)
                st.session_state["out_df"] = df_out
                st.session_state["is_tr"] = is_tr

if "out_df" in st.session_state:
    st.divider()
    st.subheader(L["preview_header"])
    st.info(L["preview_tip"])
    
    edited_df = st.data_editor(
        st.session_state["out_df"],
        use_container_width=True,
        num_rows="dynamic"
    )
    
    col_deb = "Borç" if st.session_state["is_tr"] else "Debit"
    col_crd = "Alacak" if st.session_state["is_tr"] else "Credit"
    sym = "TL" if st.session_state["is_tr"] else "$"
    
    tot_deb = edited_df[col_deb].sum()
    tot_crd = edited_df[col_crd].sum()
    
    c1, c2, c3 = st.columns(3)
    c1.metric(L["total_debit"], f"{sym} {tot_deb:,.2f}")
    c2.metric(L["total_credit"], f"{sym} {tot_crd:,.2f}")
    if abs(tot_deb - tot_crd) < 0.05:
        c3.success(L["balanced"])
    else:
        c3.warning(L["unbalanced"])
        
    excel_file = excel_olustur(edited_df)
    st.download_button(
        label=L["btn_download"],
        data=excel_file,
        file_name="ledger_journal_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
