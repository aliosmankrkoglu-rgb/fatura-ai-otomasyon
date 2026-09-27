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
    page_title="LedgerAI", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- MODERN, MİNİMALİST VE FERAH TEMA ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    .stApp {
        background-color: #090D16;
        color: #E2E8F0;
    }
    .metric-box {
        background: #111827;
        border: 1px solid #1F2937;
        border-radius: 12px;
        padding: 16px 20px;
    }
    div[data-testid="stFileUploader"] {
        background: #0F172A;
        border: 1px dashed #334155;
        border-radius: 14px;
        padding: 20px;
    }
</style>
""", unsafe_allow_html=True)

# API İstemcisi
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# --- YAN PANEL: ESNEK HESAP GRUPLARI ---
with st.sidebar:
    st.markdown("### ⚙️ Hesap Ayarları")
    
    with st.expander("📁 Özel Hesap Planı Yükle", expanded=False):
        hesap_plani = st.file_uploader("Excel veya CSV", type=["xlsx", "xls", "csv"], label_visibility="collapsed")
    
    hesap_ozeti = ""
    if hesap_plani:
        try:
            df_p = pd.read_csv(hesap_plani) if hesap_plani.name.endswith(".csv") else pd.read_excel(hesap_plani)
            c = df_p.columns[:2]
            hesap_ozeti = json.dumps(df_p[c].dropna().head(100).to_dict(orient="records"), ensure_ascii=False)
            st.success(f"✓ {len(df_p)} hesap tanımlandı")
        except Exception:
            pass

    st.markdown("---")
    st.markdown("##### Varsayılan Ana Hesaplar")
    custom_gider = st.text_input("Gider Hesabı Kökü", value="Otomatik (AI)", help="Örn: 770, 740 veya 6000")
    custom_kdv = st.text_input("KDV / Vergi Kökü", value="Otomatik (AI)", help="Örn: 191 veya 2200")
    custom_cari = st.text_input("Cari / AP Formatı", value="{KOK}.{AD}", help="{KOK} ana grup, {AD} firma adı, {VKN} vergi no")

# --- EXCEL OLUŞTURUCU ---
def excel_olustur(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Journal')
        ws = writer.sheets['Journal']

        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        body_font = Font(name="Calibri", size=10)
        border = Border(
            left=Side(style='thin', color='E2E8F0'),
            right=Side(style='thin', color='E2E8F0'),
            top=Side(style='thin', color='E2E8F0'),
            bottom=Side(style='thin', color='E2E8F0')
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

# --- ANA EKRAN ---
st.title("⚡ LedgerAI")
st.caption("Otomatik Belge & Dil Tespiti • Akıllı Gider Sınıflandırma • Dengeli Yevmiye Fişi")

yuklenen_dosyalar = st.file_uploader(
    "Fatura veya fişleri yükleyin (PDF, PNG, JPG)", 
    type=["pdf", "png", "jpg", "jpeg"], 
    accept_multiple_files=True,
    label_visibility="collapsed"
)

if yuklenen_dosyalar:
    if len(yuklenen_dosyalar) > 5:
        st.error("🛑 Demo sürümünde aynı anda en fazla 5 belge yükleyebilirsiniz.")
    else:
        st.write(f"📁 **{len(yuklenen_dosyalar)}** belge seçildi.")
        
        if st.button("⚡ Muhasebe Fişlerini Oluştur", type="primary", use_container_width=True):
            ham_veriler = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            toplam_dosya = len(yuklenen_dosyalar)
            
            for index, dosya in enumerate(yuklenen_dosyalar):
                status_text.text(f"İşleniyor ({index + 1}/{toplam_dosya}): {dosya.name}...")
                dosya_baytlari = dosya.read()
                mime_tipi = dosya.type if dosya.type else "application/pdf"
                
                # Belgenin dilini, ülkesini ve standartlarını otonom algılayan Prompt
                prompt = f"""
                You are a global autonomous accountant.
                1. Detect document country/language automatically:
                   - If Turkish/TL -> TR Tek Düzen (Expense: 770/153, Tax: 191, AP: 320)
                   - If German/DE/EUR -> Datev SKR03/04 (Expense: 4900, Tax: 1576, AP: 70000)
                   - If French/FR -> PCG (Expense: 606/618, Tax: 44566, AP: 401)
                   - If English/US/Global -> US GAAP (Expense: 6000 series, Tax: 2200, AP: 2000)
                2. Output column headers and text in the DETECTED DOCUMENT'S primary business language.
                3. Overrides from user if not 'Otomatik (AI)': Expense root: '{custom_gider}', Tax root: '{custom_kdv}'.
                {'Match against chart: ' + hesap_ozeti if hesap_ozeti else ''}

                Return ONLY a JSON object:
                {{
                  "doc_lang": "TR/EN/DE/FR",
                  "currency": "USD/EUR/TL",
                  "invoice_no": "...",
                  "date": "YYYY-MM-DD",
                  "vendor": "...",
                  "tax_id": "...",
                  "expense_code": "...",
                  "expense_name": "...",
                  "net": 0.0,
                  "tax_rate": 0,
                  "tax": 0.0,
                  "total": 0.0
                }}
                Numeric fields must be float. No markdown.
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
                        veri["filename"] = dosya.name
                        ham_veriler.append(veri)
                        break
                    except Exception as e:
                        hata_metni = str(e)
                        if ("503" in hata_metni or "429" in hata_metni) and deneme < maksimum_deneme - 1:
                            time.sleep(3 * (deneme + 1))
                            continue
                        else:
                            st.warning(f"{dosya.name} okunamadı: {hata_metni[:80]}")
                            break
                
                progress_bar.progress((index + 1) / toplam_dosya)
            
            # Durum mesajı yönetimi
            if len(ham_veriler) == toplam_dosya:
                status_text.success("✓ Tüm fişler başarıyla oluşturuldu.")
            elif len(ham_veriler) > 0:
                status_text.warning(f"✓ {len(ham_veriler)} belge işlendi, {toplam_dosya - len(ham_veriler)} belge okunamadı.")
            else:
                status_text.error("Belgeler işlenemedi. Lütfen dosya netliğini kontrol edin.")

            if ham_veriler:
                fis_satirlari = []
                v_no = 1
                
                for item in ham_veriler:
                    d_lang = item.get("doc_lang", "EN")
                    curr = item.get("currency", "USD")
                    inv_no = str(item.get("invoice_no") or "").strip()
                    date_val = str(item.get("date") or "").strip()
                    vendor = str(item.get("vendor") or "Vendor").strip()
                    tax_id = str(item.get("tax_id") or "").strip()
                    exp_code = str(item.get("expense_code") or "6000").strip()
                    exp_name = str(item.get("expense_name") or "Expense").strip()
                    
                    net = float(item.get("net") or 0.0)
                    tax = float(item.get("tax") or 0.0)
                    total = float(item.get("total") or (net + tax))
                    tax_rate = item.get("tax_rate") or 0
                    
                    # Cari hesap kökü tespiti
                    ap_root = "320" if d_lang == "TR" else ("70000" if d_lang == "DE" else ("401" if d_lang == "FR" else "2000"))
                    clean_name = "".join(c for c in vendor[:10] if c.isalnum()).upper() or "VEND"
                    
                    # Cari kod şablonu çözümü
                    ap_code = custom_cari.replace("{KOK}", ap_root).replace("{AD}", clean_name).replace("{VKN}", tax_id or clean_name)
                    
                    # Başlıkları belgenin diline göre dinamik belirleme
                    if d_lang == "TR":
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Fiş No", "Tarih", "Hesap Kodu", "Hesap Adı", "Açıklama", "Borç", "Alacak"
                        tax_name = f"%{tax_rate} İndirilecek KDV"
                        tax_code = f"191.{int(tax_rate):02d}"
                    elif d_lang == "DE":
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Beleg", "Datum", "Konto", "Bezeichnung", "Buchungstext", "Soll", "Haben"
                        tax_name = f"Vorsteuer {tax_rate}%"
                        tax_code = "1576"
                    elif d_lang == "FR":
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Pièce", "Date", "Compte", "Libellé Compte", "Libellé Écriture", "Débit", "Crédit"
                        tax_name = f"TVA Déductible {tax_rate}%"
                        tax_code = "44566"
                    else:
                        h_v, h_d, h_c, h_n, h_m, h_deb, h_crd = "Voucher #", "Date", "Account Code", "Account Name", "Description", "Debit", "Credit"
                        tax_name = f"Tax ({tax_rate}%)"
                        tax_code = f"2200-{tax_rate}"

                    # 1. Gider Satırı
                    fis_satirlari.append({
                        h_v: v_no, h_d: date_val, h_c: exp_code, h_n: exp_name,
                        h_m: f"{vendor} - {inv_no}", "Currency": curr, h_deb: net, h_crd: 0.0
                    })
                    
                    # 2. Vergi Satırı
                    if tax > 0:
                        fis_satirlari.append({
                            h_v: v_no, h_d: date_val, h_c: tax_code, h_n: tax_name,
                            h_m: f"{vendor} - Tax", "Currency": curr, h_deb: tax, h_crd: 0.0
                        })
                    
                    # 3. Cari / Satıcı Satırı
                    fis_satirlari.append({
                        h_v: v_no, h_d: date_val, h_c: ap_code, h_n: vendor,
                        h_m: f"{vendor} - {inv_no}", "Currency": curr, h_deb: 0.0, h_crd: total
                    })
                    
                    v_no += 1

                st.session_state["out_df"] = pd.DataFrame(fis_satirlari)
                st.session_state["deb_col"] = h_deb
                st.session_state["crd_col"] = h_crd

if "out_df" in st.session_state:
    st.divider()
    df_preview = st.session_state["out_df"]
    deb_col = st.session_state["deb_col"]
    crd_col = st.session_state["crd_col"]
    
    # Doğrudan düzenlenebilir modern veri tablosu
    guncel_df = st.data_editor(df_preview, use_container_width=True, num_rows="dynamic")
    
    t_deb = guncel_df[deb_col].sum()
    t_crd = guncel_df[crd_col].sum()
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Toplam Borç / Debit", f"{t_deb:,.2f}")
    col2.metric("Toplam Alacak / Credit", f"{t_crd:,.2f}")
    if abs(t_deb - t_crd) < 0.05:
        col3.success("✅ Fiş Dengeli (Balanced)")
    else:
        col3.warning("⚠️ Bakiye Farkı Var!")
        
    excel_dosya = excel_olustur(guncel_df)
    st.download_button(
        label="📥 Muhasebe Excel Dosyasını İndir (.xlsx)",
        data=excel_dosya,
        file_name="ledger_journal_voucher.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
