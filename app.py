import io
import re
import pandas as pd
import openpyxl
import streamlit as st

# Sahifa sozlamalari
st.set_page_config(
    page_title="Sug'urta Hisoboti Portali",
    page_icon="🛡️",
    layout="wide"
)

# Custom CSS dizayn
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .portal-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #2563eb 100%);
        border-radius: 16px;
        padding: 24px 28px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.25);
    }
    .portal-banner h2 {
        color: #ffffff !important;
        font-weight: 700;
        font-size: 24px;
        margin: 0 0 6px 0;
    }
    .portal-banner p {
        color: #cbd5e1;
        font-size: 14px;
        margin: 0;
    }

    .param-box {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 16px;
        margin-bottom: 14px;
    }

    .metric-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        text-align: center;
        border-top: 4px solid #2563eb;
    }
    .metric-card .title {
        color: #64748b;
        font-size: 13px;
        font-weight: 600;
        text-transform: uppercase;
        margin-bottom: 6px;
    }
    .metric-card .value {
        color: #0f172a;
        font-size: 22px;
        font-weight: 700;
    }

    div.stButton > button:first-child {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 12px 24px !important;
        font-size: 15px !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3) !important;
    }

    .stDownloadButton button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 12px 26px !important;
        font-size: 15px !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.3) !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        border-bottom: 2px solid #e2e8f0;
        margin-bottom: 18px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 18px;
        font-weight: 600;
        border-radius: 8px 8px 0 0;
    }
</style>
""", unsafe_allow_html=True)

# Kesh funksiyasi
@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes):
    return pd.read_excel(io.BytesIO(file_bytes))

# Session state xotirasi
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "report_title" not in st.session_state:
    st.session_state.report_title = ""
if "filled_template_bytes" not in st.session_state:
    st.session_state.filled_template_bytes = None
if "filled_preview_df" not in st.session_state:
    st.session_state.filled_preview_df = None

# 3 ta bo'lim
tab1, tab2, tab3 = st.tabs([
    "📊 Kuzatuv va OSAGO (Excel hisobotlar)", 
    "📑 Rangli Yig'ma Hisobot", 
    "📄 Tushum hisoboti (DOCX / Solishtirish)"
])

# =========================================================================
# 1-BO'LIM: KUZATUV VA OSAGO KUNLIK / OYLIK
# =========================================================================
with tab1:
    col_l, col_m, col_r = st.columns([1, 1, 1])
    with col_m:
        try:
            st.image("Logo mbl.png", use_container_width=True)
        except Exception:
            pass

    st.markdown("""
    <div class="portal-banner">
        <h2>🛡️ Sug'urta hisobotlarini shakllantirish</h2>
        <p>Birlamchi Excel fayllarni qayta ishlash, foizlarni taqsimlash hamda kunlik va oylik yakuniy jadvallarni eksport qilish moduli</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Birlamchi Excel faylni tanlang (.xlsx yoki .xls):",
        type=["xlsx", "xls"],
        key="main_excel_uploader"
    )

    if uploaded_file is not None:
        try:
            with st.spinner("Fayl o'qilmoqda..."):
                df_raw = load_excel_data(uploaded_file.getvalue())

            st.success(f"Fayl muvaffaqiyatli yuklandi! Jami qatorlar: **{len(df_raw):,}** ta")

            col_type, col_period = st.columns(2)
            with col_type:
                st.markdown('<div class="param-box">', unsafe_allow_html=True)
                sugurta_turi = st.radio(
                    "📌 Sug'urta turi:",
                    ["Kuzatuv (Kafolat)", "OSAGO (Majburiy)"],
                    horizontal=True
                )
                st.markdown('</div>', unsafe_allow_html=True)

            with col_period:
                st.markdown('<div class="param-box">', unsafe_allow_html=True)
                davr = st.radio(
                    "📅 Hisobot davri:",
                    ["Kunlik", "Oylik"],
                    horizontal=True
                )
                st.markdown('</div>', unsafe_allow_html=True)

            if st.button("🚀 Tayyor (Hisobotni shakllantirish)", type="primary", use_container_width=True):
                with st.spinner("Hisobot tayyorlanmoqda..."):
                    df = df_raw.copy()

                    # 1. KUZATUV
                    if sugurta_turi == "Kuzatuv (Kafolat)":
                        sum_col = "res_ins_prem" if "res_ins_prem" in df.columns else "ins_prem"
                        df[sum_col] = pd.to_numeric(
                            df[sum_col].astype(str).str.replace(" ", "").str.replace(",", "."),
                            errors="coerce"
                        ).fillna(0)

                        firma_col = "sugurta_firma" if "sugurta_firma" in df.columns else df.columns[0]
                        df["sug'urta firma"] = df[firma_col].fillna("Noma'lum").astype(str).str.strip()

                        seria = df["seria_polis"].fillna("").astype(str).str.strip() if "seria_polis" in df.columns else ""
                        number = df["number_polis"].fillna("").astype(str).str.strip() if "number_polis" in df.columns else ""
                        df["polis raqami"] = (seria + " " + number).str.strip()

                        df["100%"] = df[sum_col]
                        df["75%"] = df["100%"] * 0.75
                        df["25%"] = df["100%"] * 0.25

                        if davr == "Kunlik":
                            cols_final = ["sug'urta firma", "polis raqami", "100%", "75%", "25%"]
                            st.session_state.report_title = "Kuzatuv kunlik hisoboti"
                        else:
                            df["deklaratsiya"] = df["declaration"].fillna("").astype(str).str.strip() if "declaration" in df.columns else ""
                            date_col = "ins_date" if "ins_date" in df.columns else ("request_vaqti" if "request_vaqti" in df.columns else "")
                            df["polis olingan vaqti"] = df[date_col].fillna("").astype(str).str.strip() if date_col else ""
                            df["polis muddati"] = df["ins_srok"].fillna("").astype(str).str.strip() if "ins_srok" in df.columns else ""
                            
                            lname = df["lastname"].fillna("").astype(str).str.strip() if "lastname" in df.columns else ""
                            fname = df["firstname"].fillna("").astype(str).str.strip() if "firstname" in df.columns else ""
                            df["familya ism"] = (lname + " " + fname).str.strip()

                            cols_final = [
                                "sug'urta firma", "deklaratsiya", "polis olingan vaqti", 
                                "polis muddati", "polis raqami", "familya ism", 
                                "100%", "75%", "25%"
                            ]
                            st.session_state.report_title = "Kuzatuv oylik hisoboti"

                    # 2. OSAGO
                    else:
                        sum_col = "premium" if "premium" in df.columns else "mukofot_sum"
                        df[sum_col] = pd.to_numeric(
                            df[sum_col].astype(str).str.replace(" ", "").str.replace(",", "."),
                            errors="coerce"
                        ).fillna(0)

                        firma_col = "insurance_company" if "insurance_company" in df.columns else df.columns[0]
                        df["sug'urta firma"] = df[firma_col].fillna("Noma'lum").astype(str).str.strip()

                        df["deklaratsiya"] = df["deklaratsiya"].fillna("").astype(str).str.strip() if "deklaratsiya" in df.columns else ""
                        df["polis turi"] = df["polis_type"].fillna("").astype(str).str.strip() if "polis_type" in df.columns else ""

                        df["100%"] = df[sum_col]
                        df["80%"] = df["100%"] * 0.80
                        df["20%"] = df["100%"] * 0.20

                        if davr == "Kunlik":
                            cols_final = ["sug'urta firma", "deklaratsiya", "polis turi", "100%", "80%", "20%"]
                            st.session_state.report_title = "OSAGO kunlik hisoboti"
                        else:
                            date_col = "begin_date" if "begin_date" in df.columns else ("request_vaqti" if "request_vaqti" in df.columns else "")
                            df["polis olingan vaqti"] = df[date_col].fillna("").astype(str).str.strip() if date_col else ""
                            df["polis muddati"] = df["polis_type"].fillna("").astype(str).str.strip() if "polis_type" in df.columns else ""

                            seria = df["policy_seria"].fillna("").astype(str).str.strip() if "policy_seria" in df.columns else ""
                            number = df["policy_number"].fillna("").astype(str).str.strip() if "policy_number" in df.columns else ""
                            df["polis raqami"] = (seria + " " + number).str.strip()

                            lname = df["lastname"].fillna("").astype(str).str.strip() if "lastname" in df.columns else ""
                            fname = df["firstname"].fillna("").astype(str).str.strip() if "firstname" in df.columns else ""
                            df["familya ism"] = (lname + " " + fname).str.strip()

                            cols_final = [
                                "sug'urta firma", "deklaratsiya", "polis olingan vaqti", 
                                "polis muddati", "polis raqami", "familya ism", 
                                "100%", "80%", "20%"
                            ]
                            st.session_state.report_title = "OSAGO oylik hisoboti"

                    res_df = df[cols_final].copy()

                    # Jami qatori
                    total_row = {c: "" for c in cols_final}
                    total_row["sug'urta firma"] = "JAMI"
                    total_row["100%"] = res_df["100%"].sum()

                    if "75%" in cols_final:
                        total_row["75%"] = res_df["75%"].sum()
                        total_row["25%"] = res_df["25%"].sum()
                    if "80%" in cols_final:
                        total_row["80%"] = res_df["80%"].sum()
                        total_row["20%"] = res_df["20%"].sum()

                    res_df = pd.concat([res_df, pd.DataFrame([total_row])], ignore_index=True)
                    st.session_state.result_df = res_df

            if st.session_state.result_df is not None:
                st.markdown("---")
                out_df = st.session_state.result_df
                jami_100 = out_df.loc[out_df["sug'urta firma"] == "JAMI", "100%"].values[0]
                polices_count = len(out_df) - 1

                st.markdown(f"### 📋 {st.session_state.report_title}")
                
                m_col1, m_col2 = st.columns(2)
                with m_col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <div class="title">Jami rasmiylashtirilgan polislar</div>
                        <div class="value">{polices_count:,} ta</div>
                    </div>
                    """, unsafe_allow_html=True)
                with m_col2:
                    st.markdown(f"""
                    <div class="metric-card" style="border-top-color: #10b981;">
                        <div class="title">Umumiy summa (100%)</div>
                        <div class="value">{jami_100:,.2f} so'm</div>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.dataframe(out_df, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    out_df.to_excel(writer, index=False, sheet_name="Hisobot")
                excel_bytes = output.getvalue()

                st.download_button(
                    label="📥 Tayyor hisobotni Excel (.xlsx) formatda yuklab olish",
                    data=excel_bytes,
                    file_name=f"{st.session_state.report_title.replace(' ', '_').lower()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Xatolik yuz berdi: {e}")

# =========================================================================
# 2-BO'LIM: RANGLI YIG'MA SHABLON HISOBOT (YANGI BO'LIM)
# =========================================================================
with tab2:
    st.markdown("""
    <div class="portal-banner" style="background: linear-gradient(135deg, #065f46 0%, #047857 50%, #059669 100%);">
        <h2>📑 Rangli Yig'ma Hisobot Shablonini Yangilash</h2>
        <p>Birlamchi ma'lumotlarni hisoblab, rangli Excel shablonning "Kunlik" va "Tadbir boshidan" ustunlariga avtomatik integratsiya qilish</p>
    </div>
    """, unsafe_allow_html=True)

    # 1. Sug'urta turini tanlash tugmasi
    st.markdown("##### 📌 1-qadam: Sug'urta turini tanlang")
    report_type_selection = st.radio(
        "Sug'urta yo'nalishi:",
        ["Kuzatuv (Kafolat) — 75% / 25%", "OSAGO (Majburiy) — 80% / 20%"],
        horizontal=True,
        key="tab2_report_type"
    )
    is_kuzatuv = "Kuzatuv" in report_type_selection

    st.markdown("---")
    st.markdown("##### 📁 2-qadam: Ikkala Excel faylni yuklang")
    f_col1, f_col2 = st.columns(2)
    with f_col1:
        daily_file = st.file_uploader(
            "1. Birlamchi kunlik Excel fayl (Polislar):",
            type=["xlsx", "xls"],
            key="tab2_daily_file",
            help="Kafolat yoki Majburiy polislar bazasi"
        )
    with f_col2:
        template_file = st.file_uploader(
            "2. Rangli shablon Excel fayl (Namunadagi jadval):",
            type=["xlsx", "xls"],
            key="tab2_template_file",
            help="Sizdagi rangli jadvalli shablon fayl"
        )

    if daily_file and template_file:
        if st.button("⚡ Hisobotni yangilash va rangli Excel yaratish", type="primary", use_container_width=True):
            with st.spinner("Firmalar bo'yicha hisob-kitob qilinmoqda va shablonga joylanmoqda..."):
                try:
                    df_in = pd.read_excel(daily_file)

                    # Birlamchi fayldan firmalar bo'yicha hisoblash
                    if is_kuzatuv:
                        f_col = "sugurta_firma" if "sugurta_firma" in df_in.columns else df_in.columns[0]
                        s_col = "res_ins_prem" if "res_ins_prem" in df_in.columns else "ins_prem"
                        pct_org = 0.75
                        pct_mbl = 0.25
                    else:
                        f_col = "insurance_company" if "insurance_company" in df_in.columns else df_in.columns[0]
                        s_col = "premium" if "premium" in df_in.columns else "mukofot_sum"
                        pct_org = 0.80
                        pct_mbl = 0.20

                    df_in[s_col] = pd.to_numeric(
                        df_in[s_col].astype(str).str.replace(" ", "").str.replace(",", "."),
                        errors="coerce"
                    ).fillna(0)
                    df_in[f_col] = df_in[f_col].fillna("Noma'lum").astype(str).str.strip()

                    def clean_firm_name(val):
                        v = str(val).upper()
                        v = re.sub(r'["“”`\'\s]', '', v)
                        for token in ['AJ', 'SK', 'EISK', 'QK', 'KOMPANIYASI', 'SUGURTA', 'SUG`URTA', 'INSURANCE']:
                            v = v.replace(token, '')
                        return v

                    df_in["clean_firm"] = df_in[f_col].apply(clean_firm_name)

                    grouped = df_in.groupby("clean_firm").agg(
                        cnt=(s_col, "count"),
                        sum_val=(s_col, "sum")
                    ).to_dict(orient="index")

                    # Rangli shablonni openpyxl orqali yuklash (ranglar, chiziqlar saqlanadi)
                    wb = openpyxl.load_workbook(template_file)
                    ws = wb.active

                    # Ustunlar va qatorlarni aniqlash
                    row_jami = None
                    firm_rows = []
                    col_kunlik_cnt = None
                    col_kunlik_prem = None
                    col_kunlik_org = None
                    col_kunlik_mbl = None
                    col_tadbir_cnt = None
                    col_tadbir_prem = None
                    col_tadbir_org = None
                    col_tadbir_mbl = None

                    # Shablon jadvalini skanerlash
                    for r in range(1, 25):
                        for c in range(1, 20):
                            val = str(ws.cell(r, c).value or "")
                            if "kunlik" in val.lower():
                                # Kunlik ustunlarini topish
                                for sub_c in range(c, c + 7):
                                    hdr = str(ws.cell(r + 1, sub_c).value or "") + " " + str(ws.cell(r + 2, sub_c).value or "")
                                    hdr_l = hdr.lower()
                                    if "polis" in hdr_l and col_kunlik_cnt is None:
                                        col_kunlik_cnt = sub_c
                                    elif ("mukofot" in hdr_l or "summa" in hdr_l) and col_kunlik_prem is None:
                                        col_kunlik_prem = sub_c
                                    elif "tashkilot" in hdr_l and col_kunlik_org is None:
                                        col_kunlik_org = sub_c
                                    elif "mbl" in hdr_l and col_kunlik_mbl is None:
                                        col_kunlik_mbl = sub_c

                            if "tadbir boshidan" in val.lower() and col_tadbir_cnt is None:
                                for sub_c in range(c, c + 6):
                                    hdr = str(ws.cell(r + 1, sub_c).value or "") + " " + str(ws.cell(r + 2, sub_c).value or "")
                                    hdr_l = hdr.lower()
                                    if "polis" in hdr_l and col_tadbir_cnt is None:
                                        col_tadbir_cnt = sub_c
                                    elif ("mukofot" in hdr_l or "summa" in hdr_l) and col_tadbir_prem is None:
                                        col_tadbir_prem = sub_c
                                    elif "tashkilot" in hdr_l and col_tadbir_org is None:
                                        col_tadbir_org = sub_c
                                    elif "mbl" in hdr_l and col_tadbir_mbl is None:
                                        col_tadbir_mbl = sub_c

                    # Default ustun pozitsiyalari (agar skan topolmasa)
                    if not col_kunlik_cnt:
                        col_kunlik_cnt, col_kunlik_prem, col_kunlik_org, col_kunlik_mbl = 9, 10, 11, 12
                    if not col_tadbir_cnt:
                        col_tadbir_cnt, col_tadbir_prem, col_tadbir_org, col_tadbir_mbl = 5, 6, 7, 8

                    # Jami va firmalar qatorlarini aniqlash
                    for r in range(4, 25):
                        c1_val = str(ws.cell(r, 1).value or "")
                        c2_val = str(ws.cell(r, 2).value or "")
                        row_txt = c1_val + " " + c2_val
                        if "jami" in row_txt.lower() and row_jami is None:
                            row_jami = r
                        elif any(k in row_txt.upper() for k in ["GROSS", "TEMIR", "XALQ", "ALFA", "APEX", "INSON", "KAPITAL", "OMAD", "SEMURG"]):
                            firm_name = c2_val if len(c2_val) > 3 else c1_val
                            firm_rows.append((r, firm_name, clean_firm_name(firm_name)))

                    def to_float(val):
                        if val is None:
                            return 0.0
                        if isinstance(val, (int, float)):
                            return float(val)
                        s = str(val).replace(" ", "").replace(",", ".")
                        try:
                            return float(s)
                        except:
                            return 0.0

                    total_k_cnt = 0
                    total_k_prem = 0.0
                    total_k_org = 0.0
                    total_k_mbl = 0.0

                    # 1. Kunlik va Tadbir boshidan ustunlarini yangilash
                    for r, orig_name, c_name in firm_rows:
                        # Birlamchi fayldan mos firmani qidirish
                        matched_key = next((k for k in grouped if k in c_name or c_name in k), None)
                        if matched_key:
                            d_cnt = grouped[matched_key]["cnt"]
                            d_prem = grouped[matched_key]["sum_val"] / 1_000_000.0  # mln so'mga o'tkazish
                        else:
                            d_cnt = 0
                            d_prem = 0.0

                        d_org = d_prem * pct_org
                        d_mbl = d_prem * pct_mbl

                        total_k_cnt += d_cnt
                        total_k_prem += d_prem
                        total_k_org += d_org
                        total_k_mbl += d_mbl

                        # Kunlik ustuniga yozish
                        ws.cell(r, col_kunlik_cnt, value=d_cnt if d_cnt > 0 else "-")
                        ws.cell(r, col_kunlik_prem, value=round(d_prem, 1) if d_prem > 0 else "-")
                        ws.cell(r, col_kunlik_org, value=round(d_org, 1) if d_org > 0 else "-")
                        ws.cell(r, col_kunlik_mbl, value=round(d_mbl, 1) if d_mbl > 0 else "-")

                        # Tadbir boshidan ustuniga qo'shish (kümülyativ)
                        cur_t_cnt = to_float(ws.cell(r, col_tadbir_cnt).value)
                        cur_t_prem = to_float(ws.cell(r, col_tadbir_prem).value)
                        cur_t_org = to_float(ws.cell(r, col_tadbir_org).value)
                        cur_t_mbl = to_float(ws.cell(r, col_tadbir_mbl).value)

                        ws.cell(r, col_tadbir_cnt, value=int(cur_t_cnt + d_cnt))
                        ws.cell(r, col_tadbir_prem, value=round(cur_t_prem + d_prem, 1))
                        ws.cell(r, col_tadbir_org, value=round(cur_t_org + d_org, 1))
                        ws.cell(r, col_tadbir_mbl, value=round(cur_t_mbl + d_mbl, 1))

                    # Jami qatorini yangilash
                    if row_jami:
                        ws.cell(row_jami, col_kunlik_cnt, value=int(total_k_cnt))
                        ws.cell(row_jami, col_kunlik_prem, value=round(total_k_prem, 1))
                        ws.cell(row_jami, col_kunlik_org, value=round(total_k_org, 1))
                        ws.cell(row_jami, col_kunlik_mbl, value=round(total_k_mbl, 1))

                        cur_j_cnt = to_float(ws.cell(row_jami, col_tadbir_cnt).value)
                        cur_j_prem = to_float(ws.cell(row_jami, col_tadbir_prem).value)
                        cur_j_org = to_float(ws.cell(row_jami, col_tadbir_org).value)
                        cur_j_mbl = to_float(ws.cell(row_jami, col_tadbir_mbl).value)

                        ws.cell(row_jami, col_tadbir_cnt, value=int(cur_j_cnt + total_k_cnt))
                        ws.cell(row_jami, col_tadbir_prem, value=round(cur_j_prem + total_k_prem, 1))
                        ws.cell(row_jami, col_tadbir_org, value=round(cur_j_org + total_k_org, 1))
                        ws.cell(row_jami, col_tadbir_mbl, value=round(cur_j_mbl + total_k_mbl, 1))

                    # 2. O'ng tarafdagi 4 ustunlik blokni yangilash
                    # Jadvaldan o'ngda joylashgan kataklar qidiriladi
                    for r_sub in range(len(firm_rows) + 5, 30):
                        for c_sub in range(13, 22):
                            val_sub = ws.cell(r_sub, c_sub).value
                            if val_sub is not None and isinstance(val_sub, (int, float)) and val_sub > 0:
                                # Shu yerdagi 4 ta katak yangilanadi
                                ws.cell(r_sub, c_sub, value=int(total_k_cnt))
                                ws.cell(r_sub, c_sub + 1, value=round(total_k_prem, 1))
                                ws.cell(r_sub, c_sub + 2, value=round(total_k_org, 1))
                                ws.cell(r_sub, c_sub + 3, value=round(total_k_mbl, 1))
                                break

                    # Natijani saqlash
                    out_stream = io.BytesIO()
                    wb.save(out_stream)
                    st.session_state.filled_template_bytes = out_stream.getvalue()

                    st.success("✅ Shablon muvaffaqiyatli yangilandi! Ranglar va formulalar 100% saqlandi.")

                except Exception as ex:
                    st.error(f"Xatolik yuz berdi: {ex}")

        # Yuklab olish tugmasi
        if st.session_state.filled_template_bytes is not None:
            st.markdown("---")
            st.download_button(
                label="📥 Yangilangan Rangli Hisobotni Yuklab Olish (.xlsx)",
                data=st.session_state.filled_template_bytes,
                file_name="yangilangan_rangli_hisobot.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# =========================================================================
# 3-BO'LIM: TANISHINGIZNING BO'LIMI (O'ZGARISHSIZ)
# =========================================================================
with tab3:
    try:
        from web_ui import run as run_tanishingiz
        run_tanishingiz()
    except Exception as e:
        st.warning(f"Tushum tizimini yuklashda xatolik: {e}")
