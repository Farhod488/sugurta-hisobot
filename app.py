import io
import pandas as pd
import streamlit as st

# Sahifa sozlamalari
st.set_page_config(
    page_title="Sug'urta Hisoboti Portali",
    page_icon="🛡️",
    layout="wide"
)

# Faqat sizning bo'limingiz uchun zamonaviy Custom CSS dizayn
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Asosiy banner */
    .portal-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #2563eb 100%);
        border-radius: 16px;
        padding: 28px 32px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.25);
    }
    .portal-banner h2 {
        color: #ffffff !important;
        font-weight: 700;
        font-size: 26px;
        margin: 0 0 8px 0;
        letter-spacing: -0.5px;
    }
    .portal-banner p {
        color: #cbd5e1;
        font-size: 15px;
        margin: 0;
        line-height: 1.5;
    }

    /* Parametrlar bloklari */
    .param-box {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 16px;
    }

    /* Metrika kartochkalari */
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
        font-size: 24px;
        font-weight: 700;
    }

    /* "Tayyor" tugmasi */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 12px 28px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 18px rgba(37, 99, 235, 0.4) !important;
    }

    /* Yuklab olish (Download) tugmasi */
    .stDownloadButton button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 14px 30px !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.3) !important;
    }

    /* Tablarni bezash */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        border-bottom: 2px solid #e2e8f0;
        margin-bottom: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        font-weight: 600;
        border-radius: 8px 8px 0 0;
    }
</style>
""", unsafe_allow_html=True)

# Kesh funksiyasi
@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes):
    return pd.read_excel(io.BytesIO(file_bytes))

# Xotira
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "report_title" not in st.session_state:
    st.session_state.report_title = ""

# Bo'limlar (Vkladkalar)
tab1, tab2 = st.tabs(["📊 Kuzatuv va OSAGO (Excel hisobotlar)", "📄 Tushum hisoboti (DOCX / Solishtirish)"])

# =========================================================================
# 1-BO'LIM: SIZNING BO'LIMINGIZ (DIZAYN YANGILANDI)
# =========================================================================
with tab1:
    st.markdown("""
    <div class="portal-banner">
        <h2>🛡️ Sug'urta hisobotlarini shakllantirish</h2>
        <p>Birlamchi Excel fayllarni qayta ishlash, foizlarni taqsimlash hamda kunlik va oylik yakuniy jadvallarni eksport qilish moduli</p>
    </div>
    """, unsafe_allow_html=True)

    # Fayl yuklash
    st.markdown("##### 📁 1-qadam: Birlamchi Excel faylni tanlang")
    uploaded_file = st.file_uploader(
        "Faylni yuklash maydoni",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
        key="main_excel_uploader"
    )

    if uploaded_file is not None:
        try:
            with st.spinner("Fayl o'qilmoqda..."):
                df_raw = load_excel_data(uploaded_file.getvalue())

            st.success(f"✅ Fayl muvaffaqiyatli yuklandi! Jami qatorlar soni: **{len(df_raw):,}** ta")

            # Parametrlar paneli
            st.markdown("##### ⚙️ 2-qadam: Hisobot parametrlarini belgilang")
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

            # Shakllantirish tugmasi
            if st.button("🚀 Tayyor (Hisobotni shakllantirish)", type="primary", use_container_width=True):
                with st.spinner("Hisobot generatsiya qilinmoqda..."):
                    df = df_raw.copy()

                    # --- 1. KUZATUV (KAFOLAT) ---
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

                    # --- 2. OSAGO (MAJBURIIY) ---
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

            # Natija va statistika bloki
            if st.session_state.result_df is not None:
                st.markdown("---")
                out_df = st.session_state.result_df
                jami_100 = out_df.loc[out_df["sug'urta firma"] == "JAMI", "100%"].values[0]
                polices_count = len(out_df) - 1

                st.markdown(f"### 📋 {st.session_state.report_title}")
                
                # Kartochkalar ko'rinishida statistika
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

                # Excelga eksport qilish
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
# 2-BO'LIM: TANISHINGIZNING BO'LIMI (O'ZGARISHSIZ QOLADI)
# =========================================================================
with tab2:
    try:
        from web_ui import run as run_tanishingiz
        run_tanishingiz()
    except Exception as e:
        st.warning(f"Tushum tizimini yuklashda xatolik: {e}")
