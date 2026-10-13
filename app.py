import io
import pandas as pd
import streamlit as st

# Sahifa sozlamalari
st.set_page_config(
    page_title="Sug'urta Hisobotlari Portali",
    page_icon="📑",
    layout="wide"
)

# Zamonaviy vizual dizayn
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 24px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.08);
    }
    .main-header h1 {
        color: white;
        margin: 0;
        font-size: 26px;
    }
    .main-header p {
        color: #e2e8f0;
        margin-top: 6px;
        font-size: 14px;
    }
    .stDownloadButton button {
        background-color: #10b981 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 10px 24px !important;
        font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📑 Sug'urta hisobotlarini shakllantirish portali</h1>
    <p>Kuzatuv (Kafolat) va OSAGO (Majburiy) bo'yicha kunlik hamda oylik hisobotlar tizimi</p>
</div>
""", unsafe_allow_html=True)

# Kesh funksiyasi (katta hajmdagi fayllar tez yuklanishi uchun)
@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes):
    return pd.read_excel(io.BytesIO(file_bytes))

# Natijani saqlash uchun session_state
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "report_title" not in st.session_state:
    st.session_state.report_title = ""

# 1-qadam: Fayl yuklash
uploaded_file = st.file_uploader(
    "1. Birlamchi Excel faylni yuklang (.xlsx yoki .xls):",
    type=["xlsx", "xls"],
    help="Hajmi 50 MB gacha bo'lgan fayllarni qo'llab-quvvatlaydi"
)

if uploaded_file is not None:
    try:
        with st.spinner("Fayl o'qilmoqda..."):
            df_raw = load_excel_data(uploaded_file.getvalue())

        st.success(f"Fayl muvaffaqiyatli yuklandi! Jami qatorlar: **{len(df_raw):,}** ta")

        st.markdown("---")
        st.markdown("### 2. Hisobot parametrlarini tanlang")
        
        col_type, col_period = st.columns(2)
        with col_type:
            sugurta_turi = st.radio(
                "Sug'urta turini tanlang:",
                ["Kuzatuv (Kafolat)", "OSAGO (Majburiy)"],
                horizontal=True
            )
        with col_period:
            davr = st.radio(
                "Hisobot davrini tanlang:",
                ["Kunlik", "Oylik"],
                horizontal=True
            )

        # "Tayyor" tugmasi
        if st.button("Tayyor (Hisobotni shakllantirish)", type="primary", use_container_width=True):
            with st.spinner("Hisobot tayyorlanmoqda..."):
                df = df_raw.copy()

                # --- 1. KUZATUV (KAFOLAT) ---
                if sugurta_turi == "Kuzatuv (Kafolat)":
                    # Summani tozalash va songa o'tkazish
                    sum_col = "res_ins_prem" if "res_ins_prem" in df.columns else "ins_prem"
                    df[sum_col] = pd.to_numeric(
                        df[sum_col].astype(str).str.replace(" ", "").str.replace(",", "."),
                        errors="coerce"
                    ).fillna(0)

                    # Firma nomi
                    firma_col = "sugurta_firma" if "sugurta_firma" in df.columns else df.columns[0]
                    df["sug'urta firma"] = df[firma_col].fillna("Noma'lum").astype(str).str.strip()

                    # Polis raqami: seria_polis + number_polis
                    seria = df["seria_polis"].fillna("").astype(str).str.strip() if "seria_polis" in df.columns else ""
                    number = df["number_polis"].fillna("").astype(str).str.strip() if "number_polis" in df.columns else ""
                    df["polis raqami"] = (seria + " " + number).str.strip()

                    # Foizlar
                    df["100%"] = df[sum_col]
                    df["75%"] = df["100%"] * 0.75
                    df["25%"] = df["100%"] * 0.25

                    if davr == "Kunlik":
                        cols_final = ["sug'urta firma", "polis raqami", "100%", "75%", "25%"]
                        st.session_state.report_title = "Kuzatuv kunlik hisoboti"
                    else: # Oylik
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
                    else: # Oylik
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

                # Tanlangan ustunlarni filtrlash
                res_df = df[cols_final].copy()

                # Jami (Umumiy natija) qatorini qo'shish
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

        # Natijani ko'rsatish
        if st.session_state.result_df is not None:
            st.markdown("---")
            st.subheader(f"📊 Natija: {st.session_state.report_title}")

            out_df = st.session_state.result_df
            jami_val = out_df.loc[out_df["sug'urta firma"] == "JAMI", "100%"].values[0]

            m1, m2 = st.columns(2)
            m1.metric("Jami yozuvlar soni", f"{len(out_df) - 1:,} ta")
            m2.metric("Umumiy 100% summa", f"{jami_val:,.2f} so'm")

            st.dataframe(out_df, use_container_width=True)

            # Excelga saqlash
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                out_df.to_excel(writer, index=False, sheet_name="Hisobot")
            excel_bytes = output.getvalue()

            st.download_button(
                label="📥 Hisobotni Excel fayl sifatida yuklab olish (.xlsx)",
                data=excel_bytes,
                file_name=f"{st.session_state.report_title.replace(' ', '_').lower()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Xatolik yuz berdi: {e}")
