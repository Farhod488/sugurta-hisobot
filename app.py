import io
import pandas as pd
import streamlit as st

# Sahifa sozlamalari
st.set_page_config(
    page_title="Sug'urta Hisoboti Portali",
    page_icon="📑",
    layout="wide"
)

# Kesh funksiyasi
@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes):
    return pd.read_excel(io.BytesIO(file_bytes))

# Session state xotirasi
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "report_title" not in st.session_state:
    st.session_state.report_title = ""

# Ikkala tizimni vkladkalar (Tabs) orqali birlashtiramiz
tab1, tab2 = st.tabs(["📊 Kuzatuv va OSAGO (Excel hisobotlar)", "📄 Tushum hisoboti (DOCX / Solishtirish)"])

# =========================================================================
# 1-BO'LIM: SIZNING KUZATUV VA OSAGO HISOBOTLARINGIZ
# =========================================================================
with tab1:
    st.markdown("### 📑 Kuzatuv (Kafolat) va OSAGO (Majburiy) hisobotlari")
    st.write("Excel faylni yuklang, sug'urta turini hamda kunlik/oylik rejimni tanlang:")

    uploaded_file = st.file_uploader(
        "Excel faylni yuklang (.xlsx yoki .xls):",
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

            if st.button("Tayyor (Hisobotni shakllantirish)", type="primary", use_container_width=True):
                with st.spinner("Hisobot tayyorlanmoqda..."):
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

            # Natijani ko'rsatish
            if st.session_state.result_df is not None:
                st.markdown("---")
                st.subheader(f"📊 Natija: {st.session_state.report_title}")

                out_df = st.session_state.result_df
                st.dataframe(out_df, use_container_width=True)

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

# =========================================================================
# 2-BO'LIM: TANISHINGIZ QO'SHGAN TUSHUM VA WEB_UI QISMI
# =========================================================================
with tab2:
    try:
        from web_ui import run as run_tanishingiz
        run_tanishingiz()
    except Exception as e:
        st.warning(f"Tushum tizimini yuklashda xatolik: {e}")
