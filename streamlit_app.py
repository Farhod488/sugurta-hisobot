import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Sug'urta hisoboti", layout="wide")
st.title("Sug'urta hisobotini shakllantirish dasturi")
st.write("Birlamchi Excel faylni yuklang va natijaviy hisobotni oling.")

uploaded_file = st.file_uploader(
    "Excel faylni tanlang (.xlsx yoki .xls)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
  try:
    df = pd.read_excel(uploaded_file)
    st.success("Fayl muvaffaqiyatli yuklandi!")

    with st.expander("Yuklangan faylning dastlabki 5 ta qatori"):
      st.dataframe(df.head())

    columns = list(df.columns)

    default_firma = next(
        (
            c
            for c in columns
            if any(
                k in str(c).lower()
                for k in ["firma", "company", "наименование", "страховщик"]
            )
        ),
        columns[0],
    )
    default_polis = next(
        (
            c
            for c in columns
            if any(
                k in str(c).lower()
                for k in ["polis", "полис", "декларация", "nd", "номер"]
            )
        ),
        columns[1] if len(columns) > 1 else columns[0],
    )
    default_summa = next(
        (
            c
            for c in columns
            if any(
                k in str(c).lower()
                for k in ["prem", "res_ins_prem", "summa", "сумма"]
            )
        ),
        columns[2] if len(columns) > 2 else columns[0],
    )

    st.subheader("Ustunlarni moslashtirish")
    col1, col2, col3 = st.columns(3)
    with col1:
      col_firma = st.selectbox(
          "Firma nomi ustuni:",
          columns,
          index=columns.index(default_firma),
      )
    with col2:
      col_polis = st.selectbox(
          "Polis / Deklaratsiya ustuni:",
          columns,
          index=columns.index(default_polis),
      )
    with col3:
      col_summa = st.selectbox(
          "Umumiy summa (res_ins_prem) ustuni:",
          columns,
          index=columns.index(default_summa),
      )

    if st.button("Hisobotni shakllantirish", type="primary"):
      df[col_summa] = (
          pd.to_numeric(
              df[col_summa]
              .astype(str)
              .str.replace(" ", "")
              .str.replace(",", "."),
              errors="coerce",
          )
          .fillna(0)
      )
      df[col_firma] = df[col_firma].fillna("Noma'lum").astype(str).str.strip()

      hisobot = (
          df.groupby(col_firma)
          .agg(
              polis_soni=(col_polis, "count"),
              res_ins_prem=(col_summa, "sum"),
          )
          .reset_index()
      )

      hisobot["summa_75"] = hisobot["res_ins_prem"] * 0.75
      hisobot["summa_25"] = hisobot["res_ins_prem"] * 0.25

      hisobot["mln_100"] = hisobot["res_ins_prem"] / 1_000_000
      hisobot["mln_75"] = hisobot["summa_75"] / 1_000_000
      hisobot["mln_25"] = hisobot["summa_25"] / 1_000_000

      jami_qator = pd.DataFrame([{
          col_firma: "Общий итог",
          "polis_soni": hisobot["polis_soni"].sum(),
          "res_ins_prem": hisobot["res_ins_prem"].sum(),
          "summa_75": hisobot["summa_75"].sum(),
          "summa_25": hisobot["summa_25"].sum(),
          "mln_100": hisobot["mln_100"].sum(),
          "mln_75": hisobot["mln_75"].sum(),
          "mln_25": hisobot["mln_25"].sum(),
      }])

      yakuniy_df = pd.concat([hisobot, jami_qator], ignore_index=True)

      yakuniy_df = yakuniy_df.rename(
          columns={
              col_firma: "Названия строк",
              "polis_soni": "Количество по полю polis",
              "res_ins_prem": "Сумма по полю res_ins_prem",
              "summa_75": "Сумма по полю 75%",
              "summa_25": "Сумма по полю 25%",
              "mln_100": "100% (млн)",
              "mln_75": "75% (млн)",
              "mln_25": "25% (млн)",
          }
      )

      st.subheader("Natija")
      st.dataframe(yakuniy_df, use_container_width=True)

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        yakuniy_df.to_excel(writer, index=False, sheet_name="Hisobot")
      excel_data = output.getvalue()

      st.download_button(
          label="Tayyor hisobotni yuklab olish (.xlsx)",
          data=excel_data,
          file_name="sugurta_hisobot.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

  except Exception as e:
    st.error(f"Xatolik yuz berdi: {e}")
