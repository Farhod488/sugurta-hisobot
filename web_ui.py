"""Streamlit UI: Tushum DOCX reports + Excel сводка."""

from __future__ import annotations

import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from app.tushum import Inspector, TushumInputs, generate_tushum_report
from app.tushum.assignments import PERSONAL_INSPECTORS, default_selected_names
from app.tushum.parsers import parse_kuzatuv_xlsx, parse_osago_xlsx

ROOT = Path(__file__).resolve().parent
TEMPLATES_DIR = ROOT / "templates"


def _save_upload(uploaded, folder: Path, preferred_name: str) -> Path | None:
    if uploaded is None:
        return None
    suffix = Path(uploaded.name).suffix or Path(preferred_name).suffix
    path = folder / f"{Path(preferred_name).stem}{suffix}"
    path.write_bytes(uploaded.getvalue())
    return path


def _company_names_from_xlsx(path: Path | None, *, kind: str) -> list[str]:
    if path is None or not path.exists():
        return []
    data = parse_kuzatuv_xlsx(path) if kind == "kuzatuv" else parse_osago_xlsx(path)
    return [row.name for row in data.rows]


def render_tushum_tab() -> None:
    st.subheader("Отчёты Тушум (DOCX)")
    st.caption(
        "PDF/Excel необязательны по отдельности — нужен хотя бы один файл. "
        "Без файла соответствующий абзац останется из шаблона."
    )

    period = st.text_input("Период", value="24-31-avgust", help="Например: 24-31-avgust")

    c1, c2 = st.columns(2)
    with c1:
        kuz_pdf = st.file_uploader("KUZATUV PDF (абз. 4)", type=["pdf"], key="kuz_pdf")
        kuz_xlsx = st.file_uploader(
            "KUZATUV Excel (абз. 5)", type=["xlsx"], key="kuz_xlsx"
        )
    with c2:
        osa_pdf = st.file_uploader("OSAGO PDF (абз. 8)", type=["pdf"], key="osa_pdf")
        osa_xlsx = st.file_uploader(
            "OSAGO Excel (абз. 9)", type=["xlsx"], key="osa_xlsx"
        )

    inspector_labels = {item.label_ru: item for item in Inspector}
    selected_label = st.radio(
        "Тип отчёта",
        list(inspector_labels.keys()),
        horizontal=True,
    )
    selected_inspector = inspector_labels[selected_label]

    kuz_assign: dict[Inspector, list[str]] = {}
    osa_assign: dict[Inspector, list[str]] = {}

    with tempfile.TemporaryDirectory(prefix="tushum_preview_") as tmp:
        tmp_path = Path(tmp)
        kuz_xlsx_path = _save_upload(kuz_xlsx, tmp_path, "KUZATUV.xlsx")
        osa_xlsx_path = _save_upload(osa_xlsx, tmp_path, "OSAGO.xlsx")

        try:
            kuz_names = _company_names_from_xlsx(kuz_xlsx_path, kind="kuzatuv")
            osa_names = _company_names_from_xlsx(osa_xlsx_path, kind="osago")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Не удалось прочитать Excel: {exc}")
            return

        if kuz_names or osa_names:
            st.markdown("#### Назначение компаний")
            st.caption(
                "Хамракулов всегда берёт итог по всем компаниям Excel. "
                "Ниже — только Арипов и Абдимуродов."
            )
            for item in PERSONAL_INSPECTORS:
                st.markdown(f"**{item.label_ru}**")
                cols = st.columns(2)
                with cols[0]:
                    if kuz_names:
                        defaults = default_selected_names(
                            kuz_names, item, source="kuzatuv"
                        )
                        kuz_assign[item] = st.multiselect(
                            "KUZATUV",
                            kuz_names,
                            default=defaults,
                            key=f"kuz_{item.value}",
                        )
                    else:
                        st.caption("KUZATUV Excel не загружен")
                        kuz_assign[item] = []
                with cols[1]:
                    if osa_names:
                        defaults = default_selected_names(
                            osa_names, item, source="osago"
                        )
                        osa_assign[item] = st.multiselect(
                            "OSAGO",
                            osa_names,
                            default=defaults,
                            key=f"osa_{item.value}",
                        )
                    else:
                        st.caption("OSAGO Excel не загружен")
                        osa_assign[item] = []

    gen_one, gen_all = st.columns(2)
    do_one = gen_one.button("Сформировать отчёт", type="primary", use_container_width=True)
    do_all = gen_all.button("Сформировать все 3", use_container_width=True)

    if not (do_one or do_all):
        return

    inspectors = list(Inspector) if do_all else [selected_inspector]
    if not any([kuz_pdf, kuz_xlsx, osa_pdf, osa_xlsx]):
        st.warning("Загрузите хотя бы один файл.")
        return

    for item in inspectors:
        if item not in PERSONAL_INSPECTORS:
            continue
        if kuz_xlsx and not kuz_assign.get(item):
            st.warning(f"{item.label_ru}: выберите хотя бы одну компанию KUZATUV.")
            return
        if osa_xlsx and not osa_assign.get(item):
            st.warning(f"{item.label_ru}: выберите хотя бы одну компанию OSAGO.")
            return

    if not TEMPLATES_DIR.exists():
        st.error(f"Не найдена папка шаблонов: {TEMPLATES_DIR}")
        return

    with st.spinner("Формирование отчёта…"):
        try:
            with tempfile.TemporaryDirectory(prefix="tushum_out_") as tmp:
                folder = Path(tmp)
                out_dir = folder / "output"
                out_dir.mkdir()
                inputs = TushumInputs(
                    kuzatuv_pdf=_save_upload(kuz_pdf, folder, "KUZATUV.pdf"),
                    kuzatuv_xlsx=_save_upload(kuz_xlsx, folder, "KUZATUV.xlsx"),
                    osago_pdf=_save_upload(osa_pdf, folder, "OSAGO.pdf"),
                    osago_xlsx=_save_upload(osa_xlsx, folder, "OSAGO.xlsx"),
                    period_text=period.strip() or "24-31-avgust",
                    templates_dir=TEMPLATES_DIR,
                    output_dir=out_dir,
                    kuzatuv_assignments=kuz_assign if kuz_xlsx else None,
                    osago_assignments=osa_assign if osa_xlsx else None,
                )
                results: list[tuple[str, bytes]] = []
                for item in inspectors:
                    path = generate_tushum_report(item, inputs)
                    results.append((path.name, path.read_bytes()))
        except Exception as exc:  # noqa: BLE001
            st.error(f"Ошибка: {exc}")
            return

    st.success(f"Готово: {len(results)} файл(ов)")
    for name, data in results:
        st.download_button(
            label=f"Скачать {name}",
            data=data,
            file_name=name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            key=f"dl_{name}",
        )


def render_excel_tab() -> None:
    st.subheader("Excel-сводка (75% / 25%)")
    st.write("Birlamchi Excel faylni yuklang va natijaviy hisobotni oling.")

    uploaded_file = st.file_uploader(
        "Excel faylni tanlang (.xlsx yoki .xls)", type=["xlsx", "xls"], key="excel_sum"
    )
    if uploaded_file is None:
        return

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

        if not st.button("Hisobotni shakllantirish", type="primary"):
            return

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

        jami_qator = pd.DataFrame(
            [
                {
                    col_firma: "Общий итог",
                    "polis_soni": hisobot["polis_soni"].sum(),
                    "res_ins_prem": hisobot["res_ins_prem"].sum(),
                    "summa_75": hisobot["summa_75"].sum(),
                    "summa_25": hisobot["summa_25"].sum(),
                    "mln_100": hisobot["mln_100"].sum(),
                    "mln_75": hisobot["mln_75"].sum(),
                    "mln_25": hisobot["mln_25"].sum(),
                }
            ]
        )
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

        st.download_button(
            label="Tayyor hisobotni yuklab olish (.xlsx)",
            data=output.getvalue(),
            file_name="sugurta_hisobot.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
    except Exception as exc:  # noqa: BLE001
        st.error(f"Xatolik yuz berdi: {exc}")


def run() -> None:
    st.set_page_config(page_title="Sug'urta hisoboti", layout="wide")
    st.title("Sug'urta hisobotini shakllantirish dasturi")

    tab_tushum, tab_excel = st.tabs(["Тушум (DOCX)", "Excel-сводка"])
    with tab_tushum:
        render_tushum_tab()
    with tab_excel:
        render_excel_tab()
