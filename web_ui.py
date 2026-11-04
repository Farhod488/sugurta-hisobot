"""Streamlit UI: Tushum DOCX reports + Excel сводка (fixed)."""

from __future__ import annotations

import io
import logging
import re
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from app.tushum import Inspector, TushumInputs, generate_tushum_report
from app.tushum.assignments import PERSONAL_INSPECTORS, default_selected_names
from app.tushum.parsers import parse_kuzatuv_xlsx, parse_osago_xlsx

logger = logging.getLogger(__name__)

ROOT = Path(__file__).resolve().parent
TEMPLATES_DIR = ROOT / "templates"

DOCX_MIME = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DEFAULT_PERIOD = "24-31-avgust"
UNSAFE_PERIOD_RE = re.compile(r'[\\/:*?"<>|]|\.\.')
TOTAL_ROW_RE = re.compile(r"^\s*(общий\s+итог|итого|итог|jami|umumiy|total|grand\s+total)\b", re.I)


# --------------------------------------------------------------------------
# Yordamchi funksiyalar
# --------------------------------------------------------------------------
def _file_key(uploaded) -> str:
    """Yuklangan fayl uchun barqaror kalit (yangi fayl -> yangi kalit)."""
    return f"{uploaded.name}_{uploaded.size}"


def _save_upload(uploaded, folder: Path, preferred_name: str) -> Path | None:
    if uploaded is None:
        return None
    suffix = Path(uploaded.name).suffix or Path(preferred_name).suffix
    path = folder / f"{Path(preferred_name).stem}{suffix}"
    path.write_bytes(uploaded.getvalue())
    return path


@st.cache_data(show_spinner=False, max_entries=10)
def _get_company_names_cached(file_bytes: bytes, filename: str, kind: str) -> list[str]:
    with tempfile.NamedTemporaryFile(suffix=Path(filename).suffix, delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = Path(tmp.name)
    try:
        data = parse_kuzatuv_xlsx(tmp_path) if kind == "kuzatuv" else parse_osago_xlsx(tmp_path)
        return [row.name for row in data.rows]
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


@st.cache_data(show_spinner=False, max_entries=10)
def _read_excel_cached(file_bytes: bytes, filename: str, header_row: int) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), header=header_row)


def _guess_column(
    columns: list,
    pattern: str,
    fallback_index: int,
    exclude: set | None = None,
):
    """Nomi bo'yicha ustunni topadi; topilmasa fallback_index bo'yicha oladi.
    `exclude` dagi ustunlar qaytarilmaydi (3 ta tanlov bir-biriga to'g'ri kelmasin)."""
    exclude = exclude or set()
    rx = re.compile(pattern, re.I)
    for c in columns:
        if c not in exclude and rx.search(str(c)):
            return c
    for c in columns[fallback_index:] + columns[:fallback_index]:
        if c not in exclude:
            return c
    return columns[0]


def _clean_amount(series: pd.Series) -> pd.Series:
    """Summani songa aylantiradi. Allaqachon raqamli bo'lsa tegmaydi."""
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)
    s = (
        series.astype(str)
        .str.replace(r"[\s\u00a0\u202f]", "", regex=True)
        .str.replace(",", ".", regex=False)
    )
    # "1.234.567.89" -> faqat oxirgi nuqta o'nlik ajratgich bo'lib qoladi
    s = s.str.replace(r"\.(?=.*\.)", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


# --------------------------------------------------------------------------
# Tushum tabi
# --------------------------------------------------------------------------
def _generate_tushum(
    inspectors: list[Inspector],
    period: str,
    files: dict,
    kuz_assign: dict,
    osa_assign: dict,
) -> list[tuple[str, bytes]] | None:
    with st.spinner("Формирование отчёта…"):
        try:
            with tempfile.TemporaryDirectory(prefix="tushum_out_") as tmp:
                folder = Path(tmp)
                out_dir = folder / "output"
                out_dir.mkdir(exist_ok=True)
                inputs = TushumInputs(
                    kuzatuv_pdf=_save_upload(files["kuz_pdf"], folder, "KUZATUV.pdf"),
                    kuzatuv_xlsx=_save_upload(files["kuz_xlsx"], folder, "KUZATUV.xlsx"),
                    osago_pdf=_save_upload(files["osa_pdf"], folder, "OSAGO.pdf"),
                    osago_xlsx=_save_upload(files["osa_xlsx"], folder, "OSAGO.xlsx"),
                    period_text=period.strip() or DEFAULT_PERIOD,
                    templates_dir=TEMPLATES_DIR,
                    output_dir=out_dir,
                    kuzatuv_assignments=kuz_assign if files["kuz_xlsx"] else None,
                    osago_assignments=osa_assign if files["osa_xlsx"] else None,
                )
                results: list[tuple[str, bytes]] = []
                for item in inspectors:
                    path = generate_tushum_report(item, inputs)
                    results.append((path.name, path.read_bytes()))
                return results
        except Exception as exc:  # noqa: BLE001
            logger.exception("Tushum report generation failed")
            st.error(f"Ошибка: {exc}")
            return None


def _show_tushum_results() -> None:
    results = st.session_state.get("tushum_results")
    if not results:
        return
    st.success(f"Готово: {len(results)} файл(ов)")
    for name, data in results:
        st.download_button(
            label=f"Скачать {name}",
            data=data,
            file_name=name,
            mime=DOCX_MIME,
            key=f"dl_{name}",
        )


def render_tushum_tab() -> None:
    st.subheader("Отчёты Тушум (DOCX)")
    st.caption(
        "PDF/Excel необязательны по отдельности — нужен хотя бы один файл. "
        "Без файла соответствующий абзац останется из шаблона."
    )

    period = st.text_input("Период", value=DEFAULT_PERIOD, help="Например: 24-31-avgust")

    c1, c2 = st.columns(2)
    with c1:
        kuz_pdf = st.file_uploader("KUZATUV PDF (абз. 4)", type=["pdf"], key="kuz_pdf")
        kuz_xlsx = st.file_uploader("KUZATUV Excel (абз. 5)", type=["xlsx"], key="kuz_xlsx")
    with c2:
        osa_pdf = st.file_uploader("OSAGO PDF (абз. 8)", type=["pdf"], key="osa_pdf")
        osa_xlsx = st.file_uploader("OSAGO Excel (абз. 9)", type=["xlsx"], key="osa_xlsx")

    inspector_labels = {item.label_ru: item for item in Inspector}
    selected_label = st.radio(
        "Тип отчёта",
        list(inspector_labels.keys()),
        horizontal=True,
    )
    selected_inspector = inspector_labels[selected_label]

    kuz_assign: dict[Inspector, list[str]] = {}
    osa_assign: dict[Inspector, list[str]] = {}
    kuz_names: list[str] = []
    osa_names: list[str] = []

    try:
        if kuz_xlsx is not None:
            kuz_names = _get_company_names_cached(kuz_xlsx.getvalue(), kuz_xlsx.name, "kuzatuv")
        if osa_xlsx is not None:
            osa_names = _get_company_names_cached(osa_xlsx.getvalue(), osa_xlsx.name, "osago")
    except Exception as exc:  # noqa: BLE001
        logger.exception("Excel parse failed")
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
                    defaults = default_selected_names(kuz_names, item, source="kuzatuv")
                    kuz_assign[item] = st.multiselect(
                        "KUZATUV",
                        kuz_names,
                        default=defaults,
                        # yangi fayl yuklansa eski tanlov qolib ketmasin
                        key=f"kuz_{item.value}_{_file_key(kuz_xlsx)}",
                    )
                else:
                    st.caption("KUZATUV Excel не загружен")
                    kuz_assign[item] = []
            with cols[1]:
                if osa_names:
                    defaults = default_selected_names(osa_names, item, source="osago")
                    osa_assign[item] = st.multiselect(
                        "OSAGO",
                        osa_names,
                        default=defaults,
                        key=f"osa_{item.value}_{_file_key(osa_xlsx)}",
                    )
                else:
                    st.caption("OSAGO Excel не загружен")
                    osa_assign[item] = []

    gen_one, gen_all = st.columns(2)
    do_one = gen_one.button("Сформировать отчёт", type="primary", use_container_width=True)
    do_all = gen_all.button(
        f"Сформировать все {len(Inspector)}", use_container_width=True
    )

    if do_one or do_all:
        # Eski natijani tozalaymiz (tekshiruv xato bersa eskisi ko'rinib qolmasin)
        st.session_state.pop("tushum_results", None)

        inspectors = list(Inspector) if do_all else [selected_inspector]
        error: str | None = None

        if not any([kuz_pdf, kuz_xlsx, osa_pdf, osa_xlsx]):
            error = "Загрузите хотя бы один файл."
        elif UNSAFE_PERIOD_RE.search(period):
            error = 'Период содержит недопустимые символы (\\ / : * ? " < > | ..).'
        elif not TEMPLATES_DIR.exists():
            error = f"Не найдена папка шаблонов: {TEMPLATES_DIR}"
        else:
            for item in inspectors:
                if item not in PERSONAL_INSPECTORS:
                    continue
                if kuz_xlsx and not kuz_assign.get(item):
                    error = f"{item.label_ru}: выберите хотя бы одну компанию KUZATUV."
                    break
                if osa_xlsx and not osa_assign.get(item):
                    error = f"{item.label_ru}: выберите хотя бы одну компанию OSAGO."
                    break

        if error:
            st.warning(error)
        else:
            results = _generate_tushum(
                inspectors,
                period,
                {
                    "kuz_pdf": kuz_pdf,
                    "kuz_xlsx": kuz_xlsx,
                    "osa_pdf": osa_pdf,
                    "osa_xlsx": osa_xlsx,
                },
                kuz_assign,
                osa_assign,
            )
            if results:
                st.session_state["tushum_results"] = results

    # Tugmadan mustaqil: download bosilganda ham natija yo'qolmaydi
    _show_tushum_results()


# --------------------------------------------------------------------------
# Excel-сводка tabi
# --------------------------------------------------------------------------
def _build_report(
    df: pd.DataFrame, col_firma: str, col_polis: str, col_summa: str
) -> tuple[pd.DataFrame, list[str]]:
    """Hisobotni quradi. (natija, ogohlantirishlar) qaytaradi."""
    warnings: list[str] = []
    df = df.copy()

    # Summa
    original = df[col_summa]
    amounts = _clean_amount(original)
    bad = amounts.isna() & original.notna()
    if bad.any():
        warnings.append(f"{int(bad.sum())} ta qatordagi summani o'qib bo'lmadi (0 deb olindi).")
    df[col_summa] = amounts.fillna(0)

    # Firma
    df[col_firma] = df[col_firma].fillna("Noma'lum").astype(str).str.strip()

    # Manba faylda tayyor "Итого" qatori bo'lsa, ikki marta sanalmasin
    is_total = df[col_firma].str.match(TOTAL_ROW_RE)
    if is_total.any():
        warnings.append(
            f"Manba faylda {int(is_total.sum())} ta 'Итого' qatori topildi va hisobdan chiqarildi."
        )
        df = df[~is_total]

    empty_polis = int(df[col_polis].isna().sum())
    if empty_polis:
        warnings.append(
            f"Polis ustunida {empty_polis} ta bo'sh qiymat bor (ular 'Количество' ga kirmaydi)."
        )

    hisobot = (
        df.groupby(col_firma)
        .agg(polis_soni=(col_polis, "count"), res_ins_prem=(col_summa, "sum"))
        .reset_index()
    )
    hisobot["summa_75"] = (hisobot["res_ins_prem"] * 0.75).round(2)
    hisobot["summa_25"] = (hisobot["res_ins_prem"] - hisobot["summa_75"]).round(2)
    hisobot["res_ins_prem"] = hisobot["res_ins_prem"].round(2)
    hisobot["mln_100"] = (hisobot["res_ins_prem"] / 1_000_000).round(3)
    hisobot["mln_75"] = (hisobot["summa_75"] / 1_000_000).round(3)
    hisobot["mln_25"] = (hisobot["summa_25"] / 1_000_000).round(3)

    numeric_cols = [
        "polis_soni",
        "res_ins_prem",
        "summa_75",
        "summa_25",
        "mln_100",
        "mln_75",
        "mln_25",
    ]
    jami = {col_firma: "Общий итог"}
    jami.update({c: hisobot[c].sum() for c in numeric_cols})
    yakuniy = pd.concat([hisobot, pd.DataFrame([jami])], ignore_index=True)
    yakuniy = yakuniy.rename(
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
    return yakuniy, warnings


def _show_excel_results() -> None:
    res = st.session_state.get("excel_result")
    if not res:
        return
    for w in res["warnings"]:
        st.warning(w)
    st.subheader("Natija")
    st.dataframe(res["df"], use_container_width=True)
    st.download_button(
        label="Tayyor hisobotni yuklab olish (.xlsx)",
        data=res["xlsx"],
        file_name="sugurta_hisobot.xlsx",
        mime=XLSX_MIME,
        key="dl_excel_report",
    )


def render_excel_tab() -> None:
    st.subheader("Excel-сводка (75% / 25%)")
    st.write("Birlamchi Excel faylni yuklang va natijaviy hisobotni oling.")

    uploaded_file = st.file_uploader(
        "Excel faylni tanlang (.xlsx yoki .xls)", type=["xlsx", "xls"], key="excel_sum"
    )
    if uploaded_file is None:
        st.session_state.pop("excel_result", None)
        return

    header_row = int(
        st.number_input(
            "Sarlavha qatori raqami (0 = birinchi qator)",
            min_value=0,
            max_value=50,
            value=0,
            step=1,
        )
    )

    try:
        df = _read_excel_cached(uploaded_file.getvalue(), uploaded_file.name, header_row)
    except Exception as exc:  # noqa: BLE001
        logger.exception("Excel read failed")
        st.error(f"Faylni o'qib bo'lmadi: {exc}")
        return

    st.success("Fayl muvaffaqiyatli yuklandi!")
    with st.expander("Yuklangan faylning dastlabki 5 ta qatori"):
        st.dataframe(df.head())

    columns = list(df.columns)
    if len(columns) < 3:
        st.error("Faylda kamida 3 ta ustun bo'lishi kerak.")
        return
    if len(set(columns)) != len(columns):
        st.error("Ustun nomlari takrorlanmoqda. Sarlavha qatorini to'g'rilang.")
        return

    default_firma = _guess_column(
        columns, r"firma|company|наименование|страховщик", 0
    )
    default_polis = _guess_column(
        columns, r"polis|полис|deklar|декларация|\bnd\b|номер", 1, {default_firma}
    )
    default_summa = _guess_column(
        columns,
        r"res_ins_prem|prem|summa|сумма",
        2,
        {default_firma, default_polis},
    )

    st.subheader("Ustunlarni moslashtirish")
    col1, col2, col3 = st.columns(3)
    with col1:
        col_firma = st.selectbox(
            "Firma nomi ustuni:", columns, index=columns.index(default_firma)
        )
    with col2:
        col_polis = st.selectbox(
            "Polis / Deklaratsiya ustuni:", columns, index=columns.index(default_polis)
        )
    with col3:
        col_summa = st.selectbox(
            "Umumiy summa (res_ins_prem) ustuni:",
            columns,
            index=columns.index(default_summa),
        )

    if st.button("Hisobotni shakllantirish", type="primary"):
        st.session_state.pop("excel_result", None)
        if len({col_firma, col_polis, col_summa}) < 3:
            st.warning("Firma, polis va summa uchun turli ustunlarni tanlang.")
        else:
            try:
                yakuniy_df, warns = _build_report(df, col_firma, col_polis, col_summa)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    yakuniy_df.to_excel(writer, index=False, sheet_name="Hisobot")
                st.session_state["excel_result"] = {
                    "df": yakuniy_df,
                    "xlsx": output.getvalue(),
                    "warnings": warns,
                }
            except Exception as exc:  # noqa: BLE001
                logger.exception("Excel report failed")
                st.error(f"Xatolik yuz berdi: {exc}")

    _show_excel_results()


# --------------------------------------------------------------------------
def run() -> None:
    st.set_page_config(page_title="Sug'urta hisoboti", layout="wide")
    st.title("Sug'urta hisobotini shakllantirish dasturi")

    tab_tushum, tab_excel = st.tabs(["Тушум (DOCX)", "Excel-сводка"])
    with tab_tushum:
        render_tushum_tab()
    with tab_excel:
        render_excel_tab()


if __name__ == "__main__":
    run()
