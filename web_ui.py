"""Streamlit UI: Tushum DOCX reports."""

from __future__ import annotations

import logging
import re
import tempfile
from pathlib import Path

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
DEFAULT_PERIOD = "24-31-avgust"
UNSAFE_PERIOD_RE = re.compile(r'[\\/:*?"<>|]|\.\.')


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

    _show_tushum_results()


def run() -> None:
    render_tushum_tab()


if __name__ == "__main__":
    run()
