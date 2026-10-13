from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd
import pdfplumber

from app.tushum.formatters import company_matches, parse_number


@dataclass
class PeriodTotals:
    policies: float
    premium_mln: float
    company_share_mln: float
    mbl_share_mln: float


@dataclass
class ExcelCompanyRow:
    name: str
    policies: float
    premium_soum: float
    company_share_soum: float
    mbl_share_soum: float


@dataclass
class ExcelTotals:
    rows: list[ExcelCompanyRow]
    policies: float
    premium_soum: float
    company_share_soum: float
    mbl_share_soum: float


def parse_kuzatuv_pdf(path: Path | str) -> PeriodTotals:
    path = Path(path)
    with pdfplumber.open(path) as pdf:
        tables = pdf.pages[0].extract_tables() or []
    if not tables:
        raise ValueError(f"В PDF нет таблиц: {path.name}")

    for row in tables[0]:
        if not row:
            continue
        label = str(row[0] or "").strip()
        if label.startswith("Jami"):
            # Jami: | | 648,5 | 100% | policies | premium | company | mbl | ...
            return PeriodTotals(
                policies=parse_number(row[4]),
                premium_mln=parse_number(row[5]),
                company_share_mln=parse_number(row[6]),
                mbl_share_mln=parse_number(row[7]),
            )
    raise ValueError(f"Строка Jami не найдена в {path.name}")


def parse_osago_pdf(path: Path | str) -> PeriodTotals:
    path = Path(path)
    with pdfplumber.open(path) as pdf:
        tables = pdf.pages[0].extract_tables() or []
    if not tables:
        raise ValueError(f"В PDF нет таблиц: {path.name}")

    for row in tables[0]:
        if not row:
            continue
        label = str(row[1] or row[0] or "").strip()
        if label.startswith("Jami"):
            # Jami: deposit% transport policies AT MB ... premium company mbl ...
            return PeriodTotals(
                policies=parse_number(row[5]),
                premium_mln=parse_number(row[12]),
                company_share_mln=parse_number(row[13]),
                mbl_share_mln=parse_number(row[14]),
            )
    raise ValueError(f"Строка Jami не найдена в {path.name}")


def _read_excel_rows(path: Path, *, is_osago: bool) -> ExcelTotals:
    frame = pd.read_excel(path, header=0)
    if frame.shape[1] < 5:
        raise ValueError(f"Ожидалось минимум 5 колонок в {path.name}")

    name_col, pol_col, prem_col, company_col, mbl_col = frame.columns[:5]
    rows: list[ExcelCompanyRow] = []
    total: ExcelCompanyRow | None = None

    for _, item in frame.iterrows():
        name = str(item[name_col]).strip()
        if not name or name.lower() == "nan":
            continue
        row = ExcelCompanyRow(
            name=name,
            policies=parse_number(item[pol_col]),
            premium_soum=parse_number(item[prem_col]),
            company_share_soum=parse_number(item[company_col]),
            mbl_share_soum=parse_number(item[mbl_col]),
        )
        lowered = name.lower().replace("ʻ", "'").replace("'", "")
        # Match only total labels, not company names like "jamiyati".
        if (
            "общий" in lowered
            or "общая" in lowered
            or lowered.strip() in {"jami", "jami:", "total", "итого"}
            or lowered.startswith("общий итог")
            or lowered.startswith("jami:")
        ):
            total = row
            continue
        rows.append(row)

    if total is None:
        total = ExcelCompanyRow(
            name="TOTAL",
            policies=sum(r.policies for r in rows),
            premium_soum=sum(r.premium_soum for r in rows),
            company_share_soum=sum(r.company_share_soum for r in rows),
            mbl_share_soum=sum(r.mbl_share_soum for r in rows),
        )

    # For KUZATUV.xlsx column 4 is agent/MBL share; OSAGO uses explicit MBL column.
    _ = is_osago
    return ExcelTotals(
        rows=rows,
        policies=total.policies,
        premium_soum=total.premium_soum,
        company_share_soum=total.company_share_soum,
        mbl_share_soum=total.mbl_share_soum,
    )


def parse_kuzatuv_xlsx(path: Path | str) -> ExcelTotals:
    return _read_excel_rows(Path(path), is_osago=False)


def parse_osago_xlsx(path: Path | str) -> ExcelTotals:
    return _read_excel_rows(Path(path), is_osago=True)


def filter_companies(
    data: ExcelTotals,
    keywords: list[str],
    *,
    source_label: str | None = None,
) -> ExcelTotals:
    matched = [row for row in data.rows if company_matches(row.name, keywords)]
    if not matched:
        found = ", ".join(row.name for row in data.rows) or "(пусто)"
        where = f" «{source_label}»" if source_label else ""
        raise ValueError(
            f"В Excel{where} не найдены компании: {', '.join(keywords)}.\n"
            f"Сейчас в файле: {found}.\n"
            f"Проверьте, что выбран правильный файл "
            f"(для Арипова / абзаца 5 нужен KUZATUV.xlsx, не OSAGO)."
        )
    return ExcelTotals(
        rows=matched,
        policies=sum(r.policies for r in matched),
        premium_soum=sum(r.premium_soum for r in matched),
        company_share_soum=sum(r.company_share_soum for r in matched),
        mbl_share_soum=sum(r.mbl_share_soum for r in matched),
    )
