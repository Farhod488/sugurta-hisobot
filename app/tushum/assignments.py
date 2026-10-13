from __future__ import annotations

from app.tushum.formatters import company_matches, normalize_company
from app.tushum.parsers import ExcelTotals
from app.tushum.profiles import PROFILES, Inspector


PERSONAL_INSPECTORS = (Inspector.ARIPOV, Inspector.ABDIMURODOV)


def format_company_count(count: int) -> str:
    return f"{count} ta"


def format_company_names_uz(names: list[str]) -> str:
    quoted = [f"“{name.strip()}”" for name in names if name.strip()]
    if not quoted:
        return ""
    if len(quoted) == 1:
        return quoted[0]
    if len(quoted) == 2:
        return f"{quoted[0]} va {quoted[1]}"
    return ", ".join(quoted[:-1]) + f" va {quoted[-1]}"


def company_noun(count: int) -> str:
    return "sugʻurta tashkiloti" if count == 1 else "sugʻurta tashkilotlari"


def default_selected_names(
    company_names: list[str],
    inspector: Inspector,
    *,
    source: str,
) -> list[str]:
    profile = PROFILES[inspector]
    keywords = (
        profile.kuzatuv_company_keywords
        if source == "kuzatuv"
        else profile.osago_company_keywords
    )
    if not keywords:
        return []
    selected: list[str] = []
    for name in company_names:
        if company_matches(name, keywords):
            selected.append(name)
    return selected


def filter_by_names(data: ExcelTotals, names: list[str], *, source_label: str) -> ExcelTotals:
    wanted = {normalize_company(name) for name in names}
    matched = [row for row in data.rows if normalize_company(row.name) in wanted]
    if not matched:
        found = ", ".join(row.name for row in data.rows) or "(пусто)"
        raise ValueError(
            f"В Excel «{source_label}» не найдены выбранные компании: "
            f"{', '.join(names) or '(не выбрано)'}.\n"
            f"Сейчас в файле: {found}."
        )
    return ExcelTotals(
        rows=matched,
        policies=sum(r.policies for r in matched),
        premium_soum=sum(r.premium_soum for r in matched),
        company_share_soum=sum(r.company_share_soum for r in matched),
        mbl_share_soum=sum(r.mbl_share_soum for r in matched),
    )
