from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path


class Inspector(str, Enum):
    ARIPOV = "aripov"
    ABDIMURODOV = "abdimurodov"
    XAMRAKULOV = "xamrakulov"

    @property
    def label_ru(self) -> str:
        return {
            Inspector.ARIPOV: "Тушум — Ш.Арипов",
            Inspector.ABDIMURODOV: "Тушум — Ш.Абдимуродов",
            Inspector.XAMRAKULOV: "Тушум — Б.Хамракулов",
        }[self]


@dataclass(frozen=True)
class InspectorProfile:
    inspector: Inspector
    template_name: str
    output_prefix: str
    # Personal assigned companies for period paragraphs (None => use Excel totals)
    kuzatuv_company_keywords: list[str] | None
    osago_company_keywords: list[str] | None
    # Fixed company counts in text (do not recompute)
    kuzatuv_company_count_text: str | None
    osago_company_count_text: str | None
    # Display names inside parentheses for personal reports
    kuzatuv_company_names_text: str | None
    osago_company_names_text: str | None
    style: str  # "personal" | "chief"


PROFILES: dict[Inspector, InspectorProfile] = {
    Inspector.ARIPOV: InspectorProfile(
        inspector=Inspector.ARIPOV,
        template_name="tushum_aripov.docx",
        output_prefix="Tushum_Aripov",
        kuzatuv_company_keywords=[
            "TEMIR YOL",
            "OZBEKINVEST",
            "XALQ",
            "INFINITY",
        ],
        osago_company_keywords=["ALFA INVEST", "INSON"],
        kuzatuv_company_count_text="4 ta",
        osago_company_count_text="2 ta",
        kuzatuv_company_names_text=(
            "“TEMIR YOʻL SUGʻURTA”, “OʻZBEKINVEST SUGʻURTA”, "
            "“XALQ SUGʻURTA” va “INFINITY INSURANCE”"
        ),
        osago_company_names_text="“ALFA INVEST” AJ va “INSON” AJ",
        style="personal",
    ),
    Inspector.ABDIMURODOV: InspectorProfile(
        inspector=Inspector.ABDIMURODOV,
        template_name="tushum_abdimurodov.docx",
        output_prefix="Tushum_Abdimurodov",
        kuzatuv_company_keywords=["ALFA", "GROSS INSURANCE", "INSON"],
        osago_company_keywords=["GROSS"],
        kuzatuv_company_count_text="3 ta",
        osago_company_count_text="1 ta",
        kuzatuv_company_names_text=(
            "“ALFA”, “GROSS INSURANCE” va “INSON”"
        ),
        osago_company_names_text="“GROSS Sugurta Kompaniyasi” AJ",
        style="personal",
    ),
    Inspector.XAMRAKULOV: InspectorProfile(
        inspector=Inspector.XAMRAKULOV,
        template_name="tushum_xamrakulov.docx",
        output_prefix="Tushum_Xamrakulov",
        kuzatuv_company_keywords=None,
        osago_company_keywords=None,
        kuzatuv_company_count_text=None,
        osago_company_count_text=None,
        kuzatuv_company_names_text=None,
        osago_company_names_text=None,
        style="chief",
    ),
}


def template_path(profile: InspectorProfile, templates_dir: Path) -> Path:
    return templates_dir / profile.template_name
