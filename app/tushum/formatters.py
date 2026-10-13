from __future__ import annotations

import math
import re


def parse_number(value: object) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text in {"-", "—", "–"}:
        return 0.0
    text = text.replace("\xa0", " ").replace(" ", "").replace("%", "")
    text = text.replace(",", ".")
    return float(text)


def format_int_spaces(value: float | int) -> str:
    number = int(round(float(value)))
    return f"{number:,}".replace(",", " ")


def format_decimal(value: float, digits: int = 1, *, drop_trailing_zero: bool = False) -> str:
    quantized = round(float(value) + 1e-9, digits)
    if digits == 0:
        return format_int_spaces(quantized)
    text = f"{quantized:.{digits}f}"
    whole, frac = text.split(".")
    whole_fmt = format_int_spaces(int(whole))
    if drop_trailing_zero and int(frac) == 0:
        return whole_fmt
    return f"{whole_fmt},{frac}"


def format_policy_count_words(count: float | int) -> str:
    """1345623 -> '1 mln 345 ming 623'; 480536 -> '480 536'."""
    number = int(round(float(count)))
    if number < 1_000_000:
        return format_int_spaces(number)
    millions = number // 1_000_000
    rest = number % 1_000_000
    thousands = rest // 1_000
    units = rest % 1_000
    parts = [f"{millions} mln"]
    if thousands:
        parts.append(f"{thousands} ming")
    if units:
        parts.append(str(units))
    elif not thousands:
        pass
    return " ".join(parts)


def mln_to_mlrd(mln: float, digits: int = 1) -> str:
    return format_decimal(mln / 1000.0, digits=digits, drop_trailing_zero=True)


def soum_to_display(soum: float) -> tuple[str, str]:
    """Return (formatted_value, unit) where unit is 'mlrd' or 'mln'."""
    if abs(soum) >= 1_000_000_000:
        return format_decimal(soum / 1_000_000_000.0, 1), "mlrd"
    return format_decimal(soum / 1_000_000.0, 1), "mln"


def soum_to_mln(soum: float, digits: int = 1) -> str:
    return format_decimal(soum / 1_000_000.0, digits=digits, drop_trailing_zero=True)


def normalize_company(name: str) -> str:
    text = str(name).upper()
    # Common OCR / typography variants and Excel quirks
    replacements = {
        "ʻ": "",
        "ʼ": "",
        "'": "",
        "’": "",
        "‘": "",
        "“": "",
        "”": "",
        '"': "",
        "«": "",
        "»": "",
        "-": " ",
        "_": " ",
        ".": " ",
        "\u00a0": " ",
        "\u202f": " ",
        "Ё": "Е",
        "Ğ": "G",
        "Ö": "O",
        "Ü": "U",
        "ʻO": "O",
        "Oʻ": "O",
        "O'": "O",
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    text = re.sub(r"[^0-9A-ZА-ЯЎҚҒҲ\s]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def company_matches(name: str, keywords: list[str]) -> bool:
    hay = normalize_company(name)
    if not hay:
        return False
    for key in keywords:
        needle = normalize_company(key)
        if not needle:
            continue
        if needle in hay:
            return True
        # Token fallback: all keyword tokens must appear in the name.
        tokens = [t for t in needle.split() if len(t) > 1]
        if tokens and all(token in hay for token in tokens):
            return True
    return False