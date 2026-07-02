from __future__ import annotations

import math
import re
import unicodedata
from datetime import date, datetime
from typing import Iterable

import pandas as pd


_SPACE_RE = re.compile(r"\s+")
_HEADER_PUNCT_RE = re.compile(r"[\s_\-./\\()\[\]{}:;|]+")


def is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def display_value(value: object) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return _SPACE_RE.sub(" ", str(value).strip())


def normalize_header(value: object) -> str:
    text = unicodedata.normalize("NFKC", display_value(value)).lower()
    return _HEADER_PUNCT_RE.sub("", text)


def normalize_key_value(value: object, mode: str = "text") -> str:
    text = unicodedata.normalize("NFKC", display_value(value))
    text = _SPACE_RE.sub(" ", text).strip()
    if mode == "loose_numeric" and re.fullmatch(r"0*\d+", text):
        stripped = text.lstrip("0")
        return stripped or "0"
    return text


def normalize_key_parts(values: Iterable[object], mode: str = "text") -> str:
    return "\x1f".join(normalize_key_value(value, mode) for value in values)


def looks_like_numeric_id_mismatch(left: object, right: object) -> bool:
    left_text = normalize_key_value(left, "text")
    right_text = normalize_key_value(right, "text")
    if left_text == right_text:
        return False
    return normalize_key_value(left, "loose_numeric") == normalize_key_value(
        right, "loose_numeric"
    )
