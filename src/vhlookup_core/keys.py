from __future__ import annotations

from itertools import combinations

import pandas as pd

from vhlookup_core.mapper import ColumnMapper
from vhlookup_core.models import KeyCandidate
from vhlookup_core.normalization import normalize_header, normalize_key_parts


_IDENTIFIER_KEY_HINTS = ("사번", "직원번호", "직원id", "employeeid", "empid", "기관코드", "orgcode")
_PERIOD_KEY_HINTS = ("지급월", "기준월", "귀속월", "월", "날짜", "일자", "date", "month")


class KeyRecommender:
    def __init__(self, mapper: ColumnMapper | None = None) -> None:
        self.mapper = mapper or ColumnMapper()

    def recommend(
        self,
        reference: pd.DataFrame,
        target: pd.DataFrame,
        max_composite_size: int = 2,
        limit: int = 5,
    ) -> list[KeyCandidate]:
        mapping = self.mapper.map_columns(list(target.columns), list(reference.columns), threshold=0.80)
        pairs = [
            (reference_column, target_column)
            for reference_column, target_column in mapping.target_to_source.items()
        ]
        pairs = self._with_data_overlap_pairs(reference, target, pairs)
        candidates: list[KeyCandidate] = []
        for size in range(1, min(max_composite_size, len(pairs)) + 1):
            for combo in combinations(pairs, size):
                reference_cols = tuple(item[0] for item in combo)
                target_cols = tuple(item[1] for item in combo)
                candidates.append(self._score(reference, target, reference_cols, target_cols))
        return sorted(candidates, key=lambda item: item.score, reverse=True)[:limit]

    def _with_data_overlap_pairs(
        self,
        reference: pd.DataFrame,
        target: pd.DataFrame,
        name_pairs: list[tuple[str, str]],
    ) -> list[tuple[str, str]]:
        seen = set(name_pairs)
        used_reference_columns = {reference_column for reference_column, _target_column in name_pairs}
        used_target_columns = {target_column for _reference_column, target_column in name_pairs}
        data_pairs: list[tuple[float, tuple[str, str]]] = []
        for reference_column in reference.columns:
            if reference_column in used_reference_columns:
                continue
            for target_column in target.columns:
                if target_column in used_target_columns:
                    continue
                pair = (reference_column, target_column)
                if pair in seen:
                    continue
                score = self._single_column_data_score(reference[reference_column], target[target_column])
                if score >= 0.75:
                    data_pairs.append((score, pair))

        ordered_pairs = list(name_pairs)
        for _score, pair in sorted(data_pairs, key=lambda item: item[0], reverse=True):
            reference_column, target_column = pair
            if pair in seen or reference_column in used_reference_columns or target_column in used_target_columns:
                continue
            ordered_pairs.append(pair)
            seen.add(pair)
            used_reference_columns.add(reference_column)
            used_target_columns.add(target_column)
            if len(ordered_pairs) >= 12:
                break
        return ordered_pairs

    def _single_column_data_score(self, reference: pd.Series, target: pd.Series) -> float:
        reference_keys = reference.to_frame().apply(lambda row: normalize_key_parts(row.tolist()), axis=1)
        target_keys = target.to_frame().apply(lambda row: normalize_key_parts(row.tolist()), axis=1)
        loose_reference_keys = reference.to_frame().apply(
            lambda row: normalize_key_parts(row.tolist(), "loose_numeric"),
            axis=1,
        )
        loose_target_keys = target.to_frame().apply(
            lambda row: normalize_key_parts(row.tolist(), "loose_numeric"),
            axis=1,
        )

        reference_non_blank = reference_keys[reference_keys != ""]
        target_non_blank = target_keys[target_keys != ""]
        loose_reference_non_blank = loose_reference_keys[loose_reference_keys != ""]
        loose_target_non_blank = loose_target_keys[loose_target_keys != ""]
        if reference_non_blank.empty or target_non_blank.empty:
            return 0.0

        exact_overlap = len(set(reference_non_blank).intersection(set(target_non_blank)))
        exact_denominator = max(min(reference_non_blank.nunique(), target_non_blank.nunique()), 1)
        loose_overlap = len(set(loose_reference_non_blank).intersection(set(loose_target_non_blank)))
        loose_denominator = max(min(loose_reference_non_blank.nunique(), loose_target_non_blank.nunique()), 1)
        overlap_ratio = max(exact_overlap / exact_denominator, (loose_overlap / loose_denominator) * 0.92)
        completeness = min(
            len(reference_non_blank) / max(len(reference), 1),
            len(target_non_blank) / max(len(target), 1),
        )
        uniqueness = 1 - max(duplicate_ratio(reference_non_blank), duplicate_ratio(target_non_blank))
        return 0.55 * overlap_ratio + 0.20 * completeness + 0.25 * uniqueness

    def _score(
        self,
        reference: pd.DataFrame,
        target: pd.DataFrame,
        reference_cols: tuple[str, ...],
        target_cols: tuple[str, ...],
    ) -> KeyCandidate:
        reference_keys = build_key_series(reference, reference_cols)
        target_keys = build_key_series(target, target_cols)
        loose_reference_keys = build_key_series(reference, reference_cols, normalization="loose_numeric")
        loose_target_keys = build_key_series(target, target_cols, normalization="loose_numeric")
        reference_non_blank = reference_keys[reference_keys != ""]
        target_non_blank = target_keys[target_keys != ""]
        loose_reference_non_blank = loose_reference_keys[loose_reference_keys != ""]
        loose_target_non_blank = loose_target_keys[loose_target_keys != ""]
        duplicate_reference = duplicate_ratio(reference_non_blank)
        duplicate_target = duplicate_ratio(target_non_blank)
        overlap = len(set(reference_non_blank).intersection(set(target_non_blank)))
        denominator = max(min(reference_non_blank.nunique(), target_non_blank.nunique()), 1)
        overlap_ratio = overlap / denominator
        loose_overlap = len(set(loose_reference_non_blank).intersection(set(loose_target_non_blank)))
        loose_denominator = max(min(loose_reference_non_blank.nunique(), loose_target_non_blank.nunique()), 1)
        loose_overlap_ratio = loose_overlap / loose_denominator
        effective_overlap_ratio = max(overlap_ratio, loose_overlap_ratio * 0.92)
        completeness = min(
            len(reference_non_blank) / max(len(reference), 1),
            len(target_non_blank) / max(len(target), 1),
        )
        uniqueness = 1 - max(duplicate_reference, duplicate_target)
        business_key_bonus = 0.03 if self._looks_like_period_composite_key(reference_cols, target_cols) else 0.0
        score = 0.50 * effective_overlap_ratio + 0.25 * completeness + 0.25 * uniqueness + business_key_bonus
        return KeyCandidate(
            reference_columns=reference_cols,
            target_columns=target_cols,
            score=round(max(0.0, min(score, 1.0)), 4),
            duplicate_ratio_reference=round(duplicate_reference, 4),
            duplicate_ratio_target=round(duplicate_target, 4),
            loose_overlap_score=round(loose_overlap_ratio, 4),
        )

    def _looks_like_period_composite_key(
        self,
        reference_cols: tuple[str, ...],
        target_cols: tuple[str, ...],
    ) -> bool:
        if len(reference_cols) < 2:
            return False
        normalized = [normalize_header(column) for column in (*reference_cols, *target_cols)]
        has_identifier = any(any(hint in column for hint in _IDENTIFIER_KEY_HINTS) for column in normalized)
        has_period = any(any(hint in column for hint in _PERIOD_KEY_HINTS) for column in normalized)
        return has_identifier and has_period


def build_key_series(
    frame: pd.DataFrame,
    columns: tuple[str, ...],
    normalization: str = "text",
) -> pd.Series:
    if not columns:
        raise ValueError("At least one key column is required.")
    missing = [column for column in columns if column not in frame.columns]
    if missing:
        raise KeyError(f"Missing key columns: {missing}")
    return frame.loc[:, list(columns)].apply(
        lambda row: normalize_key_parts(row.tolist(), normalization),
        axis=1,
    )


def duplicate_ratio(series: pd.Series) -> float:
    if series.empty:
        return 0.0
    duplicate_count = int(series.duplicated(keep=False).sum())
    return duplicate_count / len(series)
