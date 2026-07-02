from __future__ import annotations

from itertools import combinations

import pandas as pd

from vhlookup_core.mapper import ColumnMapper
from vhlookup_core.models import KeyCandidate
from vhlookup_core.normalization import normalize_key_parts


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
        candidates: list[KeyCandidate] = []
        for size in range(1, min(max_composite_size, len(pairs)) + 1):
            for combo in combinations(pairs, size):
                reference_cols = tuple(item[0] for item in combo)
                target_cols = tuple(item[1] for item in combo)
                candidates.append(self._score(reference, target, reference_cols, target_cols))
        return sorted(candidates, key=lambda item: item.score, reverse=True)[:limit]

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
        score = 0.52 * effective_overlap_ratio + 0.28 * completeness + 0.20 * (1 - duplicate_reference)
        return KeyCandidate(
            reference_columns=reference_cols,
            target_columns=target_cols,
            score=round(max(0.0, min(score, 1.0)), 4),
            duplicate_ratio_reference=round(duplicate_reference, 4),
            duplicate_ratio_target=round(duplicate_target, 4),
            loose_overlap_score=round(loose_overlap_ratio, 4),
        )


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
