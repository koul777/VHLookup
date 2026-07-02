from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from typing import Sequence

import pandas as pd

from vhlookup_core.models import HeaderDetectionResult, SheetData
from vhlookup_core.normalization import display_value, is_blank, normalize_header


HEADER_KEYWORDS = {
    "사번",
    "직원번호",
    "직원명",
    "성명",
    "이름",
    "부서",
    "소속",
    "직급",
    "직위",
    "일자",
    "날짜",
    "금액",
    "수량",
    "단가",
    "예산과목",
    "사업명",
    "학교명",
    "학교코드",
    "기관명",
    "기관코드",
    "담당자",
    "현원",
    "연락처",
    "휴대폰",
    "이메일",
    "제출일",
    "지급월",
    "이수여부",
    "교육명",
    "주소",
    "구분",
    "항목",
    "월",
    "분기",
    "id",
    "name",
    "date",
    "amount",
    "department",
    "position",
}


@dataclass(frozen=True)
class _Candidate:
    row_index: int
    row_count: int
    headers: tuple[str, ...]
    score: float
    warnings: tuple[str, ...]


class HeaderDetector:
    def detect(
        self,
        sheet: SheetData,
        scan_rows: int | str = 30,
    ) -> HeaderDetectionResult:
        frame = sheet.frame
        if frame.empty:
            return HeaderDetectionResult(
                sheet_name=sheet.name,
                header_row_index=0,
                header_row_count=1,
                data_start_row_index=1,
                headers=(),
                confidence=0.0,
                warnings=("빈 시트입니다.",),
            )

        row_limit = len(frame.index) if scan_rows == "all" else min(int(scan_rows), len(frame.index))
        candidates: list[_Candidate] = []
        for row_index in range(row_limit):
            single = self._score_single(frame, row_index)
            if single is not None:
                candidates.append(single)
            if row_index + 1 < row_limit:
                pair = self._score_two_line(frame, row_index)
                if pair is not None:
                    candidates.append(pair)

        if not candidates:
            headers = unique_headers([f"Column {index + 1}" for index in range(frame.shape[1])])
            return HeaderDetectionResult(
                sheet_name=sheet.name,
                header_row_index=0,
                header_row_count=1,
                data_start_row_index=1,
                headers=tuple(headers),
                confidence=0.0,
                warnings=("헤더 후보를 찾지 못했습니다. 사용자가 직접 선택해야 합니다.",),
            )

        best = max(candidates, key=lambda item: item.score)
        confidence = max(0.0, min(1.0, best.score))
        warnings = list(best.warnings)
        if confidence < 0.55:
            warnings.append("헤더 자동탐지 신뢰도가 낮습니다. 수동 확인이 필요합니다.")

        return HeaderDetectionResult(
            sheet_name=sheet.name,
            header_row_index=best.row_index,
            header_row_count=best.row_count,
            data_start_row_index=best.row_index + best.row_count,
            headers=best.headers,
            confidence=confidence,
            warnings=tuple(warnings),
        )

    def apply(self, sheet: SheetData, detection: HeaderDetectionResult) -> pd.DataFrame:
        data = sheet.frame.iloc[detection.data_start_row_index :].copy()
        data = data.iloc[:, : len(detection.headers)]
        data.columns = list(detection.headers)
        data = data.dropna(how="all").reset_index(drop=True)
        return data

    def _score_single(self, frame: pd.DataFrame, row_index: int) -> _Candidate | None:
        values = trim_trailing_blanks([display_value(value) for value in frame.iloc[row_index].tolist()])
        non_empty = [value for value in values if value]
        if len(non_empty) < 2:
            return None
        row_width = max(1, len(values))
        fill_ratio = len(non_empty) / row_width
        avg_length = sum(len(value) for value in non_empty) / len(non_empty)
        keyword_ratio = self._keyword_ratio(non_empty)
        next_data_score = self._next_rows_data_score(frame, row_index + 1, len(values))
        title_penalty = 0.25 if avg_length > 28 and fill_ratio < 0.45 else 0.0
        sparse_penalty = 0.15 if fill_ratio < 0.25 else 0.0
        duplicate_penalty = self._duplicate_penalty(non_empty)
        score = (
            0.28 * min(fill_ratio * 1.8, 1.0)
            + 0.34 * keyword_ratio
            + 0.28 * next_data_score
            + 0.10 * min(len(non_empty) / 5, 1.0)
            - title_penalty
            - sparse_penalty
            - duplicate_penalty
        )
        headers = unique_headers(values)
        warnings = ("중복 컬럼명이 있어 자동 구분명을 부여했습니다.",) if duplicate_penalty else ()
        return _Candidate(row_index, 1, tuple(headers), score, warnings)

    def _score_two_line(self, frame: pd.DataFrame, row_index: int) -> _Candidate | None:
        top = [display_value(value) for value in frame.iloc[row_index].tolist()]
        bottom = [display_value(value) for value in frame.iloc[row_index + 1].tolist()]
        bottom_non_empty = [value for value in bottom if value]
        if len(bottom_non_empty) < 2:
            return None
        top_non_empty = [value for value in top if value]
        if not top_non_empty:
            return None
        top_unique_count = len({normalize_header(value) for value in top_non_empty})
        if len(top_non_empty) == 1 and top_unique_count == 1:
            return None
        has_grouping_signal = len(top_non_empty) < len(bottom_non_empty) or top_unique_count < len(top_non_empty)
        if not has_grouping_signal:
            return None

        filled_top = forward_fill_labels(top)
        combined = []
        for upper, lower in zip(filled_top, bottom):
            if lower and upper and normalize_header(upper) != normalize_header(lower):
                combined.append(f"{upper}_{lower}")
            elif lower:
                combined.append(lower)
            elif upper:
                combined.append(upper)
            else:
                combined.append("")
        combined = trim_trailing_blanks(combined)
        combined_non_empty = [value for value in combined if value]
        if len(combined_non_empty) < 2:
            return None

        keyword_ratio = self._keyword_ratio(combined_non_empty)
        next_data_score = self._next_rows_data_score(frame, row_index + 2, len(combined))
        group_header_bonus = min(len(top_non_empty) / max(len(bottom_non_empty), 1), 0.5)
        score = 0.30 * keyword_ratio + 0.38 * next_data_score + 0.32 + 0.10 * group_header_bonus
        score = min(score, 1.0)
        duplicate_penalty = self._duplicate_penalty(combined_non_empty)
        headers = unique_headers(combined_non_empty)
        warnings = ("2줄 헤더를 결합했습니다.",)
        if duplicate_penalty:
            warnings = warnings + ("중복 컬럼명이 있어 자동 구분명을 부여했습니다.",)
        return _Candidate(row_index, 2, tuple(headers), score - duplicate_penalty, warnings)

    def _keyword_ratio(self, values: Sequence[str]) -> float:
        if not values:
            return 0.0
        hits = 0
        for value in values:
            normalized = normalize_header(value)
            if normalized in HEADER_KEYWORDS or any(keyword in normalized for keyword in HEADER_KEYWORDS):
                hits += 1
        return min(hits / max(len(values), 1) * 2.0, 1.0)

    def _next_rows_data_score(self, frame: pd.DataFrame, start: int, width: int) -> float:
        if start >= len(frame.index):
            return 0.0
        sample = frame.iloc[start : min(start + 5, len(frame.index)), :width]
        if sample.empty:
            return 0.0
        row_scores = []
        for _, row in sample.iterrows():
            non_empty = sum(0 if is_blank(value) else 1 for value in row.tolist())
            row_scores.append(non_empty / max(width, 1))
        return min(sum(row_scores) / max(len(row_scores), 1) * 1.5, 1.0)

    def _duplicate_penalty(self, values: Sequence[str]) -> float:
        normalized = [normalize_header(value) for value in values if value]
        counts = Counter(normalized)
        duplicates = sum(count - 1 for count in counts.values() if count > 1)
        return min(duplicates * 0.05, 0.20)


def forward_fill_labels(values: Sequence[str]) -> list[str]:
    labels: list[str] = []
    current = ""
    for value in values:
        if value:
            current = value
        labels.append(current)
    return labels


def trim_trailing_blanks(values: Sequence[str]) -> list[str]:
    result = list(values)
    while result and not result[-1]:
        result.pop()
    return result


def unique_headers(values: Sequence[str]) -> list[str]:
    counts: Counter[str] = Counter()
    result = []
    for index, value in enumerate(values):
        base = value.strip() if value and value.strip() else f"Column {index + 1}"
        key = normalize_header(base) or f"column{index + 1}"
        counts[key] += 1
        if counts[key] == 1:
            result.append(base)
        else:
            result.append(f"{base}__{counts[key]}")
    return result
