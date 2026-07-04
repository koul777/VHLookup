from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Literal

import pandas as pd

from vhlookup_core.keys import KeyRecommender, build_key_series
from vhlookup_core.mapper import ColumnMapper
from vhlookup_core.models import KeySpec
from vhlookup_core.normalization import normalize_header


@dataclass(frozen=True)
class AutoLookupPlan:
    key_spec: KeySpec
    value_columns: tuple[str, ...]
    confidence: float
    evidence_rows: tuple[dict[str, Any], ...] = ()
    warnings: tuple[str, ...] = ()


class AutoLookupPlanner:
    def __init__(
        self,
        key_recommender: KeyRecommender | None = None,
        mapper: ColumnMapper | None = None,
    ) -> None:
        self.mapper = mapper or ColumnMapper()
        self.key_recommender = key_recommender or KeyRecommender(self.mapper)

    def infer_lookup_plan(
        self,
        reference: pd.DataFrame,
        target: pd.DataFrame,
        preferred_reference_key_columns: tuple[str, ...] = (),
        preferred_target_key_columns: tuple[str, ...] = (),
        preferred_value_columns: tuple[str, ...] = (),
    ) -> AutoLookupPlan:
        warnings: list[str] = []
        evidence_rows: list[dict[str, Any]] = []
        key_spec: KeySpec
        confidence = 1.0

        if preferred_reference_key_columns:
            target_key_columns = self._resolve_target_key_columns(
                target,
                preferred_reference_key_columns,
                preferred_target_key_columns,
            )
            normalization = self.infer_key_normalization(
                reference,
                target,
                preferred_reference_key_columns,
                target_key_columns,
            )
            key_spec = KeySpec(
                reference_key_columns=preferred_reference_key_columns,
                target_key_columns=target_key_columns,
                normalization=normalization,
            )
            for reference_column, target_column in zip(preferred_reference_key_columns, target_key_columns):
                evidence_rows.append(
                    {
                        "역할": "키 컬럼",
                        "기준표 컬럼": reference_column,
                        "대상표 컬럼": target_column,
                        "신뢰도": 1.0,
                        "추천 방식": "템플릿/사용자 입력 기반 자동 매칭",
                        "검토 메모": "선택한 업무 템플릿 또는 입력값을 기준으로 대상표 컬럼명을 맞췄습니다.",
                    }
                )
        else:
            candidates = self.key_recommender.recommend(reference, target, limit=1)
            if not candidates:
                raise ValueError("자동으로 키 컬럼 후보를 찾지 못했습니다.")
            best = candidates[0]
            confidence = best.score
            normalization = self.infer_key_normalization(
                reference,
                target,
                best.reference_columns,
                best.target_columns,
            )
            key_spec = KeySpec(
                reference_key_columns=best.reference_columns,
                target_key_columns=best.target_columns,
                normalization=normalization,
            )
            for reference_column, target_column in zip(best.reference_columns, best.target_columns):
                evidence_rows.append(
                    {
                        "역할": "키 컬럼",
                        "기준표 컬럼": reference_column,
                        "대상표 컬럼": target_column,
                        "신뢰도": best.score,
                        "추천 방식": "컬럼명 유사도 + 데이터 겹침",
                        "검토 메모": f"느슨한 형식 비교 겹침 점수: {best.loose_overlap_score}",
                    }
                )
            if best.score < 0.55:
                warnings.append("키 컬럼 자동 추천 신뢰도가 낮습니다. 실행 전 미리보기 확인이 필요합니다.")
            if normalization == "loose_numeric":
                warnings.append("앞자리 0 등 표시 형식이 다른 키 값이 있을 수 있습니다.")

        value_columns = self._resolve_value_columns(
            reference,
            target,
            key_spec.reference_key_columns,
            preferred_value_columns,
        )
        if not value_columns:
            raise ValueError("기준표에서 가져올 컬럼을 자동으로 찾지 못했습니다.")
        for value_column in value_columns:
            evidence_rows.append(
                {
                    "역할": "가져올 컬럼",
                    "기준표 컬럼": value_column,
                    "대상표 컬럼": "",
                    "신뢰도": confidence,
                    "추천 방식": "대상표에 없는 기준표 컬럼 자동 선택",
                    "검토 메모": "키 컬럼과 대상표 기존 컬럼을 제외하고 결과에 붙일 값으로 선택했습니다.",
                }
            )

        return AutoLookupPlan(
            key_spec=key_spec,
            value_columns=value_columns,
            confidence=round(confidence, 4),
            evidence_rows=tuple(evidence_rows),
            warnings=tuple(warnings),
        )

    def infer_key_normalization(
        self,
        reference: pd.DataFrame,
        target: pd.DataFrame,
        reference_key_columns: tuple[str, ...],
        target_key_columns: tuple[str, ...],
    ) -> Literal["text", "loose_numeric"]:
        reference_text = build_key_series(reference, reference_key_columns)
        target_text = build_key_series(target, target_key_columns)
        reference_loose = build_key_series(reference, reference_key_columns, normalization="loose_numeric")
        target_loose = build_key_series(target, target_key_columns, normalization="loose_numeric")

        exact_overlap = len(set(reference_text[reference_text != ""]).intersection(set(target_text[target_text != ""])))
        loose_overlap = len(set(reference_loose[reference_loose != ""]).intersection(set(target_loose[target_loose != ""])))
        if loose_overlap > exact_overlap:
            return "loose_numeric"
        return "text"

    def infer_reconciliation_key_spec(
        self,
        reference: pd.DataFrame,
        target: pd.DataFrame,
        preferred_reference_key_columns: tuple[str, ...] = (),
        preferred_target_key_columns: tuple[str, ...] = (),
    ) -> AutoLookupPlan:
        plan = self.infer_lookup_plan(
            reference,
            target,
            preferred_reference_key_columns=preferred_reference_key_columns,
            preferred_target_key_columns=preferred_target_key_columns,
            preferred_value_columns=tuple(column for column in reference.columns if column not in preferred_reference_key_columns),
        )
        key_evidence = tuple(row for row in plan.evidence_rows if row.get("역할") == "키 컬럼")
        return AutoLookupPlan(
            key_spec=plan.key_spec,
            value_columns=(),
            confidence=plan.confidence,
            evidence_rows=key_evidence,
            warnings=plan.warnings,
        )

    def _resolve_target_key_columns(
        self,
        target: pd.DataFrame,
        reference_key_columns: tuple[str, ...],
        preferred_target_key_columns: tuple[str, ...],
    ) -> tuple[str, ...]:
        if preferred_target_key_columns and all(column in target.columns for column in preferred_target_key_columns):
            return preferred_target_key_columns
        if all(column in target.columns for column in reference_key_columns):
            return reference_key_columns
        mapping = self.mapper.map_columns(list(target.columns), list(reference_key_columns), threshold=0.70)
        resolved = tuple(mapping.target_to_source.get(column, column) for column in reference_key_columns)
        missing = [column for column in resolved if column not in target.columns]
        if missing:
            raise KeyError(f"대상표에서 키 컬럼을 자동 매칭하지 못했습니다: {missing}")
        return resolved

    def _resolve_value_columns(
        self,
        reference: pd.DataFrame,
        target: pd.DataFrame,
        reference_key_columns: tuple[str, ...],
        preferred_value_columns: tuple[str, ...],
    ) -> tuple[str, ...]:
        preferred = tuple(column for column in preferred_value_columns if column in reference.columns)
        if preferred:
            return preferred

        target_normalized = {normalize_header(column) for column in target.columns}
        key_normalized = {normalize_header(column) for column in reference_key_columns}
        values = []
        for column in reference.columns:
            normalized = normalize_header(column)
            if normalized in key_normalized:
                continue
            if normalized in target_normalized:
                continue
            values.append(column)
        if values:
            return tuple(values)
        return tuple(column for column in reference.columns if normalize_header(column) not in key_normalized)
