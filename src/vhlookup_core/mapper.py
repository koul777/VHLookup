from __future__ import annotations

from difflib import SequenceMatcher

from vhlookup_core.models import ColumnMapping
from vhlookup_core.normalization import normalize_header


SYNONYMS: dict[str, tuple[str, ...]] = {
    "사번": ("사번", "직원번호", "직원id", "employeeid", "empid", "id"),
    "성명": ("성명", "이름", "직원명", "담당자명", "대상자명", "신청자명", "name", "employee"),
    "부서": ("부서", "소속", "팀", "부서명", "담당부서", "department", "dept"),
    "직급": ("직급", "직위", "position", "rank", "grade"),
    "일자": ("일자", "날짜", "기준일", "작성일", "제출일", "date", "day"),
    "금액": ("금액", "금원", "지급액", "교부액", "집행액", "신청액", "amount", "price", "cost"),
    "학교코드": ("학교코드", "기관코드", "schoolcode", "orgcode"),
    "기관코드": ("기관코드", "제출기관코드", "학교코드", "부서코드", "소속코드", "orgcode", "agencycode"),
    "학교명": ("학교명", "기관명", "학교", "부서명", "제출기관", "schoolname", "orgname"),
    "기관명": ("기관명", "학교명", "부서명", "소속기관", "제출기관", "orgname", "agency"),
    "연락처": ("연락처", "전화번호", "휴대폰", "내선번호", "phone", "mobile"),
    "이메일": ("이메일", "메일", "email", "e-mail"),
    "담당자": ("담당자", "담당자명", "작성자", "담당", "manager", "owner"),
    "사업명": ("사업명", "사업", "프로그램명", "과제명", "project", "program"),
    "항목": ("항목", "세부항목", "구분", "내역", "item", "category"),
    "지급월": ("지급월", "월", "귀속월", "기준월", "month"),
    "단가": ("단가", "기준단가", "지급단가", "unitprice"),
    "예산과목": ("예산과목", "세목", "목", "세부사업", "budgetcode", "budgetitem"),
    "이수여부": ("이수여부", "수료여부", "완료여부", "completion", "completed"),
    "교육명": ("교육명", "과정명", "교육과정", "training", "course"),
    "비고": ("비고", "메모", "참고", "특이사항", "remark", "note"),
    "현원": ("현원", "인원수", "인원", "정원", "현재인원", "headcount"),
}


class ColumnMapper:
    def map_columns(
        self,
        source_headers: list[str] | tuple[str, ...],
        target_headers: list[str] | tuple[str, ...],
        saved_mappings: dict[str, str] | None = None,
        threshold: float = 0.72,
    ) -> ColumnMapping:
        saved_mappings = saved_mappings or {}
        source_to_target: dict[str, str] = {}
        target_to_source: dict[str, str] = {}
        confidence_by_target: dict[str, float] = {}
        warnings: list[str] = []

        available_sources = set(source_headers)
        for target in target_headers:
            saved_source = saved_mappings.get(target)
            if saved_source in available_sources:
                self._assign(source_to_target, target_to_source, confidence_by_target, saved_source, target, 1.0)
                available_sources.remove(saved_source)
                continue

            best_source = None
            best_score = 0.0
            for source in available_sources:
                score = self._score(source, target)
                if score > best_score:
                    best_score = score
                    best_source = source
            if best_source is not None and best_score >= threshold:
                self._assign(
                    source_to_target,
                    target_to_source,
                    confidence_by_target,
                    best_source,
                    target,
                    best_score,
                )
                available_sources.remove(best_source)

        unmapped = tuple(target for target in target_headers if target not in target_to_source)
        if unmapped:
            warnings.append("일부 표준 컬럼을 자동 매칭하지 못했습니다.")
        return ColumnMapping(
            source_to_target=source_to_target,
            target_to_source=target_to_source,
            confidence_by_target=confidence_by_target,
            unmapped_targets=unmapped,
            warnings=tuple(warnings),
        )

    def _assign(
        self,
        source_to_target: dict[str, str],
        target_to_source: dict[str, str],
        confidence_by_target: dict[str, float],
        source: str,
        target: str,
        score: float,
    ) -> None:
        source_to_target[source] = target
        target_to_source[target] = source
        confidence_by_target[target] = round(score, 4)

    def _score(self, source: str, target: str) -> float:
        source_norm = normalize_header(source)
        target_norm = normalize_header(target)
        if source_norm == target_norm:
            return 1.0
        if self._same_synonym_group(source_norm, target_norm):
            return 0.94
        if self._contains_synonym_for_target(source_norm, target_norm):
            return 0.90
        if source_norm in target_norm or target_norm in source_norm:
            return 0.86
        return SequenceMatcher(None, source_norm, target_norm).ratio()

    def _same_synonym_group(self, source_norm: str, target_norm: str) -> bool:
        for aliases in SYNONYMS.values():
            normalized = {normalize_header(alias) for alias in aliases}
            if source_norm in normalized and target_norm in normalized:
                return True
        return False

    def _contains_synonym_for_target(self, source_norm: str, target_norm: str) -> bool:
        for aliases in SYNONYMS.values():
            normalized = {normalize_header(alias) for alias in aliases}
            if target_norm in normalized and any(alias in source_norm for alias in normalized):
                return True
        return False
