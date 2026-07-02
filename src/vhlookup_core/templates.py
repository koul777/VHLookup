from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


WorkflowMode = Literal["consolidation", "lookup", "reconciliation", "horizontal"]


@dataclass(frozen=True)
class WorkflowTemplate:
    id: str
    name: str
    mode: WorkflowMode
    headline: str
    description: str
    standard_columns: tuple[str, ...] = ()
    key_columns: tuple[str, ...] = ()
    target_key_columns: tuple[str, ...] = ()
    value_columns: tuple[str, ...] = ()
    required_columns: tuple[str, ...] = ()
    numeric_columns: tuple[str, ...] = ()
    date_columns: tuple[str, ...] = ()
    duplicate_key_columns: tuple[str, ...] = ()
    output_name: str = "vhlookup_result.xlsx"
    safety_note: str = "원본 파일은 수정하지 않고 결과 파일만 새로 만듭니다."


TEMPLATES: tuple[WorkflowTemplate, ...] = (
    WorkflowTemplate(
        id="school_submission_consolidation",
        name="학교/부서 제출자료 한 번에 수합",
        mode="consolidation",
        headline="폴더 하나 선택해서 제출자료 통합본 만들기",
        description="학교, 부서, 산하기관에서 받은 양식이 조금 달라도 표준 컬럼으로 맞춰 합칩니다.",
        standard_columns=(
            "기관명",
            "기관코드",
            "담당자",
            "연락처",
            "제출일",
            "사업명",
            "항목",
            "금액",
            "비고",
        ),
        required_columns=("기관명", "기관코드", "담당자", "제출일", "사업명", "항목", "금액"),
        numeric_columns=("금액",),
        date_columns=("제출일",),
        duplicate_key_columns=("기관코드", "사업명", "항목"),
        output_name="제출자료_수합결과.xlsx",
    ),
    WorkflowTemplate(
        id="department_status_consolidation",
        name="부서별 현황자료 수합",
        mode="consolidation",
        headline="부서별 현황 제출자료를 하나로 통합",
        description="제목과 안내문이 붙은 부서별 현황 파일에서 실제 헤더를 찾고 현원 자료를 합칩니다.",
        standard_columns=("부서", "담당자", "연락처", "기준일", "현원", "비고"),
        required_columns=("부서", "담당자", "기준일", "현원"),
        numeric_columns=("현원",),
        date_columns=("기준일",),
        duplicate_key_columns=("부서", "기준일"),
        output_name="부서별현황_수합결과.xlsx",
    ),
    WorkflowTemplate(
        id="hr_training_lookup",
        name="직원 교육이수 명단에 부서/직급 붙이기",
        mode="lookup",
        headline="교육이수 명단을 직원 기본명단과 자동 대조",
        description="사번 기준으로 부서, 직급, 소속 정보를 붙이고 누락자와 형식 오류를 따로 표시합니다.",
        key_columns=("사번",),
        target_key_columns=("직원번호",),
        value_columns=("성명", "부서", "직급", "소속"),
        required_columns=("사번",),
        duplicate_key_columns=("사번",),
        output_name="교육이수_명단대조.xlsx",
    ),
    WorkflowTemplate(
        id="allowance_budget_lookup",
        name="수당/예산 지급대상자 기준표 대조",
        mode="lookup",
        headline="지급대상자 명단에 기준 단가와 예산과목 붙이기",
        description="사번과 지급월 같은 복합 키로 기준표를 붙이고 중복 지급 위험을 먼저 보여줍니다.",
        key_columns=("사번", "지급월"),
        target_key_columns=("사번", "지급월"),
        value_columns=("단가", "지급기준", "예산과목", "비고"),
        required_columns=("사번", "지급월"),
        numeric_columns=("단가",),
        duplicate_key_columns=("사번", "지급월"),
        output_name="수당예산_대조결과.xlsx",
    ),
    WorkflowTemplate(
        id="missing_people_reconciliation",
        name="빠진 사람/누락 제출자 찾기",
        mode="reconciliation",
        headline="두 명단을 비교해 어느 쪽에 빠졌는지 표시",
        description="직원, 제출자, 교육대상자, 지급대상자 명단을 비교해 기준표에만 있음/대상표에만 있음을 분리합니다.",
        key_columns=("사번",),
        required_columns=("사번",),
        duplicate_key_columns=("사번",),
        output_name="누락자_대조결과.xlsx",
    ),
    WorkflowTemplate(
        id="monthly_wide_to_long",
        name="월별 가로표 세로 변환",
        mode="horizontal",
        headline="1월, 2월, 3월로 옆으로 긴 표를 세로형 자료로 변환",
        description="월별, 분기별, 항목별 가로표를 수합과 대조에 다시 쓸 수 있는 세로형 표로 바꿉니다.",
        key_columns=("기관명", "항목"),
        required_columns=("기관명", "항목"),
        output_name="월별가로표_세로변환.xlsx",
    ),
)


def all_templates() -> tuple[WorkflowTemplate, ...]:
    return TEMPLATES


def templates_for_mode(mode: WorkflowMode) -> tuple[WorkflowTemplate, ...]:
    return tuple(template for template in TEMPLATES if template.mode == mode)


def get_template(template_id: str | None) -> WorkflowTemplate | None:
    if not template_id:
        return None
    for template in TEMPLATES:
        if template.id == template_id:
            return template
    raise KeyError(f"Unknown workflow template: {template_id}")


def default_template(mode: WorkflowMode) -> WorkflowTemplate:
    templates = templates_for_mode(mode)
    if not templates:
        raise KeyError(f"No templates for mode: {mode}")
    return templates[0]
