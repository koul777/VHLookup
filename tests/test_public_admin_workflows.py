from pathlib import Path

import pandas as pd

from vhlookup_core import AutoLookupPlanner, ExcelLoader, HeaderDetector, SheetDetector
from vhlookup_core.consolidation import ConsolidationEngine
from vhlookup_core.merge import MergeEngine
from vhlookup_core.models import KeySpec
from vhlookup_core.reconciliation import ReconciliationEngine
from vhlookup_core.templates import get_template, templates_for_mode


SAMPLES = Path("samples/public_admin")
MERGE_SAMPLES = SAMPLES / "03_merge_files"
HR_SAMPLES = MERGE_SAMPLES / "column_merge_hr_training"
ALLOWANCE_SAMPLES = MERGE_SAMPLES / "column_merge_allowance_budget"
EXTRA_SAMPLES = SAMPLES / "90_extra_cli_samples"


def _load_table(path: Path):
    loader = ExcelLoader()
    header_detector = HeaderDetector()
    sheet_detector = SheetDetector(header_detector)
    source = loader.load(path)
    sheet = sheet_detector.select(source)
    detection = header_detector.detect(sheet)
    return header_detector.apply(sheet, detection)


def test_public_admin_templates_cover_each_primary_mode():
    assert templates_for_mode("consolidation")
    assert templates_for_mode("lookup")
    assert templates_for_mode("reconciliation")
    assert templates_for_mode("horizontal")
    assert get_template("school_submission_consolidation").standard_columns


def test_school_submission_sample_consolidates_with_template():
    result = ConsolidationEngine().consolidate_folder(
        MERGE_SAMPLES / "row_merge_school_submissions",
        template="school_submission_consolidation",
    )

    assert len(result.result_frame) == 4
    assert list(result.result_frame.columns) == [
        "기관명",
        "기관코드",
        "담당자",
        "연락처",
        "제출일",
        "사업명",
        "항목",
        "금액",
        "비고",
    ]
    assert set(result.result_frame["기관명"]) == {"강북초", "강남중"}
    assert result.summary["workflow"] == "학교/부서 제출자료 한 번에 수합"
    assert result.mapping_records


def test_submission_error_sample_surfaces_admin_validation_issues():
    result = ConsolidationEngine().consolidate_folder(
        MERGE_SAMPLES / "row_merge_submission_errors",
        template="school_submission_consolidation",
    )

    issue_types = {issue.issue_type for issue in result.issues}
    assert {"required_value_missing", "numeric_value_invalid", "date_value_invalid", "duplicate_business_key"} <= issue_types
    assert result.summary["validation_issue_count"] >= 4


def test_training_lookup_sample_flags_numeric_id_format_mismatch():
    reference = _load_table(HR_SAMPLES / "hr_employee_master.csv")
    target = _load_table(HR_SAMPLES / "hr_training_completion.csv")

    result = MergeEngine().merge_lookup(
        reference,
        target,
        KeySpec(reference_key_columns=("사번",), target_key_columns=("직원번호",)),
        value_columns=["성명", "부서", "직급", "소속"],
    )

    issue_types = {issue.issue_type for issue in result.issues}
    assert "format_mismatch" in issue_types
    assert "match_failed" in issue_types
    assert result.result_frame.loc[1, "부서"] == "인사과"


def test_lookup_plan_auto_matches_columns_without_examples_or_manual_keys():
    reference = _load_table(HR_SAMPLES / "hr_employee_master.csv")
    target = _load_table(HR_SAMPLES / "hr_training_completion.csv")

    plan = AutoLookupPlanner().infer_lookup_plan(reference, target)

    assert plan.key_spec.reference_key_columns == ("사번",)
    assert plan.key_spec.target_columns() == ("직원번호",)
    assert plan.key_spec.normalization == "loose_numeric"
    assert set(plan.value_columns) == {"성명", "부서", "직급", "소속"}
    assert plan.confidence >= 0.55
    assert plan.evidence_rows

    result = MergeEngine().merge_lookup(
        reference,
        target,
        plan.key_spec,
        value_columns=list(plan.value_columns),
    )

    assert result.result_frame.loc[0, "부서"] == "총무과"
    assert result.result_frame.loc[1, "부서"] == "인사과"
    assert "match_failed" in {issue.issue_type for issue in result.issues}


def test_reconciliation_finds_missing_public_admin_rows():
    reference = _load_table(HR_SAMPLES / "hr_employee_master.csv")
    target = _load_table(HR_SAMPLES / "hr_training_completion.csv")

    result = ReconciliationEngine().compare_lists(
        reference,
        target,
        KeySpec(reference_key_columns=("사번",), target_key_columns=("직원번호",)),
    )

    assert result.summary["missing_in_target_rows"] >= 1
    assert "대조 상태" in result.result_frame.columns
    assert "missing_in_target" in {issue.issue_type for issue in result.issues}


def test_reconciliation_auto_plan_reports_key_evidence_only():
    reference = _load_table(HR_SAMPLES / "hr_employee_master.csv")
    target = _load_table(HR_SAMPLES / "hr_training_completion.csv")

    plan = AutoLookupPlanner().infer_reconciliation_key_spec(reference, target)

    assert plan.key_spec.reference_key_columns == ("사번",)
    assert plan.key_spec.target_columns() == ("직원번호",)
    assert plan.key_spec.normalization == "loose_numeric"
    assert plan.value_columns == ()
    assert {row["역할"] for row in plan.evidence_rows} == {"키 컬럼"}


def test_allowance_budget_sample_supports_composite_key_lookup():
    reference = _load_table(ALLOWANCE_SAMPLES / "rate_reference.csv")
    target = _load_table(ALLOWANCE_SAMPLES / "payment_requests.csv")

    plan = AutoLookupPlanner().infer_lookup_plan(
        reference,
        target,
        preferred_reference_key_columns=("사번", "지급월"),
        preferred_target_key_columns=("직원번호", "지급월"),
        preferred_value_columns=("단가", "지급기준", "예산과목", "비고"),
    )
    result = MergeEngine().merge_lookup(reference, target, plan.key_spec, list(plan.value_columns))

    assert plan.key_spec.reference_key_columns == ("사번", "지급월")
    assert plan.key_spec.target_columns() == ("직원번호", "지급월")
    assert plan.key_spec.normalization == "loose_numeric"
    assert result.result_frame.loc[0, "단가"] == "50000"
    assert result.result_frame.loc[1, "단가"] == "75000"
    assert {issue.issue_type for issue in result.issues} == {"match_failed"}


def test_column_merge_samples_use_loose_numeric_and_composite_keys():
    training = ConsolidationEngine().merge_files_by_columns(
        [HR_SAMPLES / "hr_training_completion.csv", HR_SAMPLES / "hr_employee_master.csv"]
    )
    allowance = ConsolidationEngine().merge_files_by_columns(
        [ALLOWANCE_SAMPLES / "payment_requests.csv", ALLOWANCE_SAMPLES / "rate_reference.csv"]
    )

    assert training.result_frame.loc[0, "부서"] == "총무과"
    assert training.result_frame.loc[1, "부서"] == "인사과"
    assert allowance.result_frame.loc[0, "단가"] == "50000"
    assert allowance.result_frame.loc[1, "단가"] == "75000"
    assert {
        (row["기준표 컬럼"], row["대상표 컬럼"])
        for row in allowance.mapping_records
        if row.get("역할") == "키 컬럼"
    } == {("사번", "직원번호"), ("지급월", "지급월")}


def test_merge_mode_recommendation_distinguishes_row_and_column_samples():
    engine = ConsolidationEngine()

    assert engine.recommend_merge_mode(
        [
            MERGE_SAMPLES / "row_merge_school_submissions" / "gangbuk_school.csv",
            MERGE_SAMPLES / "row_merge_school_submissions" / "gangnam_school.csv",
        ]
    ) == "rows"
    assert engine.recommend_merge_mode(
        [
            MERGE_SAMPLES / "row_merge_messy_headers" / "department_status_a.csv",
            MERGE_SAMPLES / "row_merge_messy_headers" / "department_status_b.csv",
        ]
    ) == "rows"
    assert engine.recommend_merge_mode(
        [
            HR_SAMPLES / "hr_employee_master.csv",
            HR_SAMPLES / "hr_training_completion.csv",
        ]
    ) == "columns"
    assert engine.recommend_merge_mode(
        [
            ALLOWANCE_SAMPLES / "payment_requests.csv",
            ALLOWANCE_SAMPLES / "rate_reference.csv",
        ]
    ) == "columns"


def test_submission_reconciliation_sample_finds_missing_and_unknown_submitters():
    expected = _load_table(EXTRA_SAMPLES / "submission_reconciliation" / "expected_submitters.csv")
    received = _load_table(EXTRA_SAMPLES / "submission_reconciliation" / "received_submitters.csv")

    plan = AutoLookupPlanner().infer_reconciliation_key_spec(
        expected,
        received,
        preferred_reference_key_columns=("기관코드",),
        preferred_target_key_columns=("제출기관코드",),
    )
    result = ReconciliationEngine().compare_lists(
        expected,
        received,
        plan.key_spec,
        reference_label="제출 대상 명단",
        target_label="실제 제출 명단",
    )

    assert result.summary["missing_in_target_rows"] == 2
    assert result.summary["missing_in_reference_rows"] == 1


def test_messy_header_department_status_sample_consolidates():
    result = ConsolidationEngine().consolidate_folder(
        MERGE_SAMPLES / "row_merge_messy_headers",
        template="department_status_consolidation",
    )

    assert len(result.result_frame) == 4
    assert list(result.result_frame.columns) == ["부서", "담당자", "연락처", "기준일", "현원", "비고"]
    assert set(result.result_frame["부서"]) == {"총무과", "인사과", "예산과", "기획과"}
    assert not [issue for issue in result.issues if issue.severity == "error"]
