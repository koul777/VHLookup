from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from vhlookup_cli.main import main
from vhlookup_core.inspection import InspectionEngine
from vhlookup_core.inspection_report import InspectionReportWriter


SAMPLES = Path("samples/public_admin")
MERGE_SAMPLES = SAMPLES / "03_merge_files"
HR_SAMPLES = MERGE_SAMPLES / "column_merge_hr_training"
EXTRA_SAMPLES = SAMPLES / "90_extra_cli_samples"


def test_cli_consolidate_generates_review_ready_workbook(tmp_path):
    output = tmp_path / "consolidated.xlsx"

    main(
        [
            "consolidate",
            "--folder",
            str(MERGE_SAMPLES / "row_merge_school_submissions"),
            "--template",
            "school_submission_consolidation",
            "--out",
            str(output),
        ]
    )

    sheets = pd.read_excel(output, sheet_name=None)
    assert set(sheets) == {"결과", "확인사항"}
    assert len(sheets["결과"]) == 4
    privacy_rows = sheets["확인사항"][sheets["확인사항"]["구분"] == "개인정보 의심"]
    assert privacy_rows.empty


def test_cli_templates_exports_catalog(tmp_path):
    output = tmp_path / "templates.xlsx"

    main(["templates", "--out", str(output)])

    sheets = pd.read_excel(output, sheet_name=None)
    assert set(sheets) == {"업무템플릿", "명령예시"}
    assert "school_submission_consolidation" in set(sheets["업무템플릿"]["템플릿 ID"])
    assert "추천 명령" in sheets["명령예시"].columns


def test_cli_inspect_generates_preflight_workbook(tmp_path):
    output = tmp_path / "inspect.xlsx"

    main(
        [
            "inspect",
            "--path",
            str(MERGE_SAMPLES / "row_merge_school_submissions"),
            "--template",
            "school_submission_consolidation",
            "--out",
            str(output),
        ]
    )

    sheets = pd.read_excel(output, sheet_name=None)
    assert set(sheets) >= {
        "먼저확인",
        "파일점검",
        "시트헤더점검",
        "열비교",
        "컬럼요약",
        "컬럼매칭점검",
        "자동추천근거",
        "확인필요",
        "점검요약",
    }
    assert len(sheets["파일점검"]) == 2
    assert "헤더 행" in sheets["시트헤더점검"].columns
    assert "표준 컬럼" in sheets["컬럼매칭점검"].columns


def test_inspection_compares_selected_files_and_flags_sensitive_columns(tmp_path):
    file_a = tmp_path / "dept_a.xlsx"
    file_b = tmp_path / "dept_b.xlsx"
    output = tmp_path / "preflight.xlsx"

    pd.DataFrame(
        {
            "직원명": ["홍길동", "김영희"],
            "부서": ["총무", "회계"],
            "연락처": ["010-1234-5678", "010-2222-3333"],
        }
    ).to_excel(file_a, index=False)
    pd.DataFrame(
        {
            "직원명": ["박철수"],
            "부서": ["민원"],
            "예산액": [100000],
        }
    ).to_excel(file_b, index=False)

    result = InspectionEngine().inspect_files([file_a, file_b])
    InspectionReportWriter().write_xlsx(result, output)
    sheets = pd.read_excel(output, sheet_name=None)
    workbook = load_workbook(output)
    source_sheet = workbook[f"원본확인_{file_a.stem}"]

    assert any(issue.issue_type == "column_schema_mismatch" for issue in result.issues)
    assert list(sheets)[0] == f"원본확인_{file_a.stem}"
    assert "열비교" in sheets
    assert "개인정보점검" in sheets
    assert file_b.name in set(result.column_comparison["누락 파일"])
    assert {"직원명", "연락처"} <= set(sheets["개인정보점검"]["컬럼명"])
    assert source_sheet["C1"].comment is not None
    assert source_sheet["C2"].fill.fgColor.rgb in {"00FDE68A", "FDE68A"}


def test_cli_lookup_auto_matches_without_manual_columns(tmp_path):
    output = tmp_path / "lookup.xlsx"

    main(
        [
            "lookup",
            "--reference",
            str(HR_SAMPLES / "hr_employee_master.csv"),
            "--target",
            str(HR_SAMPLES / "hr_training_completion.csv"),
            "--out",
            str(output),
        ]
    )

    sheets = pd.read_excel(output, sheet_name=None)
    review = sheets["확인사항"]
    assert "자동 선택 기준표 키" in set(review["기준/컬럼"])
    assert review.loc[review["구분"] == "키 컬럼", "기준/컬럼"].iloc[0] == "사번 -> 직원번호"


def test_cli_reconcile_and_horizontal_workflows(tmp_path):
    reconcile_out = tmp_path / "reconcile.xlsx"
    horizontal_out = tmp_path / "horizontal.xlsx"

    main(
        [
            "reconcile",
            "--reference",
            str(EXTRA_SAMPLES / "submission_reconciliation" / "expected_submitters.csv"),
            "--target",
            str(EXTRA_SAMPLES / "submission_reconciliation" / "received_submitters.csv"),
            "--key",
            "기관코드",
            "--target-key",
            "제출기관코드",
            "--out",
            str(reconcile_out),
        ]
    )
    main(
        [
            "horizontal",
            "--file",
            str(EXTRA_SAMPLES / "horizontal_table" / "monthly_budget_wide.csv"),
            "--out",
            str(horizontal_out),
        ]
    )

    reconcile = pd.read_excel(reconcile_out, sheet_name="확인사항")
    horizontal = pd.read_excel(horizontal_out, sheet_name="결과")
    assert "대상표에 없는 건수" in set(reconcile["기준/컬럼"])
    assert set(horizontal.columns) == {"기관명", "항목", "비고", "열 기준", "값"}
