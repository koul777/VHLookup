from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from vhlookup_core.consolidation import ConsolidationEngine
from vhlookup_core.horizontal import HorizontalTableEngine
from vhlookup_core.models import JobResult
from vhlookup_core.report import (
    ReportWriter,
    SHEET_AUTO_EVIDENCE,
    SHEET_GUIDE,
    SHEET_ISSUES,
    SHEET_PRIVACY,
    SHEET_RESULT,
    SHEET_SUMMARY,
)
from vhlookup_core.privacy import PrivacyScanner


SAMPLES = Path("samples/public_admin")


def test_consolidation_folder_merges_files_with_different_headers(tmp_path):
    file_one = tmp_path / "one.xlsx"
    file_two = tmp_path / "two.xlsx"
    pd.DataFrame(
        [
            ["직원 현황", None, None],
            ["사번", "성명", "부서"],
            ["001", "홍길동", "총무"],
        ]
    ).to_excel(file_one, header=False, index=False)
    pd.DataFrame(
        [
            ["제출자료", None, None],
            ["직원번호", "이름", "소속"],
            ["002", "김영희", "인사"],
        ]
    ).to_excel(file_two, header=False, index=False)

    result = ConsolidationEngine().consolidate_folder(
        tmp_path,
        standard_columns=["사번", "성명", "부서"],
    )

    assert len(result.result_frame) == 2
    assert list(result.result_frame.columns) == ["원본 파일명", "원본 시트명", "원본 행 번호", "사번", "성명", "부서"]
    assert set(result.result_frame["성명"]) == {"홍길동", "김영희"}


def test_consolidation_preserves_unmapped_source_columns(tmp_path):
    file_one = tmp_path / "one.xlsx"
    file_two = tmp_path / "two.xlsx"
    pd.DataFrame(
        [
            ["직원 현황", None, None, None],
            ["사번", "성명", "부서", "첨부파일"],
            ["001", "홍길동", "총무", "있음"],
        ]
    ).to_excel(file_one, header=False, index=False)
    pd.DataFrame(
        [
            ["제출자료", None, None, None],
            ["직원번호", "이름", "소속", "비고2"],
            ["002", "김영희", "인사", "추가 확인"],
        ]
    ).to_excel(file_two, header=False, index=False)

    result = ConsolidationEngine().consolidate_folder(
        tmp_path,
        standard_columns=["사번", "성명", "부서"],
    )

    assert {"첨부파일", "비고2"} <= set(result.result_frame.columns)
    assert result.result_frame.loc[0, "첨부파일"] == "있음"
    assert result.result_frame.loc[1, "비고2"] == "추가 확인"
    assert "원본 컬럼 보존" in {row["역할"] for row in result.mapping_records}


def test_consolidation_accepts_user_corrected_column_mapping(tmp_path):
    file_one = tmp_path / "one.xlsx"
    file_two = tmp_path / "two.xlsx"
    pd.DataFrame({"사번": ["001"], "성명": ["홍길동"], "부서": ["총무"]}).to_excel(file_one, index=False)
    pd.DataFrame({"직원번호": ["002"], "이름": ["김영희"], "팀명": ["인사"]}).to_excel(file_two, index=False)

    result = ConsolidationEngine().consolidate_files(
        [file_one, file_two],
        standard_columns=["사번", "성명", "부서"],
        saved_mappings_by_file={str(file_two.resolve()): {"사번": "직원번호", "성명": "이름", "부서": "팀명"}},
    )

    assert result.result_frame.loc[1, "사번"] == "002"
    assert result.result_frame.loc[1, "성명"] == "김영희"
    assert result.result_frame.loc[1, "부서"] == "인사"


def test_column_merge_attaches_columns_by_common_key(tmp_path):
    file_one = tmp_path / "base.xlsx"
    file_two = tmp_path / "extra.xlsx"
    pd.DataFrame(
        {
            "사번": ["001", "002"],
            "성명": ["홍길동", "김영희"],
        }
    ).to_excel(file_one, index=False)
    pd.DataFrame(
        {
            "사번": ["001", "002"],
            "직급": ["주무관", "팀장"],
            "부서": ["총무", "인사"],
        }
    ).to_excel(file_two, index=False)

    result = ConsolidationEngine().merge_files_by_columns([file_one, file_two])

    assert len(result.result_frame) == 2
    assert {"사번", "성명", "직급", "부서"} <= set(result.result_frame.columns)
    assert result.result_frame.loc[0, "직급"] == "주무관"
    assert result.summary["workflow"] == "열 방향 파일 합치기"
    assert "키 컬럼" in {row["역할"] for row in result.mapping_records}


def test_column_merge_marks_unmatched_attached_cells_with_comment(tmp_path):
    output = tmp_path / "result.xlsx"
    file_one = SAMPLES / "hr_employee_master.csv"
    file_two = SAMPLES / "hr_training_completion.csv"

    result = ConsolidationEngine().merge_files_by_columns([file_one, file_two])
    ReportWriter().write_xlsx(result, output)

    assert len(result.result_frame) == 3
    assert result.result_frame.loc[2, "사번"] == "00125"
    assert pd.isna(result.result_frame.loc[2, "교육명"])
    issue = next(issue for issue in result.issues if issue.issue_type == "match_failed")
    assert issue.message == "2번째 파일(hr_training_completion.csv)에 해당 기준값이 없습니다."

    workbook = load_workbook(output)
    result_sheet = workbook[SHEET_RESULT]
    headers = [cell.value for cell in result_sheet[1]]
    education_column = headers.index("교육명") + 1
    base_column = headers.index("직급") + 1
    missing_cell = result_sheet.cell(row=4, column=education_column)
    base_cell = result_sheet.cell(row=4, column=base_column)

    assert missing_cell.value is None
    assert missing_cell.fill.fgColor.rgb in {"00FCA5A5", "FCA5A5"}
    assert missing_cell.comment is not None
    assert "2번째 파일(hr_training_completion.csv)에 해당 기준값이 없습니다." in missing_cell.comment.text
    assert base_cell.comment is None


def test_report_writer_extracts_error_rows(tmp_path):
    output = tmp_path / "result.xlsx"
    result = ConsolidationEngine().consolidate_folder(
        "samples/public_admin/submission_errors",
        template="school_submission_consolidation",
    )

    ReportWriter().write_xlsx(result, output)
    sheets = pd.read_excel(output, sheet_name=None)

    assert "오류행만" in sheets
    assert not sheets["오류행만"].empty
    assert {"확인 유형", "안내 문구", "원본 행 번호"} <= set(sheets["오류행만"].columns)


def test_report_writer_can_leave_merge_result_unmarked(tmp_path):
    output = tmp_path / "result.xlsx"
    result = ConsolidationEngine().consolidate_folder(
        "samples/public_admin/submission_errors",
        template="school_submission_consolidation",
    )

    ReportWriter().write_xlsx(result, output, mark_result_cells=False)
    workbook = load_workbook(output)
    result_sheet = workbook[SHEET_RESULT]

    assert workbook.sheetnames[0] == SHEET_RESULT
    assert all(cell.comment is None for row in result_sheet.iter_rows() for cell in row)


def test_horizontal_table_engine_detects_and_converts_month_columns():
    frame = pd.DataFrame(
        {
            "사번": ["001", "002"],
            "성명": ["홍길동", "김영희"],
            "1월": [10, 20],
            "2월": [11, 21],
            "3월": [12, 22],
        }
    )
    engine = HorizontalTableEngine()

    detection = engine.detect(frame)
    converted = engine.wide_to_long(frame, id_columns=["사번", "성명"])

    assert detection.is_horizontal
    assert set(detection.value_columns) == {"1월", "2월", "3월"}
    assert len(converted) == 6
    assert set(converted.columns) == {"사번", "성명", "열 기준", "값"}


def test_report_writer_creates_expected_sheets(tmp_path):
    output = tmp_path / "result.xlsx"
    result = JobResult(
        result_frame=pd.DataFrame({"성명": ["홍길동"]}),
        summary={"row_count": 1},
        mapping_records=[
            {
                "역할": "키 컬럼",
                "기준표 컬럼": "사번",
                "대상표 컬럼": "직원번호",
                "신뢰도": 0.9,
                "추천 방식": "테스트",
                "검토 메모": "",
            }
        ],
    )

    ReportWriter().write_xlsx(result, output)
    sheets = pd.read_excel(output, sheet_name=None)

    assert set(sheets) == {SHEET_GUIDE, SHEET_RESULT, SHEET_ISSUES, SHEET_SUMMARY, SHEET_AUTO_EVIDENCE, SHEET_PRIVACY}
    assert list(sheets)[0] == SHEET_RESULT
    assert sheets[SHEET_RESULT].loc[0, "성명"] == "홍길동"
    assert sheets[SHEET_GUIDE].loc[0, "먼저 볼 내용"] == "업무 결과"
    assert sheets[SHEET_AUTO_EVIDENCE].loc[0, "역할"] == "키 컬럼"
    assert sheets[SHEET_PRIVACY].loc[0, "컬럼명"] == "성명"


def test_privacy_scanner_flags_sensitive_columns_without_values():
    frame = pd.DataFrame(
        {
            "성명": ["홍길동"],
            "연락처": ["010-1234-5678"],
            "비고": ["주민번호 900101-1234567 확인"],
        }
    )

    records = PrivacyScanner().scan_frame(frame)

    assert {record["점검 유형"] for record in records} >= {"개인 식별 이름", "연락처", "주민등록번호 패턴"}
    serialized = str(records)
    assert "홍길동" not in serialized
    assert "010-1234-5678" not in serialized
    assert "900101-1234567" not in serialized
