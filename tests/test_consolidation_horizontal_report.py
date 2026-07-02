import pandas as pd

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
    assert sheets[SHEET_RESULT].loc[0, "성명"] == "홍길동"
    assert sheets[SHEET_GUIDE].loc[0, "확인 유형"] == "처리 결과"
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
