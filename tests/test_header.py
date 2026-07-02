import pandas as pd

from vhlookup_core.header import HeaderDetector
from vhlookup_core.models import SheetData


def test_header_detector_skips_title_and_guidance_rows():
    frame = pd.DataFrame(
        [
            ["직원 교육 이수 현황", None, None],
            ["작성 안내: 아래 양식에 맞게 입력", None, None],
            [None, None, None],
            ["사번", "성명", "부서"],
            ["001", "홍길동", "총무"],
            ["002", "김영희", "인사"],
        ]
    )
    detection = HeaderDetector().detect(SheetData("Sheet1", frame))

    assert detection.header_row_number == 4
    assert detection.headers == ("사번", "성명", "부서")
    assert detection.confidence >= 0.75

    table = HeaderDetector().apply(SheetData("Sheet1", frame), detection)
    assert list(table.columns) == ["사번", "성명", "부서"]
    assert table.iloc[0].to_dict() == {"사번": "001", "성명": "홍길동", "부서": "총무"}


def test_header_detector_combines_two_line_headers():
    frame = pd.DataFrame(
        [
            ["부서별 현황", None, None],
            ["기본", "기본", "근무"],
            ["사번", "성명", "부서"],
            ["001", "홍길동", "총무"],
        ]
    )
    detection = HeaderDetector().detect(SheetData("Sheet1", frame))

    assert detection.header_row_number == 2
    assert detection.header_row_count == 2
    assert detection.headers == ("기본_사번", "기본_성명", "근무_부서")
