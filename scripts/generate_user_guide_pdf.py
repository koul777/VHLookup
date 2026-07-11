from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "docs" / "manuals" / "VHLookup_User_Guide.pdf"
MAIN_SCREENSHOT = ROOT / "docs" / "assets" / "vhlookup-local-main.jpg"
FEATURE_IMAGES = [
    ROOT / "docs" / "images" / "feature_01_privacy_masking.png",
    ROOT / "docs" / "images" / "feature_02_split_sheets.png",
    ROOT / "docs" / "images" / "feature_03_merge_files.png",
    ROOT / "docs" / "images" / "feature_04_before_after_validation.png",
    ROOT / "docs" / "images" / "feature_05_monthly_list.png",
    ROOT / "docs" / "images" / "feature_06_pivot_summary.png",
    ROOT / "docs" / "images" / "feature_07_custom_order_sort.png",
]

PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN_X = 46
MARGIN_TOP = 78
MARGIN_BOTTOM = 48
CONTENT_WIDTH = PAGE_WIDTH - (MARGIN_X * 2)
BLUE = colors.HexColor("#234F84")
GREEN = colors.HexColor("#317A47")
DARK = colors.HexColor("#1F2933")
MUTED = colors.HexColor("#5F6C7B")
LIGHT_LINE = colors.HexColor("#D8DEE9")
SOFT_BLUE = colors.HexColor("#EAF2FB")


def register_fonts() -> tuple[str, str]:
    regular = Path(r"C:\Windows\Fonts\malgun.ttf")
    bold = Path(r"C:\Windows\Fonts\malgunbd.ttf")
    if regular.exists() and bold.exists():
        pdfmetrics.registerFont(TTFont("MalgunGothic", str(regular)))
        pdfmetrics.registerFont(TTFont("MalgunGothic-Bold", str(bold)))
        return "MalgunGothic", "MalgunGothic-Bold"
    return "Helvetica", "Helvetica-Bold"


FONT, FONT_BOLD = register_fonts()


def file_sha256(path: Path) -> str:
    if not path.exists():
        return "배포 파일 생성 후 확인"
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def file_size(path: Path) -> str:
    if not path.exists():
        return "배포 파일 생성 후 확인"
    size = path.stat().st_size
    mb = size / (1024 * 1024)
    return f"{size:,} bytes ({mb:.1f} MB)"


def draw_wrapped(
    pdf: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_width: float,
    *,
    font: str = FONT,
    size: float = 10.5,
    leading: float = 15,
    color=colors.black,
) -> float:
    pdf.setFont(font, size)
    pdf.setFillColor(color)
    for line in wrap_text(text, max_width, font, size):
        pdf.drawString(x, y, line)
        y -= leading
    return y


def wrap_text(text: str, max_width: float, font: str, size: float) -> list[str]:
    wrapped: list[str] = []
    for paragraph in text.splitlines() or [""]:
        tokens = paragraph.split(" ")
        line = ""
        for token in tokens:
            candidates = [token]
            if pdfmetrics.stringWidth(token, font, size) > max_width:
                candidates = split_long_token(token, max_width, font, size)
            for candidate in candidates:
                proposed = candidate if not line else f"{line} {candidate}"
                if pdfmetrics.stringWidth(proposed, font, size) <= max_width:
                    line = proposed
                else:
                    if line:
                        wrapped.append(line)
                    line = candidate
        wrapped.append(line)
    return wrapped


def split_long_token(token: str, max_width: float, font: str, size: float) -> list[str]:
    chunks: list[str] = []
    current = ""
    for char in token:
        proposed = current + char
        if current and pdfmetrics.stringWidth(proposed, font, size) > max_width:
            chunks.append(current)
            current = char
        else:
            current = proposed
    if current:
        chunks.append(current)
    return chunks


class Manual:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.pdf = canvas.Canvas(str(path), pagesize=A4)
        self.page = 0
        self.y = PAGE_HEIGHT - MARGIN_TOP

    def new_page(self, title: str | None = None) -> None:
        if self.page:
            self.footer()
            self.pdf.showPage()
        self.page += 1
        self.y = PAGE_HEIGHT - MARGIN_TOP
        self.header()
        if title:
            self.section(title)

    def header(self) -> None:
        self.pdf.setFillColor(BLUE)
        self.pdf.rect(0, PAGE_HEIGHT - 30, PAGE_WIDTH, 30, fill=True, stroke=False)
        self.pdf.setFillColor(colors.white)
        self.pdf.setFont(FONT_BOLD, 10)
        self.pdf.drawString(MARGIN_X, PAGE_HEIGHT - 20, "VHLookup Local V1.4")
        self.pdf.drawRightString(PAGE_WIDTH - MARGIN_X, PAGE_HEIGHT - 20, "사용설명서")

    def footer(self) -> None:
        self.pdf.setStrokeColor(LIGHT_LINE)
        self.pdf.line(MARGIN_X, 34, PAGE_WIDTH - MARGIN_X, 34)
        self.pdf.setFillColor(MUTED)
        self.pdf.setFont(FONT, 8.5)
        self.pdf.drawString(MARGIN_X, 22, "VHLookup Local V1.4 사용설명서 | 2026-07-11")
        self.pdf.drawRightString(PAGE_WIDTH - MARGIN_X, 22, f"{self.page}")

    def ensure(self, height: float) -> None:
        if self.y - height < MARGIN_BOTTOM:
            self.new_page()

    def section(self, text: str) -> None:
        self.ensure(42)
        self.pdf.setFillColor(DARK)
        self.pdf.setFont(FONT_BOLD, 18)
        self.pdf.drawString(MARGIN_X, self.y, text)
        self.y -= 12
        self.pdf.setStrokeColor(BLUE)
        self.pdf.setLineWidth(1.2)
        self.pdf.line(MARGIN_X, self.y, PAGE_WIDTH - MARGIN_X, self.y)
        self.y -= 22

    def subsection(self, text: str) -> None:
        self.ensure(30)
        self.pdf.setFillColor(BLUE)
        self.pdf.setFont(FONT_BOLD, 13)
        self.pdf.drawString(MARGIN_X, self.y, text)
        self.y -= 18

    def paragraph(self, text: str, *, size: float = 10.5, color=DARK) -> None:
        lines = wrap_text(text, CONTENT_WIDTH, FONT, size)
        self.ensure(len(lines) * 15 + 8)
        self.y = draw_wrapped(
            self.pdf,
            text,
            MARGIN_X,
            self.y,
            CONTENT_WIDTH,
            font=FONT,
            size=size,
            leading=15,
            color=color,
        )
        self.y -= 5

    def bullets(self, items: Iterable[str], *, indent: float = 12) -> None:
        for item in items:
            lines = wrap_text(item, CONTENT_WIDTH - indent - 12, FONT, 10.2)
            self.ensure(len(lines) * 14 + 4)
            self.pdf.setFillColor(BLUE)
            self.pdf.setFont(FONT_BOLD, 10.2)
            self.pdf.drawString(MARGIN_X, self.y, "-")
            self.y = draw_wrapped(
                self.pdf,
                item,
                MARGIN_X + indent,
                self.y,
                CONTENT_WIDTH - indent,
                font=FONT,
                size=10.2,
                leading=14,
                color=DARK,
            )
            self.y -= 2
        self.y -= 4

    def callout(self, title: str, lines: Iterable[str]) -> None:
        wrapped_count = 1 + sum(
            len(wrap_text(line, CONTENT_WIDTH - 28, FONT, 9.8)) for line in lines
        )
        height = wrapped_count * 14 + 22
        self.ensure(height)
        self.pdf.setFillColor(SOFT_BLUE)
        self.pdf.roundRect(MARGIN_X, self.y - height + 10, CONTENT_WIDTH, height, 6, fill=True, stroke=False)
        self.pdf.setFillColor(BLUE)
        self.pdf.setFont(FONT_BOLD, 10.5)
        self.pdf.drawString(MARGIN_X + 12, self.y - 6, title)
        cursor = self.y - 23
        for line in lines:
            cursor = draw_wrapped(
                self.pdf,
                line,
                MARGIN_X + 14,
                cursor,
                CONTENT_WIDTH - 28,
                font=FONT,
                size=9.8,
                leading=13,
                color=DARK,
            )
        self.y -= height + 6

    def image(self, path: Path, *, max_height: float, caption: str | None = None) -> None:
        if not path.exists():
            return
        reader = ImageReader(str(path))
        image_width, image_height = reader.getSize()
        scale = min(CONTENT_WIDTH / image_width, max_height / image_height)
        draw_width = image_width * scale
        draw_height = image_height * scale
        total_height = draw_height + (18 if caption else 0) + 10
        self.ensure(total_height)
        x = MARGIN_X + (CONTENT_WIDTH - draw_width) / 2
        y = self.y - draw_height
        self.pdf.drawImage(reader, x, y, width=draw_width, height=draw_height, mask="auto")
        self.y = y - 8
        if caption:
            self.pdf.setFillColor(MUTED)
            self.pdf.setFont(FONT, 8.8)
            self.pdf.drawCentredString(PAGE_WIDTH / 2, self.y, caption)
            self.y -= 18

    def table(self, rows: list[list[str]], widths: list[float]) -> None:
        row_height = 29
        self.ensure(row_height * len(rows) + 10)
        x0 = MARGIN_X
        y0 = self.y
        for row_index, row in enumerate(rows):
            fill = BLUE if row_index == 0 else (colors.HexColor("#F7FAFC") if row_index % 2 else colors.white)
            self.pdf.setFillColor(fill)
            self.pdf.rect(x0, y0 - row_height, sum(widths), row_height, fill=True, stroke=False)
            self.pdf.setStrokeColor(LIGHT_LINE)
            self.pdf.rect(x0, y0 - row_height, sum(widths), row_height, fill=False, stroke=True)
            x = x0
            for col_index, value in enumerate(row):
                self.pdf.setFillColor(colors.white if row_index == 0 else DARK)
                self.pdf.setFont(FONT_BOLD if row_index == 0 else FONT, 8.8)
                lines = wrap_text(value, widths[col_index] - 10, FONT_BOLD if row_index == 0 else FONT, 8.8)
                text_y = y0 - 12
                for line in lines[:2]:
                    self.pdf.drawString(x + 5, text_y, line)
                    text_y -= 10
                x += widths[col_index]
            y0 -= row_height
        self.y = y0 - 12

    def save(self) -> None:
        self.footer()
        self.pdf.save()


def build_manual() -> None:
    zip_path = ROOT / "dist" / "VHLookupLocal_v1.4.zip"
    exe_path = ROOT / "dist" / "VHLookupLocal_v1.4.exe"
    manual = Manual(OUTPUT)

    manual.new_page()
    manual.pdf.setFillColor(DARK)
    manual.pdf.setFont(FONT_BOLD, 26)
    manual.pdf.drawString(MARGIN_X, manual.y, "VHLookup Local V1.4 사용설명서")
    manual.y -= 28
    manual.pdf.setFillColor(MUTED)
    manual.pdf.setFont(FONT, 12)
    manual.pdf.drawString(MARGIN_X, manual.y, "공공기관 취합 엑셀 업무를 로컬 PC에서 처리하는 Windows 실행 프로그램")
    manual.y -= 26
    manual.callout(
        "V1.4 핵심 변경",
        [
            "7번 기능으로 사용자 지정 순서 정렬을 추가했습니다.",
            "데이터 파일을 탑재한 뒤 기준 컬럼을 고르면, 해당 컬럼 안의 고유값을 자동으로 불러오고 사용자가 정한 순서대로 전체 행을 정렬합니다.",
        ],
    )
    manual.image(MAIN_SCREENSHOT, max_height=330, caption="프로그램 메인 화면 - VHLookup Local V1.4")
    manual.paragraph("최종 갱신일: 2026-07-11")

    manual.new_page("1. 다운로드 및 실행")
    manual.bullets(
        [
            "GitHub Releases에서 VHLookupLocal_v1.4.zip을 내려받습니다.",
            "압축을 풀면 나오는 VHLookupLocal_v1.4.exe를 더블클릭해 실행합니다.",
            "소스 ZIP을 받은 경우 dist 폴더와 exe가 없는 것이 정상입니다. 실행파일은 GitHub Release 첨부 파일로 배포합니다.",
            "프로그램은 업로드한 파일을 외부 서버로 보내지 않고 로컬 PC에서 처리합니다.",
        ]
    )
    manual.subsection("배포 파일 확인")
    manual.table(
        [
            ["파일", "크기", "SHA256"],
            ["VHLookupLocal_v1.4.zip", file_size(zip_path), file_sha256(zip_path)],
            ["VHLookupLocal_v1.4.exe", file_size(exe_path), file_sha256(exe_path)],
        ],
        [140, 145, CONTENT_WIDTH - 285],
    )
    manual.subsection("기본 사용 흐름")
    manual.bullets(
        [
            "메인 화면에서 필요한 메뉴 번호를 누릅니다.",
            "입력 파일을 선택하고, 화면에서 요구하는 기준 컬럼이나 옵션을 지정합니다.",
            "실행 전 미리보기와 안내 문구를 확인합니다.",
            "결과 파일 저장 위치를 지정하고 생성된 엑셀 파일을 확인합니다.",
        ]
    )

    manual.new_page("2. 주요 기능")
    feature_rows = [
        ["번호", "기능", "주로 쓰는 상황"],
        ["1", "개인정보 마스킹", "성명, 전화번호, 생년월일 등 민감정보 일부를 숨겨 공유용 파일을 만들 때"],
        ["2", "기관별 시트 분리", "하나의 취합표를 기관, 부서, 담당자 등 기준값별 시트로 나눌 때"],
        ["3", "다중 파일 병합", "여러 기관이 보낸 파일을 한 파일로 합칠 때"],
        ["4", "전후 데이터 검증", "이전 제출본과 수정 제출본의 변경 행을 확인할 때"],
        ["5", "월별 목록 생성", "시작일과 종료일 사이의 월별 업무 목록을 만들 때"],
        ["6", "피벗 요약", "기관, 월, 상태, 담당자 기준으로 건수와 금액을 빠르게 요약할 때"],
        ["7", "사용자 지정 순서 정렬", "직제순, 지역순, 우선순위처럼 사용자가 정한 순서대로 행을 정렬할 때"],
    ]
    manual.table(feature_rows, [36, 130, CONTENT_WIDTH - 166])
    manual.callout(
        "7번 기능의 기준",
        [
            "별도의 기준 컬럼을 새로 만들 필요가 없습니다.",
            "탑재한 엑셀 또는 CSV 안에 이미 있는 컬럼 중 하나를 선택하고, 그 컬럼에 들어 있는 값들의 순서만 지정합니다.",
        ],
    )

    manual.new_page("3. 1번부터 3번 기능")
    for title, description, image_path in [
        ("1. 개인정보 마스킹", "마스킹할 컬럼을 선택하고 이름, 전화번호, 주민번호 형태에 맞게 일부 값을 숨깁니다. 원본은 보존하고 결과 파일을 따로 저장합니다.", FEATURE_IMAGES[0]),
        ("2. 기관별 시트 분리", "기준 컬럼의 값별로 시트를 나눕니다. 기관명, 부서명, 담당자처럼 반복되는 값을 기준으로 취합 자료를 배포하기 좋습니다.", FEATURE_IMAGES[1]),
        ("3. 다중 파일 병합", "동일한 양식의 여러 파일을 선택해 하나의 엑셀로 합칩니다. 파일명 출처 컬럼을 남겨 어느 파일에서 온 행인지 확인할 수 있습니다.", FEATURE_IMAGES[2]),
    ]:
        manual.subsection(title)
        manual.paragraph(description)
        manual.image(image_path, max_height=115)

    manual.new_page("4. 4번부터 6번 기능")
    for title, description, image_path in [
        ("4. 전후 데이터 검증", "변경 전 파일과 변경 후 파일을 비교해 추가, 삭제, 수정된 행을 확인합니다. 제출본 검토나 정정 요청 확인에 사용합니다.", FEATURE_IMAGES[3]),
        ("5. 월별 목록 생성", "시작월과 종료월을 기준으로 월별 행을 자동 생성합니다. 반복 보고, 월별 점검, 월별 취합 대상 목록을 만들 때 사용합니다.", FEATURE_IMAGES[4]),
        ("6. 피벗 요약", "선택한 행/열/값 기준으로 요약표를 만듭니다. 별도 피벗 조작 없이 기관별 건수, 월별 금액, 상태별 합계 등을 확인할 수 있습니다.", FEATURE_IMAGES[5]),
    ]:
        manual.subsection(title)
        manual.paragraph(description)
        manual.image(image_path, max_height=115)

    manual.new_page("5. 7번 사용자 지정 순서 정렬")
    manual.paragraph(
        "사용자 지정 순서 정렬은 직제순만을 위한 고정 기능이 아닙니다. 사용자가 원하는 기준 컬럼을 고르고, 그 컬럼 안에 실제로 들어 있는 값들의 순서를 직접 지정해 전체 행을 정렬하는 기능입니다."
    )
    manual.bullets(
        [
            "7. 사용자 지정 순서 정렬을 누릅니다.",
            "정렬할 엑셀 또는 CSV 파일을 선택합니다.",
            "파일에서 읽어온 컬럼 목록 중 정렬 기준으로 쓸 컬럼을 고릅니다.",
            "선택한 컬럼에 들어 있는 고유값이 값 순서 목록에 자동으로 표시됩니다.",
            "맨 위, 위로, 아래로, 맨 아래 버튼으로 값을 원하는 순서로 바꿉니다.",
            "실행하면 지정한 순서에 맞춰 전체 행이 정렬되고, 지정하지 않은 값은 뒤쪽으로 이동합니다.",
            "결과 파일에는 정렬된 데이터와 사용자가 지정한 순서 정보가 함께 저장됩니다.",
        ]
    )
    manual.callout(
        "예시",
        [
            "부서 컬럼 값이 기획조정실, 행정지원과, 재무과, 복지정책과, 안전관리과라면 이 값을 원하는 직제순으로 재배치해 정렬할 수 있습니다.",
            "같은 방식으로 지역순, 업무 우선순위, 처리 단계순 등에도 사용할 수 있습니다.",
        ],
    )
    manual.image(FEATURE_IMAGES[6], max_height=250, caption="7번 사용자 지정 순서 정렬 전후 비교")

    manual.new_page("6. 메뉴별 샘플")
    sample_rows = [
        ["메뉴", "샘플 경로", "확인 포인트"],
        ["1", "samples/public_admin/01_privacy_masking/citizen_contacts.csv", "개인정보가 안전하게 가려지는지 확인"],
        ["2", "samples/public_admin/02_split_sheets/service_requests.csv", "기관별 시트가 생성되는지 확인"],
        ["3", "samples/public_admin/03_merge_files/*.csv", "여러 기관 파일이 하나로 병합되는지 확인"],
        ["4", "samples/public_admin/04_before_after_validation", "전후 변경 행이 분류되는지 확인"],
        ["5", "samples/public_admin/05_monthly_list/program_schedule.csv", "월별 반복 목록이 생성되는지 확인"],
        ["6", "samples/public_admin/06_pivot_summary/budget_execution.csv", "피벗 요약표가 생성되는지 확인"],
        ["7", "samples/public_admin/07_organization_order/department_tasks.csv", "선택 컬럼의 고유값을 지정한 순서대로 정렬하는지 확인"],
    ]
    manual.table(sample_rows, [42, 245, CONTENT_WIDTH - 287])
    manual.paragraph(
        "샘플 파일은 프로그램 기능 확인용입니다. 실제 업무 파일을 사용할 때도 동일하게 파일을 탑재하고 화면에서 컬럼을 선택해 실행하면 됩니다."
    )

    manual.new_page("7. 결과 파일 및 문제 해결")
    manual.subsection("결과 파일")
    manual.bullets(
        [
            "기본 결과는 엑셀 파일로 저장됩니다.",
            "입력 파일의 원본 행과 컬럼은 가능한 보존하고, 기능에 필요한 보조 시트나 검증 컬럼만 추가합니다.",
            "7번 기능 결과 파일에는 정렬된 데이터와 사용자가 지정한 값 순서가 남아 재검토가 가능합니다.",
        ]
    )
    manual.subsection("자주 막히는 상황")
    manual.bullets(
        [
            "파일이 열려 있으면 저장에 실패할 수 있습니다. 결과 파일을 닫고 다시 실행합니다.",
            "CSV의 한글이 깨지면 UTF-8 또는 CP949 인코딩으로 저장한 뒤 다시 불러옵니다.",
            "7번에서 원하는 컬럼이 보이지 않으면 첫 행이 헤더인지 확인합니다.",
            "정렬 후 일부 값이 뒤로 밀리면 값 순서 목록에 포함되지 않은 값이 있는지 확인합니다.",
            "보안 경고가 뜨면 GitHub Release에서 내려받은 파일명과 SHA256 값을 먼저 확인합니다.",
        ]
    )
    manual.callout(
        "배포 위치",
        [
            "다운로드: https://github.com/koul777/VHLookup/releases/latest/download/VHLookupLocal_v1.4.zip",
            "릴리즈 목록: https://github.com/koul777/VHLookup/releases/latest",
            "사용설명서 PDF: docs/manuals/VHLookup_User_Guide.pdf",
        ],
    )

    manual.save()


if __name__ == "__main__":
    build_manual()
    print(OUTPUT)
