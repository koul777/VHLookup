from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
DEMO_OUTPUT = ROOT / "demo_output"
IMAGE_OUTPUT = ROOT / "docs" / "images"

FEATURE_SHOTS = [
    (
        "01_개인정보_마스킹결과.xlsx",
        "마스킹결과",
        "1. 개인정보 마스킹",
        "feature_01_privacy_masking.png",
    ),
    (
        "02_분류별_시트나누기.xlsx",
        "전체",
        "2. 분류별 시트 나누기",
        "feature_02_split_sheets.png",
    ),
    (
        "03_제출자료_수합결과.xlsx",
        "결과",
        "3. 엑셀/CSV 파일 여러 개 합치기",
        "feature_03_merge_files.png",
    ),
    (
        "04_전후파일_검증결과.xlsx",
        "후파일_메모",
        "4. 전/후 파일 검증",
        "feature_04_before_after_validation.png",
    ),
    (
        "05_월별표_목록형변환.xlsx",
        "결과",
        "5. 월별표 목록형 변환",
        "feature_05_monthly_list.png",
    ),
    (
        "06_피벗요약표_결과.xlsx",
        "피벗요약",
        "6. 피벗 요약표 만들기",
        "feature_06_pivot_summary.png",
    ),
]


def _font(name: str, size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path("C:/Windows/Fonts") / name,
        Path("C:/Windows/Fonts/malgun.ttf"),
        Path("C:/Windows/Fonts/malgunbd.ttf"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


TITLE_FONT = _font("malgunbd.ttf", 24)
HEADER_FONT = _font("malgunbd.ttf", 17)
BODY_FONT = _font("malgun.ttf", 16)
SMALL_FONT = _font("malgun.ttf", 12)


def _rgb_from_fill(fill: PatternFill) -> tuple[int, int, int]:
    if fill.fill_type != "solid":
        return (255, 255, 255)
    rgb = fill.fgColor.rgb
    if not rgb or rgb == "00000000":
        return (255, 255, 255)
    value = str(rgb)[-6:]
    return tuple(int(value[index : index + 2], 16) for index in (0, 2, 4))


def _format_value(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return f"{value:,}"
    if isinstance(value, float):
        return f"{value:,.2f}".rstrip("0").rstrip(".")
    return str(value)


def _column_width(sheet, col_index: int) -> int:
    letter = sheet.cell(row=1, column=col_index).column_letter
    width = sheet.column_dimensions[letter].width or 12
    return max(92, min(int(width * 9), 180))


def _draw_text_clipped(draw: ImageDraw.ImageDraw, text: str, xy: tuple[int, int], width: int, font) -> None:
    if not text:
        return
    clipped = text
    while clipped and draw.textlength(clipped, font=font) > width - 16:
        clipped = clipped[:-1]
    if clipped != text and clipped:
        clipped = clipped[:-1] + "..."
    draw.text(xy, clipped, fill=(31, 41, 55), font=font)


def render_sheet(workbook_path: Path, sheet_name: str, title: str, output_path: Path) -> None:
    workbook = load_workbook(workbook_path, data_only=True)
    sheet = workbook[sheet_name] if sheet_name in workbook.sheetnames else workbook[workbook.sheetnames[0]]
    max_rows = min(sheet.max_row, 14)
    max_cols = min(sheet.max_column, 8)
    col_widths = [_column_width(sheet, index) for index in range(1, max_cols + 1)]
    row_height = 34
    title_height = 78
    margin = 18
    table_width = sum(col_widths)
    image_width = table_width + margin * 2
    image_height = title_height + row_height * max_rows + margin * 2

    image = Image.new("RGB", (image_width, image_height), (244, 247, 250))
    draw = ImageDraw.Draw(image)

    draw.rounded_rectangle(
        (margin, margin, image_width - margin, image_height - margin),
        radius=8,
        fill=(255, 255, 255),
        outline=(203, 213, 225),
        width=1,
    )
    draw.text((margin + 16, margin + 14), title, fill=(23, 50, 77), font=TITLE_FONT)
    draw.text(
        (margin + 18, margin + 48),
        f"{workbook_path.name} / {sheet.title}",
        fill=(82, 97, 115),
        font=SMALL_FONT,
    )

    y = margin + title_height
    for row in range(1, max_rows + 1):
        x = margin
        for col in range(1, max_cols + 1):
            cell = sheet.cell(row=row, column=col)
            width = col_widths[col - 1]
            fill = _rgb_from_fill(cell.fill)
            if row == 1 and fill == (255, 255, 255):
                fill = (31, 111, 95)
            draw.rectangle((x, y, x + width, y + row_height), fill=fill, outline=(226, 232, 240))
            font = HEADER_FONT if row == 1 else BODY_FONT
            text_color = (255, 255, 255) if row == 1 else (31, 41, 55)
            text = _format_value(cell.value)
            if text:
                clipped = text
                while clipped and draw.textlength(clipped, font=font) > width - 16:
                    clipped = clipped[:-1]
                if clipped != text and clipped:
                    clipped = clipped[:-1] + "..."
                draw.text((x + 8, y + 8), clipped, fill=text_color, font=font)
            if cell.comment:
                draw.polygon(
                    [(x + width - 12, y), (x + width, y), (x + width, y + 12)],
                    fill=(220, 38, 38),
                )
            x += width
        y += row_height

    output_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(output_path)


def main() -> int:
    missing = [name for name, *_ in FEATURE_SHOTS if not (DEMO_OUTPUT / name).exists()]
    if missing:
        joined = ", ".join(missing)
        raise SystemExit(f"demo output missing: {joined}. Run scripts/run_public_admin_demo.py first.")
    for workbook_name, sheet_name, title, output_name in FEATURE_SHOTS:
        output = IMAGE_OUTPUT / output_name
        render_sheet(DEMO_OUTPUT / workbook_name, sheet_name, title, output)
        print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
