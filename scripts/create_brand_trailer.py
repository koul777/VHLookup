from __future__ import annotations

import subprocess
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont


ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / ".tmp" / "brand-trailer-work"
VIDEOS = ROOT / "docs" / "videos"
PREVIEWS = VIDEOS / "previews"
FFMPEG = (
    ROOT
    / ".tmp"
    / "demo-video-work"
    / "node_modules"
    / "ffmpeg-static"
    / "ffmpeg.exe"
)

WIDTH = 1920
HEIGHT = 1080
FPS = 30
TRANSITION = 0.45

FONT_BOLD = Path(r"C:\Windows\Fonts\NotoSansKR-Bold.ttf")
FONT_MEDIUM = Path(r"C:\Windows\Fonts\NotoSansKR-Medium.ttf")
FONT_REGULAR = Path(r"C:\Windows\Fonts\NotoSansKR-Regular.ttf")

GREEN = "#1E7564"
MINT = "#55D6B0"
NAVY = "#0F121C"
WHITE = "#F8FBFA"
MUTED = "#C9D9D5"


def font(path: Path, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(path), size)


def run(command: list[str]) -> None:
    print(" ".join(command))
    subprocess.run(command, check=True, cwd=ROOT)


def gradient_background() -> Image.Image:
    image = Image.new("RGB", (WIDTH, HEIGHT))
    pixels = image.load()
    top = (8, 30, 32)
    bottom = (30, 117, 100)
    for y in range(HEIGHT):
        mix = y / max(HEIGHT - 1, 1)
        row = tuple(int(top[i] * (1 - mix) + bottom[i] * mix) for i in range(3))
        for x in range(WIDTH):
            side = abs(x - WIDTH / 2) / (WIDTH / 2)
            pixels[x, y] = tuple(max(0, int(value * (1 - side * 0.12))) for value in row)
    return image


def draw_brand_texture(image: Image.Image) -> None:
    draw = ImageDraw.Draw(image, "RGBA")
    for x in range(80, WIDTH, 120):
        draw.line((x, 0, x, HEIGHT), fill=(255, 255, 255, 10), width=1)
    for y in range(60, HEIGHT, 120):
        draw.line((0, y, WIDTH, y), fill=(255, 255, 255, 8), width=1)
    draw.ellipse((1350, -240, 2130, 540), outline=(85, 214, 176, 42), width=5)
    draw.ellipse((1450, -140, 2030, 440), outline=(255, 255, 255, 24), width=2)
    draw.ellipse((-240, 760, 360, 1360), outline=(85, 214, 176, 28), width=4)


def draw_wordmark(draw: ImageDraw.ImageDraw, x: int, y: int, *, large: bool = True) -> None:
    box = 150 if large else 94
    radius = 34 if large else 22
    draw.rounded_rectangle((x, y, x + box, y + box), radius=radius, fill=(85, 214, 176, 255))
    monogram_font = font(FONT_BOLD, 64 if large else 40)
    draw.text((x + box / 2, y + box / 2 - 4), "VH", font=monogram_font, fill=NAVY, anchor="mm")


def create_intro(path: Path) -> None:
    image = gradient_background()
    draw_brand_texture(image)
    draw = ImageDraw.Draw(image, "RGBA")
    draw_wordmark(draw, 170, 300, large=True)
    draw.text((370, 312), "VHLookup", font=font(FONT_BOLD, 104), fill=WHITE)
    draw.text((374, 425), "LOCAL V1.4", font=font(FONT_MEDIUM, 36), fill=MINT)
    draw.rounded_rectangle((170, 545, 1500, 550), radius=2, fill=(85, 214, 176, 220))
    draw.text(
        (170, 610),
        "반복되는 엑셀 업무를, 더 빠르고 안전하게.",
        font=font(FONT_BOLD, 56),
        fill=WHITE,
    )
    draw.text(
        (174, 700),
        "파일 수합부터 검증·요약까지, 로컬 PC에서 한 번에 처리합니다.",
        font=font(FONT_REGULAR, 31),
        fill=MUTED,
    )
    draw.text((170, 914), "LOCAL EXCEL AUTOMATION", font=font(FONT_MEDIUM, 25), fill=MINT)
    draw.text((170, 958), "WINDOWS  ·  8 WORKFLOWS  ·  100% LOCAL", font=font(FONT_REGULAR, 21), fill=MUTED)
    image.save(path)


def create_outro(path: Path) -> None:
    image = gradient_background()
    draw_brand_texture(image)
    draw = ImageDraw.Draw(image, "RGBA")
    draw_wordmark(draw, 1580, 100, large=False)
    draw.text((180, 190), "엑셀 업무의", font=font(FONT_MEDIUM, 50), fill=MINT)
    draw.text((180, 255), "새로운 기본", font=font(FONT_BOLD, 104), fill=WHITE)
    draw.text((180, 405), "VHLookup Local V1.4", font=font(FONT_BOLD, 62), fill=WHITE)
    draw.text(
        (184, 505),
        "다운로드하고 바로 시작하세요.",
        font=font(FONT_MEDIUM, 38),
        fill=MUTED,
    )
    draw.rounded_rectangle((180, 630, 1150, 755), radius=28, fill=(85, 214, 176, 255))
    draw.text((665, 691), "github.com/koul777/VHLookup", font=font(FONT_BOLD, 37), fill=NAVY, anchor="mm")
    draw.text(
        (184, 840),
        "원본 파일은 그대로  ·  결과는 새 파일로  ·  데이터는 PC 안에서",
        font=font(FONT_REGULAR, 25),
        fill=MUTED,
    )
    image.save(path)


def create_local_card(path: Path) -> None:
    source = Image.open(ROOT / "docs" / "assets" / "vhlookup-local-main.jpg").convert("RGB")
    scale = max(WIDTH / source.width, HEIGHT / source.height)
    resized = source.resize((int(source.width * scale), int(source.height * scale)), Image.Resampling.LANCZOS)
    left = (resized.width - WIDTH) // 2
    top = (resized.height - HEIGHT) // 2
    image = resized.crop((left, top, left + WIDTH, top + HEIGHT)).filter(ImageFilter.GaussianBlur(3.0))
    tint = Image.new("RGBA", (WIDTH, HEIGHT), (5, 30, 29, 184))
    image = Image.alpha_composite(image.convert("RGBA"), tint)
    draw = ImageDraw.Draw(image, "RGBA")
    draw.rounded_rectangle((250, 190, 1670, 890), radius=42, fill=(10, 35, 36, 220), outline=(85, 214, 176, 110), width=3)

    shield = [(960, 285), (1065, 330), (1048, 480), (960, 555), (872, 480), (855, 330)]
    draw.polygon(shield, fill=(85, 214, 176, 38), outline=(85, 214, 176, 255), width=8)
    draw.line((905, 414, 946, 455, 1018, 375), fill=(248, 251, 250, 255), width=14, joint="curve")
    draw.text((960, 620), "100% LOCAL", font=font(FONT_BOLD, 39), fill=MINT, anchor="mm")
    draw.text((960, 700), "파일은 PC 밖으로 나가지 않습니다.", font=font(FONT_BOLD, 55), fill=WHITE, anchor="mm")
    draw.text(
        (960, 790),
        "원본은 수정하지 않고, 결과는 새 엑셀 파일로 저장합니다.",
        font=font(FONT_REGULAR, 29),
        fill=MUTED,
        anchor="mm",
    )
    image.convert("RGB").save(path)


def create_overlay(path: Path, label: str, subtitle: str) -> None:
    image = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image, "RGBA")
    draw.rounded_rectangle((1590, 52, 1852, 112), radius=26, fill=(15, 18, 28, 210), outline=(85, 214, 176, 170), width=2)
    draw.ellipse((1620, 74, 1636, 90), fill=(85, 214, 176, 255))
    draw.text((1650, 82), "100% LOCAL", font=font(FONT_MEDIUM, 20), fill=WHITE, anchor="lm")

    draw.rectangle((0, 914, WIDTH, HEIGHT), fill=(15, 18, 28, 247))
    draw.rectangle((0, 914, WIDTH, 920), fill=(85, 214, 176, 255))
    draw.text((62, 945), "VHLOOKUP LOCAL · PRODUCT TRAILER", font=font(FONT_MEDIUM, 18), fill=MINT)

    label_font = font(FONT_BOLD, 22)
    label_box = draw.textbbox((0, 0), label, font=label_font)
    label_width = label_box[2] - label_box[0] + 58
    draw.rounded_rectangle((62, 987, 62 + label_width, 1044), radius=20, fill=(30, 117, 100, 255))
    draw.text((91, 1015), label, font=label_font, fill=WHITE, anchor="lm")

    subtitle_x = 62 + label_width + 48
    draw.text((subtitle_x, 1015), subtitle, font=font(FONT_BOLD, 33), fill=WHITE, anchor="lm")
    image.save(path)


def render_still(image: Path, duration: float, output: Path, zoom_rate: float = 0.00016) -> None:
    frames = int(duration * FPS)
    zoom = f"min(zoom+{zoom_rate:.5f},1.025)"
    filter_value = (
        f"zoompan=z='{zoom}':x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)'"
        f":d=1:s={WIDTH}x{HEIGHT}:fps={FPS},format=yuv420p"
    )
    run(
        [
            str(FFMPEG),
            "-hide_banner",
            "-loglevel",
            "error",
            "-y",
            "-framerate",
            str(FPS),
            "-loop",
            "1",
            "-i",
            str(image),
            "-vf",
            filter_value,
            "-frames:v",
            str(frames),
            "-an",
            "-c:v",
            "libx264",
            "-preset",
            "medium",
            "-crf",
            "18",
            "-pix_fmt",
            "yuv420p",
            str(output),
        ]
    )


def render_clip(source: Path, start: float, duration: float, overlay: Path, output: Path) -> None:
    run(
        [
            str(FFMPEG),
            "-hide_banner",
            "-loglevel",
            "error",
            "-y",
            "-ss",
            str(start),
            "-i",
            str(source),
            "-framerate",
            str(FPS),
            "-loop",
            "1",
            "-i",
            str(overlay),
            "-t",
            str(duration),
            "-filter_complex",
            f"[0:v]fps={FPS},scale={WIDTH}:{HEIGHT},setsar=1[base];[base][1:v]overlay=0:0:shortest=1,format=yuv420p[out]",
            "-map",
            "[out]",
            "-an",
            "-c:v",
            "libx264",
            "-preset",
            "medium",
            "-crf",
            "18",
            "-r",
            str(FPS),
            str(output),
        ]
    )


def concat_with_crossfades(segments: list[Path], durations: list[float], output: Path) -> None:
    command = [str(FFMPEG), "-hide_banner", "-loglevel", "error", "-y"]
    for segment in segments:
        command.extend(["-i", str(segment)])

    filters: list[str] = []
    current = "[0:v]"
    elapsed = durations[0]
    for index in range(1, len(segments)):
        target = f"[v{index}]"
        offset = elapsed - TRANSITION
        filters.append(
            f"{current}[{index}:v]xfade=transition=fade:duration={TRANSITION}:offset={offset:.3f}{target}"
        )
        current = target
        elapsed += durations[index] - TRANSITION

    command.extend(
        [
            "-filter_complex",
            ";".join(filters),
            "-map",
            current,
            "-an",
            "-c:v",
            "libx264",
            "-preset",
            "medium",
            "-crf",
            "18",
            "-pix_fmt",
            "yuv420p",
            "-movflags",
            "+faststart",
            str(output),
        ]
    )
    run(command)


def create_gif(source: Path, output: Path) -> None:
    run(
        [
            str(FFMPEG),
            "-hide_banner",
            "-loglevel",
            "error",
            "-y",
            "-i",
            str(source),
            "-filter_complex",
            "fps=6,scale=800:-1:flags=lanczos,split[s0][s1];"
            "[s0]palettegen=max_colors=96:stats_mode=diff[p];"
            "[s1][p]paletteuse=dither=bayer:bayer_scale=5:diff_mode=rectangle",
            "-loop",
            "0",
            str(output),
        ]
    )


def create_qa_artifacts(source: Path) -> None:
    run(
        [
            str(FFMPEG),
            "-hide_banner",
            "-loglevel",
            "error",
            "-y",
            "-i",
            str(source),
            "-vf",
            "fps=1/5,scale=640:360,tile=3x3",
            "-frames:v",
            "1",
            str(WORK / "brand_trailer_montage.png"),
        ]
    )
    for index, timestamp in enumerate((2, 8, 14, 20, 26, 32, 38), start=1):
        run(
            [
                str(FFMPEG),
                "-hide_banner",
                "-loglevel",
                "error",
                "-y",
                "-ss",
                str(timestamp),
                "-i",
                str(source),
                "-frames:v",
                "1",
                str(WORK / f"spot_{index:02d}_{timestamp:02d}s.png"),
            ]
        )


def main() -> None:
    if not FFMPEG.exists():
        raise FileNotFoundError(f"ffmpeg not found: {FFMPEG}")
    WORK.mkdir(parents=True, exist_ok=True)
    PREVIEWS.mkdir(parents=True, exist_ok=True)

    intro = WORK / "intro.png"
    outro = WORK / "outro.png"
    local = WORK / "local.png"
    create_intro(intro)
    create_outro(outro)
    create_local_card(local)

    overlays = {
        "overview": ("8 WORKFLOWS", "클릭 몇 번으로 여덟 가지 엑셀 업무를 처리합니다."),
        "privacy": ("PRIVACY FIRST", "개인정보는 가리고, 업무 데이터는 그대로 지킵니다."),
        "merge_files": ("MERGE SMARTER", "흩어진 파일과 시트를 한 번에 정리합니다."),
        "merge_sheets": ("MERGE SMARTER", "흩어진 파일과 시트를 한 번에 정리합니다."),
        "validation": ("CHECK & SUMMARIZE", "차이는 놓치지 않고, 요약은 바로 완성합니다."),
        "pivot": ("CHECK & SUMMARIZE", "차이는 놓치지 않고, 요약은 바로 완성합니다."),
    }
    overlay_paths: dict[str, Path] = {}
    for name, (label, subtitle) in overlays.items():
        target = WORK / f"overlay_{name}.png"
        create_overlay(target, label, subtitle)
        overlay_paths[name] = target

    durations = [5.0, 6.0, 6.0, 4.0, 4.0, 4.0, 4.0, 6.0, 5.0]
    segments = [WORK / f"segment_{index:02d}.mp4" for index in range(1, 10)]

    render_still(intro, durations[0], segments[0], zoom_rate=0.00018)
    render_clip(VIDEOS / "01_privacy_masking.mp4", 0.0, durations[1], overlay_paths["overview"], segments[1])
    render_clip(VIDEOS / "01_privacy_masking.mp4", 17.0, durations[2], overlay_paths["privacy"], segments[2])
    render_clip(VIDEOS / "03_merge_files.mp4", 17.0, durations[3], overlay_paths["merge_files"], segments[3])
    render_clip(VIDEOS / "08_sheet_merge.mp4", 17.0, durations[4], overlay_paths["merge_sheets"], segments[4])
    render_clip(VIDEOS / "04_before_after_validation.mp4", 17.0, durations[5], overlay_paths["validation"], segments[5])
    render_clip(VIDEOS / "06_pivot_summary.mp4", 17.0, durations[6], overlay_paths["pivot"], segments[6])
    render_still(local, durations[7], segments[7], zoom_rate=0.00010)
    render_still(outro, durations[8], segments[8], zoom_rate=0.00014)

    trailer = VIDEOS / "VHLookup_brand_trailer.mp4"
    preview = PREVIEWS / "VHLookup_brand_trailer.gif"
    concat_with_crossfades(segments, durations, trailer)
    create_gif(trailer, preview)
    create_qa_artifacts(trailer)
    print(trailer)
    print(preview)


if __name__ == "__main__":
    main()
