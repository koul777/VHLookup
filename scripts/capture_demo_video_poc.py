from __future__ import annotations

import ctypes
import subprocess
import sys
import threading
import time
from pathlib import Path
from tkinter import END, Tk, Toplevel, messagebox, ttk

import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageGrab


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from vhlookup_app.tk_main import LocalApp  # noqa: E402
from vhlookup_core import AdminWorkbookTools  # noqa: E402


INPUT = ROOT / "samples" / "public_admin" / "08_sheet_merge" / "monthly_budget_sheets.xlsx"
WORK = ROOT / ".tmp" / "demo-video-work"
RAW_DIR = WORK / "raw"
RESULT = WORK / "08_sheet_merge_result.xlsx"
RAW_VIDEO = RAW_DIR / "08_sheet_merge_poc.webm"
FFMPEG = (
    WORK
    / "node_modules"
    / "@ffmpeg-installer"
    / "win32-x64"
    / "ffmpeg.exe"
)

OUTPUT_WIDTH = 1920
OUTPUT_HEIGHT = 1080
CAPTURE_FPS = 8
SCENE_SECONDS = 29


class CaptureState:
    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.subtitle = "여러 시트에 나뉜 자료를 한 표로 합쳐보겠습니다."
        self.label = "8. 여러 시트 합치기"
        self.cursor_target = (1580.0, 880.0)
        self.cursor_position = (1580.0, 880.0)
        self.zoom_scale = 1.0
        self.zoom_center = (960.0, 520.0)
        self.click_until = 0.0
        self.stop = False

    def update(self, **values: object) -> None:
        with self.lock:
            for key, value in values.items():
                setattr(self, key, value)

    def snapshot(self) -> dict[str, object]:
        with self.lock:
            current_x, current_y = self.cursor_position
            target_x, target_y = self.cursor_target
            current_x += (target_x - current_x) * 0.16
            current_y += (target_y - current_y) * 0.16
            self.cursor_position = (current_x, current_y)
            return {
                "subtitle": self.subtitle,
                "label": self.label,
                "cursor": self.cursor_position,
                "zoom_scale": self.zoom_scale,
                "zoom_center": self.zoom_center,
                "click_until": self.click_until,
                "stop": self.stop,
            }


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path("C:/Windows/Fonts/malgunbd.ttf" if bold else "C:/Windows/Fonts/malgun.ttf"),
        Path("C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default()


SUBTITLE_FONT = font(30, bold=True)
LABEL_FONT = font(24, bold=True)


def apply_zoom(image: Image.Image, scale: float, center: tuple[float, float]) -> Image.Image:
    if scale <= 1.001:
        return image
    width, height = image.size
    crop_width = max(1, int(width / scale))
    crop_height = max(1, int(height / scale))
    center_x = max(crop_width / 2, min(float(center[0]), width - crop_width / 2))
    center_y = max(crop_height / 2, min(float(center[1]), height - crop_height / 2))
    left = int(center_x - crop_width / 2)
    top = int(center_y - crop_height / 2)
    return image.crop((left, top, left + crop_width, top + crop_height)).resize(
        (width, height), Image.Resampling.LANCZOS
    )


def draw_overlays(image: Image.Image, snapshot: dict[str, object]) -> Image.Image:
    layer = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)

    label = str(snapshot["label"])
    label_bbox = draw.textbbox((0, 0), label, font=LABEL_FONT)
    label_width = label_bbox[2] - label_bbox[0]
    label_x, label_y = 54, OUTPUT_HEIGHT - 104
    draw.rounded_rectangle(
        (label_x - 18, label_y - 12, label_x + label_width + 18, label_y + 46),
        radius=14,
        fill=(31, 111, 95, 238),
    )
    draw.text((label_x, label_y), label, font=LABEL_FONT, fill=(255, 255, 255, 255))

    subtitle = str(snapshot["subtitle"])
    if subtitle:
        subtitle_bbox = draw.textbbox((0, 0), subtitle, font=SUBTITLE_FONT)
        subtitle_width = subtitle_bbox[2] - subtitle_bbox[0]
        bar_width = min(1480, subtitle_width + 92)
        bar_left = (OUTPUT_WIDTH - bar_width) // 2
        bar_top = OUTPUT_HEIGHT - 112
        draw.rounded_rectangle(
            (bar_left, bar_top, bar_left + bar_width, bar_top + 70),
            radius=18,
            fill=(15, 18, 28, 218),
        )
        draw.text(
            ((OUTPUT_WIDTH - subtitle_width) // 2, bar_top + 15),
            subtitle,
            font=SUBTITLE_FONT,
            fill=(255, 255, 255, 255),
        )

    cursor_x, cursor_y = snapshot["cursor"]
    cursor_x = float(cursor_x)
    cursor_y = float(cursor_y)
    if float(snapshot["click_until"]) > time.monotonic():
        draw.ellipse(
            (cursor_x - 29, cursor_y - 29, cursor_x + 29, cursor_y + 29),
            outline=(46, 163, 255, 230),
            width=6,
        )
    cursor_points = [
        (cursor_x, cursor_y),
        (cursor_x + 2, cursor_y + 36),
        (cursor_x + 11, cursor_y + 27),
        (cursor_x + 21, cursor_y + 48),
        (cursor_x + 31, cursor_y + 43),
        (cursor_x + 20, cursor_y + 23),
        (cursor_x + 35, cursor_y + 21),
    ]
    draw.polygon(cursor_points, fill=(255, 255, 255, 255), outline=(20, 31, 48, 255))
    draw.line(cursor_points + [cursor_points[0]], fill=(20, 31, 48, 255), width=3, joint="curve")
    return Image.alpha_composite(image.convert("RGBA"), layer).convert("RGB")


def capture_loop(
    state: CaptureState,
    capture_box: tuple[int, int, int, int],
    ready: threading.Event,
) -> None:
    command = [
        str(FFMPEG),
        "-y",
        "-loglevel",
        "error",
        "-f",
        "rawvideo",
        "-pix_fmt",
        "rgb24",
        "-s",
        f"{OUTPUT_WIDTH}x{OUTPUT_HEIGHT}",
        "-r",
        str(CAPTURE_FPS),
        "-i",
        "-",
        "-an",
        "-c:v",
        "libvpx-vp9",
        "-deadline",
        "realtime",
        "-cpu-used",
        "6",
        "-b:v",
        "4M",
        "-pix_fmt",
        "yuv420p",
        str(RAW_VIDEO),
    ]
    process = subprocess.Popen(command, stdin=subprocess.PIPE)
    assert process.stdin is not None
    ready.set()
    frame_seconds = 1 / CAPTURE_FPS
    next_frame = time.monotonic()
    try:
        while True:
            snapshot = state.snapshot()
            if bool(snapshot["stop"]):
                break
            grabbed = ImageGrab.grab(bbox=capture_box, all_screens=True).convert("RGB")
            if grabbed.size != (OUTPUT_WIDTH, OUTPUT_HEIGHT):
                grabbed = grabbed.resize((OUTPUT_WIDTH, OUTPUT_HEIGHT), Image.Resampling.LANCZOS)
            grabbed = apply_zoom(
                grabbed,
                float(snapshot["zoom_scale"]),
                tuple(snapshot["zoom_center"]),
            )
            process.stdin.write(draw_overlays(grabbed, snapshot).tobytes())
            next_frame += frame_seconds
            time.sleep(max(0.0, next_frame - time.monotonic()))
    finally:
        process.stdin.close()
        return_code = process.wait(timeout=30)
        if return_code:
            raise RuntimeError(f"ffmpeg capture failed with exit code {return_code}")


def iter_widgets(widget):
    for child in widget.winfo_children():
        yield child
        yield from iter_widgets(child)


def button_with(root: Tk, predicate):
    for widget in iter_widgets(root):
        if isinstance(widget, ttk.Button):
            try:
                if predicate(str(widget.cget("text"))):
                    return widget
            except Exception:
                continue
    return None


def output_position(widget, capture_box: tuple[int, int, int, int]) -> tuple[float, float]:
    source_width = capture_box[2] - capture_box[0]
    source_height = capture_box[3] - capture_box[1]
    center_x = widget.winfo_rootx() + widget.winfo_width() / 2 - capture_box[0]
    center_y = widget.winfo_rooty() + widget.winfo_height() / 2 - capture_box[1]
    return center_x * OUTPUT_WIDTH / source_width, center_y * OUTPUT_HEIGHT / source_height


def create_result_preview(root: Tk, capture_box: tuple[int, int, int, int], state: CaptureState) -> None:
    if not RESULT.exists():
        root.after(350, lambda: create_result_preview(root, capture_box, state))
        return

    table = pd.read_excel(RESULT, sheet_name=0)
    preview = Toplevel(root)
    preview.title("합친결과 미리보기")
    preview.geometry("1240x630+320+175")
    preview.transient(root)
    preview.lift()

    frame = ttk.Frame(preview, padding=20)
    frame.pack(fill="both", expand=True)
    ttk.Label(frame, text="합친결과 미리보기", font=("맑은 고딕", 18, "bold")).pack(anchor="w")
    ttk.Label(
        frame,
        text="1월 · 2월 시트의 열 이름을 맞춰 한 표로 연결하고, 원본 시트명을 함께 기록했습니다.",
        foreground="#526173",
        font=("맑은 고딕", 11),
    ).pack(anchor="w", pady=(4, 14))

    tree_frame = ttk.Frame(frame)
    tree_frame.pack(fill="both", expand=True)
    columns = [str(column) for column in table.columns]
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=14)
    for column in columns:
        tree.heading(column, text=column)
        tree.column(column, width=165, minwidth=110, anchor="center")
    for row in table.head(14).itertuples(index=False, name=None):
        tree.insert("", END, values=["" if pd.isna(value) else value for value in row])
    y_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
    tree.grid(row=0, column=0, sticky="nsew")
    y_scroll.grid(row=0, column=1, sticky="ns")
    x_scroll.grid(row=1, column=0, sticky="ew")
    tree_frame.columnconfigure(0, weight=1)
    tree_frame.rowconfigure(0, weight=1)
    ttk.Label(
        frame,
        text=f"완료 · {len(table):,}행 · 결과 파일: {RESULT.name}",
        foreground="#1F6F5F",
        font=("맑은 고딕", 11, "bold"),
    ).pack(anchor="w", pady=(12, 0))

    preview.update_idletasks()
    tree_target = output_position(tree, capture_box)
    state.update(
        subtitle="열이 달라도 이름을 기준으로 맞추고, 원본 시트명까지 남깁니다.",
        cursor_target=tree_target,
    )


def main() -> int:
    if not INPUT.exists():
        raise FileNotFoundError(INPUT)
    if not FFMPEG.exists():
        raise FileNotFoundError(FFMPEG)
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    RESULT.unlink(missing_ok=True)
    RAW_VIDEO.unlink(missing_ok=True)

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    root = Tk()
    root.geometry("1600x900+60+40")
    root.update_idletasks()
    app = LocalApp(root)
    root.geometry("1600x900+60+40")
    root.update_idletasks()
    root.lift()
    root.attributes("-topmost", True)
    root.after(800, lambda: root.attributes("-topmost", False))

    capture_box = (
        root.winfo_rootx(),
        root.winfo_rooty(),
        root.winfo_rootx() + root.winfo_width(),
        root.winfo_rooty() + root.winfo_height(),
    )
    state = CaptureState()

    app._ask_file_or_none = lambda _title: INPUT
    app._ask_save_path_or_none = lambda _title, _base_name: RESULT
    app.open_folder = lambda _path: None
    messagebox.showinfo = lambda *_args, **_kwargs: None
    messagebox.showwarning = lambda *_args, **_kwargs: None
    messagebox.showerror = lambda *_args, **_kwargs: None

    original_writer = AdminWorkbookTools.write_merged_sheets_workbook

    def paced_writer(self, path, output_path, sheet_names=None, progress_callback=None):
        def paced_progress(percent: int, text: str) -> None:
            if progress_callback is not None:
                progress_callback(percent, text)
            time.sleep(0.45)

        return original_writer(
            self,
            path,
            output_path,
            sheet_names=sheet_names,
            progress_callback=paced_progress,
        )

    AdminWorkbookTools.write_merged_sheets_workbook = paced_writer

    ready = threading.Event()
    capture_thread = threading.Thread(
        target=capture_loop,
        args=(state, capture_box, ready),
        daemon=True,
    )
    capture_thread.start()
    ready.wait(timeout=10)

    def click_main_menu() -> None:
        button = button_with(root, lambda text: text.startswith("8."))
        if button is None:
            raise RuntimeError("8번 메뉴 버튼을 찾지 못했습니다.")
        state.update(
            subtitle="8번 ‘여러 시트 합치기’를 선택합니다.",
            cursor_target=output_position(button, capture_box),
        )

        def invoke() -> None:
            state.update(click_until=time.monotonic() + 0.65)
            root.after(450, prepare_sheet_dialog)
            button.invoke()

        root.after(1150, invoke)

    def prepare_sheet_dialog() -> None:
        for widget in root.winfo_children():
            if isinstance(widget, Toplevel) and "시트 선택" in widget.title():
                widget.geometry("680x520+500+210")
                widget.lift()
                break
        state.update(subtitle="합칠 시트를 확인합니다. 기본값은 전체 시트 선택입니다.")
        root.after(2450, choose_selected_sheets)

    def choose_selected_sheets() -> None:
        button = button_with(root, lambda text: "선택한 시트 합치기" in text)
        if button is None:
            root.after(250, choose_selected_sheets)
            return
        state.update(cursor_target=output_position(button, capture_box))

        def invoke() -> None:
            state.update(
                subtitle="선택한 시트를 아래로 이어 붙여 새 엑셀을 만듭니다.",
                click_until=time.monotonic() + 0.65,
            )
            button.invoke()

        root.after(1150, invoke)

    def show_preview() -> None:
        create_result_preview(root, capture_box, state)

    def zoom_preview() -> None:
        state.update(
            subtitle="추가 열은 그대로 보존되고, 없는 값은 빈칸으로 정리됩니다.",
            zoom_scale=1.35,
            zoom_center=(960.0, 485.0),
        )

    def finish_zoom() -> None:
        state.update(
            subtitle="결과 파일에는 ‘합친결과’와 ‘확인사항’ 시트가 함께 저장됩니다.",
            zoom_scale=1.0,
            cursor_target=(1470.0, 840.0),
        )

    def finish() -> None:
        state.update(subtitle="메뉴 8 사용이 완료되었습니다.")

    def stop() -> None:
        state.update(stop=True)
        root.after(300, root.destroy)

    root.after(2200, click_main_menu)
    root.after(12000, show_preview)
    root.after(16800, zoom_preview)
    root.after(21400, finish_zoom)
    root.after(26000, finish)
    root.after(SCENE_SECONDS * 1000, stop)

    try:
        root.mainloop()
    finally:
        state.update(stop=True)
        capture_thread.join(timeout=45)
        AdminWorkbookTools.write_merged_sheets_workbook = original_writer
    if not RAW_VIDEO.exists() or RAW_VIDEO.stat().st_size < 100_000:
        raise RuntimeError("대표 영상 캡처 결과가 생성되지 않았습니다.")
    print(RAW_VIDEO)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
