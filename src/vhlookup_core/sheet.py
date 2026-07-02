from __future__ import annotations

from vhlookup_core.header import HeaderDetector
from vhlookup_core.models import SheetData, SheetProfile, WorkbookSource
from vhlookup_core.normalization import is_blank


class SheetDetector:
    def __init__(self, header_detector: HeaderDetector | None = None) -> None:
        self.header_detector = header_detector or HeaderDetector()

    def profile(self, sheet: SheetData) -> SheetProfile:
        frame = sheet.frame
        non_empty = int(frame.map(lambda value: not is_blank(value)).sum().sum()) if not frame.empty else 0
        detection = self.header_detector.detect(sheet)
        return SheetProfile(
            sheet_name=sheet.name,
            row_count=len(frame.index),
            column_count=len(frame.columns),
            non_empty_cells=non_empty,
            header_confidence=detection.confidence,
        )

    def select(self, source: WorkbookSource) -> SheetData:
        if not source.sheets:
            raise ValueError(f"No sheets found in {source.path}")
        return max(
            source.sheets,
            key=lambda sheet: (
                self.profile(sheet).header_confidence,
                self.profile(sheet).non_empty_cells,
                self.profile(sheet).row_count,
            ),
        )
