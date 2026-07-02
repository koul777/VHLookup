from __future__ import annotations

from pathlib import Path

import pandas as pd

from vhlookup_core.models import SheetData, WorkbookSource


class ExcelLoader:
    """Loads workbook-like local files without mutating originals."""

    excel_suffixes = {".xlsx", ".xlsm"}
    csv_suffixes = {".csv"}
    encodings = ("utf-8-sig", "utf-8", "cp949", "euc-kr")

    def load(self, path: str | Path) -> WorkbookSource:
        file_path = Path(path)
        if not file_path.exists():
            raise FileNotFoundError(file_path)
        suffix = file_path.suffix.lower()
        if suffix in self.excel_suffixes:
            return self._load_excel(file_path)
        if suffix in self.csv_suffixes:
            return self._load_csv(file_path)
        raise ValueError(f"Unsupported file type: {file_path.suffix}")

    def _load_excel(self, path: Path) -> WorkbookSource:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
        return WorkbookSource(
            path=path,
            sheets=tuple(SheetData(name=name, frame=frame) for name, frame in sheets.items()),
        )

    def _load_csv(self, path: Path) -> WorkbookSource:
        last_error: Exception | None = None
        for encoding in self.encodings:
            try:
                frame = pd.read_csv(path, header=None, dtype=object, encoding=encoding)
                return WorkbookSource(path=path, sheets=(SheetData(name=path.stem, frame=frame),))
            except UnicodeDecodeError as exc:
                last_error = exc
        if last_error is not None:
            raise last_error
        raise ValueError(f"Could not load CSV: {path}")
