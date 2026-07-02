from __future__ import annotations

import json
import sqlite3
from pathlib import Path


class SettingsStore:
    """SQLite store for non-sensitive templates and job metadata."""

    def __init__(self, db_path: str | Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init()

    def save_mapping_template(self, name: str, mapping: dict[str, str]) -> None:
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO mapping_templates(name, mapping_json)
                VALUES (?, ?)
                ON CONFLICT(name) DO UPDATE SET mapping_json = excluded.mapping_json
                """,
                (name, json.dumps(mapping, ensure_ascii=False, sort_keys=True)),
            )

    def load_mapping_template(self, name: str) -> dict[str, str] | None:
        with sqlite3.connect(self.db_path) as conn:
            row = conn.execute(
                "SELECT mapping_json FROM mapping_templates WHERE name = ?",
                (name,),
            ).fetchone()
        if row is None:
            return None
        return json.loads(row[0])

    def record_job(self, job_type: str, output_path: str, success: bool, issue_count: int) -> None:
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                INSERT INTO job_history(job_type, output_path, success, issue_count)
                VALUES (?, ?, ?, ?)
                """,
                (job_type, output_path, int(success), issue_count),
            )

    def _init(self) -> None:
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS mapping_templates (
                    name TEXT PRIMARY KEY,
                    mapping_json TEXT NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS job_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_type TEXT NOT NULL,
                    output_path TEXT NOT NULL,
                    success INTEGER NOT NULL,
                    issue_count INTEGER NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
