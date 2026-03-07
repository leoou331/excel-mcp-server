"""Tests for file manager safety guarantees."""

import os
import tempfile
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_mcp_server import add_allowed_directory
from excel_mcp_server.utils.file_manager import FileManager


def _create_workbook(file_path: Path, value: str = "before") -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = value
    workbook.save(file_path)
    workbook.close()


def test_safe_write_uses_atomic_temp_file() -> None:
    """Workbook saves should go through a temp file before replacing the target."""
    with tempfile.TemporaryDirectory() as directory:
        add_allowed_directory(directory)
        file_path = Path(directory) / "atomic.xlsx"
        _create_workbook(file_path)

        manager = FileManager()
        save_targets: list[str] = []

        with manager.safe_write(file_path) as workbook:
            workbook.active["A1"] = "after"
            original_save = workbook.save

            def wrapped_save(target: str) -> None:
                save_targets.append(target)
                original_save(target)

            workbook.save = wrapped_save  # type: ignore[method-assign]

        assert save_targets
        assert Path(save_targets[0]) != file_path
        assert Path(save_targets[0]).parent == file_path.parent

        reloaded = load_workbook(file_path, data_only=True)
        assert reloaded.active["A1"].value == "after"
        reloaded.close()


def test_safe_open_cleans_stale_lock_files() -> None:
    """Stale lock files should be removed automatically before opening a workbook."""
    with tempfile.TemporaryDirectory() as directory:
        add_allowed_directory(directory)
        file_path = Path(directory) / "locked.xlsx"
        _create_workbook(file_path)

        lock_dir = file_path.parent / ".excel_locks"
        lock_dir.mkdir(exist_ok=True)
        stale_lock = lock_dir / "orphan.lock"
        stale_lock.write_text("stale")

        stale_at = time.time() - 3600
        os.utime(stale_lock, (stale_at, stale_at))

        manager = FileManager()
        with manager.safe_open(file_path, read_only=True):
            pass

        assert not stale_lock.exists()
        assert not (lock_dir / f"{file_path.name}.lock").exists()
