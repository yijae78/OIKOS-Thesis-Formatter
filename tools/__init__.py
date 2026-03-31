"""Thesis Editor Tools — python-docx based Word document manipulation toolkit."""

import os
import shutil
from datetime import datetime
from pathlib import Path

# Project paths
PROJECT_DIR = Path(__file__).parent.parent
ORIGINAL_FILE = PROJECT_DIR / "1. 신이재 박사학위논문 원본.docx"
WORKING_FILE = PROJECT_DIR / "working_copy.docx"
BACKUP_DIR = PROJECT_DIR / "backups"

BACKUP_DIR.mkdir(exist_ok=True)


def ensure_working_copy():
    """Create working copy from original if it doesn't exist."""
    if not WORKING_FILE.exists():
        if not ORIGINAL_FILE.exists():
            raise FileNotFoundError(f"Original file not found: {ORIGINAL_FILE}")
        shutil.copy2(ORIGINAL_FILE, WORKING_FILE)
        print(f"[INIT] Working copy created: {WORKING_FILE.name}")
    return WORKING_FILE


def create_backup(label: str = "") -> Path:
    """Create timestamped backup of working copy. Returns backup path."""
    if not WORKING_FILE.exists():
        raise FileNotFoundError("No working copy to back up")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = f"_{label}" if label else ""
    backup_name = f"backup_{timestamp}{suffix}.docx"
    backup_path = BACKUP_DIR / backup_name

    shutil.copy2(WORKING_FILE, backup_path)
    print(f"[BACKUP] Created: {backup_name} ({backup_path.stat().st_size / 1024 / 1024:.1f} MB)")
    return backup_path


def restore_backup(backup_path: Path) -> None:
    """Restore working copy from a backup."""
    if not backup_path.exists():
        raise FileNotFoundError(f"Backup not found: {backup_path}")
    shutil.copy2(backup_path, WORKING_FILE)
    print(f"[RESTORE] Restored from: {backup_path.name}")


def list_backups() -> list[Path]:
    """List all available backups, newest first."""
    backups = sorted(BACKUP_DIR.glob("backup_*.docx"), reverse=True)
    return backups


def validate_docx(filepath: Path) -> bool:
    """L0 integrity check: file exists, opens, has sections."""
    from docx import Document

    if not filepath.exists():
        print(f"[L0 FAIL] File does not exist: {filepath}")
        return False
    if filepath.stat().st_size == 0:
        print(f"[L0 FAIL] File is empty: {filepath}")
        return False

    try:
        doc = Document(str(filepath))
        section_count = len(doc.sections)
        para_count = len(doc.paragraphs)
        print(f"[L0 PASS] Sections: {section_count}, Paragraphs: {para_count}")
        return True
    except Exception as e:
        print(f"[L0 FAIL] Cannot open document: {e}")
        return False
