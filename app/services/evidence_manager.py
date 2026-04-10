from __future__ import annotations

from pathlib import Path
import shutil
import uuid


BASE_DATA_DIR = Path("reportforge_data")
EVIDENCE_DIR = BASE_DATA_DIR / "evidence"


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_project_evidence_dir(project_id: int) -> Path:
    return ensure_directory(EVIDENCE_DIR / f"project_{project_id}")


def get_finding_evidence_dir(project_id: int, finding_id: int) -> Path:
    project_dir = get_project_evidence_dir(project_id)
    return ensure_directory(project_dir / f"finding_{finding_id}")


def copy_evidence_files(
    source_paths: list[str],
    project_id: int,
    finding_id: int,
) -> list[str]:
    target_dir = get_finding_evidence_dir(project_id, finding_id)
    saved_paths: list[str] = []

    for source in source_paths:
        source_path = Path(source)
        if not source_path.exists() or not source_path.is_file():
            continue

        safe_name = source_path.name
        unique_name = f"{uuid.uuid4().hex[:8]}_{safe_name}"
        destination = target_dir / unique_name
        shutil.copy2(source_path, destination)
        saved_paths.append(str(destination))

    return saved_paths