from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class ProjectPaths:
    root: Path
    docs: Path
    src: Path
    tests: Path


def detect_project_paths(root: Path | None = None) -> ProjectPaths:
    base = (root or Path.cwd()).resolve()
    return ProjectPaths(
        root=base,
        docs=base / "Doc",
        src=base / "src",
        tests=base / "tests",
    )

