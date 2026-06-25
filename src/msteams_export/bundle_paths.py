from __future__ import annotations

from pathlib import Path, PurePosixPath


def normalize_bundle_relative_path(value: str | None) -> Path | None:
    if value is None:
        return None
    text = str(value).strip().replace("\\", "/")
    if not text:
        return None
    path = PurePosixPath(text)
    if path.is_absolute():
        return None
    parts = [part for part in path.parts if part not in {"", "."}]
    if not parts or any(part == ".." for part in parts):
        return None
    return Path(*parts)


def resolve_bundle_relative_path(bundle_root: Path, value: str | None) -> Path | None:
    relative = normalize_bundle_relative_path(value)
    if relative is None:
        return None
    candidate = (bundle_root / relative).resolve()
    try:
        candidate.relative_to(bundle_root)
    except ValueError:
        return None
    return candidate


def bundle_relative_path_string(bundle_root: Path, path: Path) -> str:
    return path.resolve().relative_to(bundle_root).as_posix()
