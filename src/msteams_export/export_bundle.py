from __future__ import annotations

from pathlib import Path


def resolve_export_root(target: Path) -> tuple[Path, Path]:
    resolved = target.expanduser().resolve()
    if resolved.is_dir():
        index_path = resolved / "index.json"
        chats_dir = resolved / "chats"
    else:
        index_path = resolved
        chats_dir = resolved.parent / "chats"
    if not index_path.is_file():
        raise FileNotFoundError(f"Could not find index.json at {index_path}")
    if not chats_dir.is_dir():
        raise FileNotFoundError(f"Could not find chats directory at {chats_dir}")
    return index_path, chats_dir
