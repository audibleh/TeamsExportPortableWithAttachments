from __future__ import annotations

from typing import Any
from urllib.parse import urlparse


_EXCLUDED_EXTENSIONS = {
    "gif",
    "gifv",
    "mp4",
    "m4v",
    "mov",
    "avi",
    "mkv",
    "webm",
    "wmv",
    "mpeg",
    "mpg",
    "3gp",
    "m3u8",
    "ts",
}


def keep_attachment(
    *,
    label: str | None,
    href: str | None,
    type_value: str | None,
    kind: str | None = None,
) -> bool:
    normalized_type = (type_value or "").strip().lower().split(";", 1)[0]
    normalized_kind = (kind or "").strip().lower()
    if normalized_kind == "video":
        return False
    if normalized_type.startswith("video/"):
        return False
    if normalized_type in {"gif", "image/gif"}:
        return False
    suffixes = {
        _suffix(label),
        _suffix(href),
        normalized_type.split("/", 1)[-1] if "/" in normalized_type else normalized_type,
    }
    return not any(suffix in _EXCLUDED_EXTENSIONS for suffix in suffixes if suffix)


def filter_attachment_dicts(items: Any) -> list[dict[str, Any]]:
    if not isinstance(items, list):
        return []
    kept: list[dict[str, Any]] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        if not keep_attachment(
            label=_optional_str(item.get("label")),
            href=_optional_str(item.get("href")),
            type_value=_optional_str(item.get("type")),
            kind=_optional_str(item.get("kind")),
        ):
            continue
        kept.append(item)
    return kept


def _suffix(value: str | None) -> str:
    if not value:
        return ""
    text = value.strip()
    if not text:
        return ""
    parsed = urlparse(text)
    target = parsed.path or text
    if "." not in target:
        return ""
    return target.rsplit(".", 1)[-1].strip().lower()


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None
