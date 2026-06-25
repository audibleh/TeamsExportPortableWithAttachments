from __future__ import annotations

from copy import deepcopy
from datetime import UTC, datetime
from typing import Any
from urllib.parse import quote


def normalize_conversations(raw_conversations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized = [normalize_conversation(record) for record in raw_conversations if isinstance(record, dict)]
    normalized.sort(key=_conversation_sort_key, reverse=True)
    return normalized


def merge_conversation_sources(
    api_conversations: list[dict[str, Any]],
    cached_conversations: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}
    for record in api_conversations:
        if not isinstance(record, dict):
            continue
        normalized = normalize_conversation(record, source_name="api")
        merged[normalized["id"]] = normalized
    for record in cached_conversations:
        if not isinstance(record, dict):
            continue
        normalized = normalize_conversation(record, source_name="cache")
        conversation_id = normalized["id"]
        existing = merged.get(conversation_id)
        if existing is None:
            merged[conversation_id] = normalized
        else:
            merged[conversation_id] = _merge_normalized_conversations(existing, normalized)
    result = list(merged.values())
    result.sort(key=_conversation_sort_key, reverse=True)
    return result


def normalize_conversation(record: dict[str, Any], source_name: str | None = None) -> dict[str, Any]:
    properties = _ensure_dict(record.get("properties"))
    thread_properties = _ensure_dict(record.get("threadProperties"))
    member_properties = _ensure_dict(record.get("memberProperties"))
    last_message = _ensure_dict(record.get("lastMessage"))
    chat_title = _ensure_dict(record.get("chatTitle"))

    conversation_id = _optional_str(record.get("id")) or ""
    title = _conversation_title(conversation_id, thread_properties, chat_title)
    thread_type = _optional_str(thread_properties.get("threadType"))
    product_thread_type = _optional_str(thread_properties.get("productThreadType"))
    hidden = _as_bool(thread_properties.get("hidden"))
    meeting = _looks_like_meeting(conversation_id, title, thread_type, product_thread_type)
    empty = _as_bool(properties.get("isemptyconversation"))
    last_message_at = _coerce_timestamp(
        properties.get("lastimreceivedtime")
        or last_message.get("originalarrivaltime")
        or record.get("lastMessageTimeUtc")
    )
    created_at = _coerce_timestamp(thread_properties.get("createdat"))
    sort_timestamp = last_message_at or created_at

    return {
        "id": conversation_id,
        "title": title,
        "type": _optional_str(record.get("type")),
        "version": _optional_number(record.get("version")),
        "threadType": thread_type,
        "productThreadType": product_thread_type,
        "hidden": hidden,
        "meeting": meeting,
        "empty": empty,
        "sortTimestamp": sort_timestamp,
        "lastMessageAt": last_message_at,
        "createdAt": created_at,
        "lastJoinAtMs": _timestamp_from_epoch_ms(thread_properties.get("lastjoinat")),
        "lastSequenceId": _optional_str(thread_properties.get("lastSequenceId")),
        "originalThreadId": _optional_str(thread_properties.get("originalThreadId")),
        "memberRole": _optional_str(member_properties.get("role")),
        "memberIsReader": _as_bool(member_properties.get("isReader")),
        "memberExpirationTime": _optional_number(member_properties.get("memberExpirationTime")),
        "alerts": _as_bool(properties.get("alerts")),
        "gapDetectionEnabled": _as_bool(thread_properties.get("gapDetectionEnabled")),
        "addedBy": _optional_str(properties.get("addedBy")),
        "addedByTenantId": _optional_str(properties.get("addedByTenantId")),
        "discoverySources": [source_name] if source_name else [],
        "raw": deepcopy(record),
    }


def conversation_filename(conversation_id: str) -> str:
    safe_id = quote(conversation_id, safe="")
    return f"{safe_id}.json"


def _conversation_title(
    conversation_id: str,
    thread_properties: dict[str, Any],
    chat_title: dict[str, Any],
) -> str:
    topic = _optional_str(thread_properties.get("topic"))
    if topic:
        return topic
    short_title = _optional_str(chat_title.get("shortTitle"))
    if short_title:
        return short_title
    long_title = _optional_str(chat_title.get("longTitle"))
    if long_title:
        return long_title
    original_thread_id = _optional_str(thread_properties.get("originalThreadId"))
    if original_thread_id:
        return original_thread_id
    return conversation_id or "Untitled conversation"


def _looks_like_meeting(
    conversation_id: str,
    title: str,
    thread_type: str | None,
    product_thread_type: str | None,
) -> bool:
    fields = [conversation_id, title, thread_type or "", product_thread_type or ""]
    joined = " ".join(fields).lower()
    return "meeting" in joined


def _conversation_sort_key(record: dict[str, Any]) -> tuple[int, str, str]:
    sort_timestamp = _optional_str(record.get("sortTimestamp")) or ""
    last_message_at = _optional_str(record.get("lastMessageAt")) or ""
    title = _optional_str(record.get("title")) or ""
    return (1 if sort_timestamp else 0, sort_timestamp or last_message_at, title.casefold())


def _merge_normalized_conversations(primary: dict[str, Any], secondary: dict[str, Any]) -> dict[str, Any]:
    merged = dict(primary)
    merged["title"] = _prefer_title(primary.get("title"), secondary.get("title"), primary["id"])
    merged["hidden"] = bool(primary.get("hidden")) or bool(secondary.get("hidden"))
    merged["meeting"] = bool(primary.get("meeting")) or bool(secondary.get("meeting"))
    merged["empty"] = bool(primary.get("empty")) and bool(secondary.get("empty"))
    merged["sortTimestamp"] = _max_timestamp(primary.get("sortTimestamp"), secondary.get("sortTimestamp"))
    merged["lastMessageAt"] = _max_timestamp(primary.get("lastMessageAt"), secondary.get("lastMessageAt"))
    merged["createdAt"] = _min_timestamp(primary.get("createdAt"), secondary.get("createdAt"))
    merged["lastSequenceId"] = _prefer_value(primary.get("lastSequenceId"), secondary.get("lastSequenceId"))
    merged["originalThreadId"] = _prefer_value(primary.get("originalThreadId"), secondary.get("originalThreadId"))
    merged["memberRole"] = _prefer_value(primary.get("memberRole"), secondary.get("memberRole"))
    merged["memberIsReader"] = bool(primary.get("memberIsReader")) or bool(secondary.get("memberIsReader"))
    merged["memberExpirationTime"] = _prefer_value(
        primary.get("memberExpirationTime"), secondary.get("memberExpirationTime")
    )
    merged["alerts"] = bool(primary.get("alerts")) or bool(secondary.get("alerts"))
    merged["gapDetectionEnabled"] = bool(primary.get("gapDetectionEnabled")) or bool(
        secondary.get("gapDetectionEnabled")
    )
    merged["addedBy"] = _prefer_value(primary.get("addedBy"), secondary.get("addedBy"))
    merged["addedByTenantId"] = _prefer_value(primary.get("addedByTenantId"), secondary.get("addedByTenantId"))
    merged["discoverySources"] = sorted(
        {str(value) for value in [*(primary.get("discoverySources") or []), *(secondary.get("discoverySources") or [])] if value}
    )
    merged["raw"] = _prefer_raw(primary, secondary)
    return merged


def _prefer_title(primary: Any, secondary: Any, conversation_id: str) -> str:
    primary_text = _optional_str(primary)
    secondary_text = _optional_str(secondary)
    if primary_text and primary_text != conversation_id:
        return primary_text
    if secondary_text and secondary_text != conversation_id:
        return secondary_text
    return primary_text or secondary_text or conversation_id


def _prefer_value(primary: Any, secondary: Any) -> Any:
    if primary not in (None, "", [], {}):
        return primary
    return secondary


def _prefer_raw(primary: dict[str, Any], secondary: dict[str, Any]) -> dict[str, Any] | None:
    if secondary.get("hidden") and not primary.get("hidden"):
        raw = secondary.get("raw")
        return deepcopy(raw) if isinstance(raw, dict) else None
    if _prefer_title(primary.get("title"), secondary.get("title"), primary["id"]) == secondary.get("title"):
        raw = secondary.get("raw")
        return deepcopy(raw) if isinstance(raw, dict) else None
    raw = primary.get("raw")
    if isinstance(raw, dict):
        return deepcopy(raw)
    raw = secondary.get("raw")
    return deepcopy(raw) if isinstance(raw, dict) else None


def _ensure_dict(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value
    return {}


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value)
    return text if text else None


def _optional_number(value: Any) -> int | float | None:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        try:
            return int(value)
        except ValueError:
            try:
                return float(value)
            except ValueError:
                return None
    return None


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes"}
    return False


def _timestamp_from_epoch_ms(value: Any) -> str | None:
    number = _optional_number(value)
    if number is None:
        return None
    return str(int(number))


def _coerce_timestamp(value: Any) -> str | None:
    text = _optional_str(value)
    if text is None:
        return None
    if "T" in text:
        return text
    number = _optional_number(value)
    if number is None:
        return text
    try:
        return datetime.fromtimestamp(float(number) / 1000, tz=UTC).isoformat().replace("+00:00", "Z")
    except (OSError, OverflowError, ValueError):
        return text


def _max_timestamp(first: Any, second: Any) -> str | None:
    timestamps = [value for value in [_coerce_timestamp(first), _coerce_timestamp(second)] if value]
    return max(timestamps) if timestamps else None


def _min_timestamp(first: Any, second: Any) -> str | None:
    timestamps = [value for value in [_coerce_timestamp(first), _coerce_timestamp(second)] if value]
    return min(timestamps) if timestamps else None
