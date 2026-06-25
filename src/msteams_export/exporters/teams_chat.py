from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime
import json
from pathlib import Path
from typing import Any

from msteams_export.browser.session import DEFAULT_TEAMS_URL
from msteams_export.exporters.teams_browser import (
    TeamsBrowserRequest,
    open_teams_page,
    resolve_browser_target,
    run_api_action,
)
from msteams_export.parsing.teams_api import convert_api_messages
from msteams_export.parsing.teams_graph import convert_graph_channel_messages
from msteams_export.parsing.teams_conversations import merge_conversation_sources, normalize_conversation


@dataclass(slots=True)
class ChatExportRequest:
    output: Path
    browser_name: str = "auto"
    profile_path: Path | None = None
    teams_url: str = DEFAULT_TEAMS_URL
    headless: bool = True
    timeout_ms: int = 30_000
    conversation_id: str | None = None
    conversation_title: str | None = None


@dataclass(slots=True)
class ChatExportResult:
    ok: bool
    message: str
    output_path: Path | None = None
    browser_name: str | None = None
    executable_path: Path | None = None
    profile_path: Path | None = None
    title: str | None = None
    conversation_id: str | None = None
    message_count: int = 0


class ExportInterrupted(RuntimeError):
    def __init__(self, message: str, *, fetched_messages: int = 0) -> None:
        super().__init__(message)
        self.fetched_messages = fetched_messages


def export_chat_to_json(request: ChatExportRequest) -> ChatExportResult:
    try:
        target = resolve_browser_target(
            TeamsBrowserRequest(
                browser_name=request.browser_name,
                profile_path=request.profile_path,
                teams_url=request.teams_url,
                headless=request.headless,
                timeout_ms=request.timeout_ms,
            )
        )
    except Exception as exc:
        return ChatExportResult(ok=False, message=str(exc))

    try:
        with open_teams_page(target) as page:
            if request.conversation_id:
                list_payload = run_api_action(page, "conversation-list")
                normalized = merge_conversation_sources(
                    [item for item in list_payload.get("rawConversations", []) if isinstance(item, dict)],
                    [item for item in list_payload.get("rawCachedConversations", []) if isinstance(item, dict)],
                )
                matched_conversation = next(
                    (item for item in normalized if item.get("id") == request.conversation_id),
                    None,
                )
                document = export_live_conversation_document(
                    page,
                    conversation_id=request.conversation_id,
                    conversation_title=request.conversation_title
                    or _optional_str(matched_conversation.get("title") if matched_conversation else None)
                    or request.conversation_id,
                    conversation_meta=matched_conversation,
                )
            else:
                scraped = run_api_action(page, "active-chat")
                if not scraped.get("ok", False):
                    return ChatExportResult(
                        ok=False,
                        message=str(scraped.get("error", "Chat export failed.")),
                        browser_name=target.browser_name,
                        executable_path=target.executable_path,
                        profile_path=target.profile_path,
                    )
                document = build_export_document(scraped)
    except Exception as exc:
        return ChatExportResult(
            ok=False,
            message=f"Export probe failed: {exc}",
            browser_name=target.browser_name,
            executable_path=target.executable_path,
            profile_path=target.profile_path,
        )

    output_path = request.output.expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(document, ensure_ascii=False, indent=2), encoding="utf-8")

    meta = document["meta"]
    messages = document["messages"]
    return ChatExportResult(
        ok=True,
        message=f"Exported {len(messages)} messages to {output_path}",
        output_path=output_path,
        browser_name=target.browser_name,
        executable_path=target.executable_path,
        profile_path=target.profile_path,
        title=str(meta.get("title", "") or ""),
        conversation_id=str(meta.get("conversationId", "") or ""),
        message_count=len(messages),
    )


def export_live_conversation_document(
    page: Any,
    *,
    conversation_id: str,
    conversation_title: str,
    conversation_meta: dict[str, Any] | None = None,
    stop_controller: Any = None,
) -> dict[str, Any]:
    meta = conversation_meta or {}
    if _is_team_space(meta):
        return build_team_space_document(meta, title=conversation_title)
    if _is_community(meta):
        return build_community_document(meta, title=conversation_title)
    if _is_channel(meta):
        errors: list[str] = []
        thread_url = _thread_url(meta)
        if thread_url:
            try:
                payload = _fetch_thread_messages_payload(
                    page,
                    thread_url=thread_url,
                    conversation_id=conversation_id,
                    title=conversation_title,
                    stop_controller=stop_controller,
                )
                payload["conversationMeta"] = meta
                payload["conversation"] = meta.get("raw")
                payload["source"] = "teams-thread-api"
                payload["exportTarget"] = "channel"
                return build_export_document(payload)
            except Exception as exc:
                errors.append(f"thread: {exc}")
        team_id = _team_group_id(meta)
        if team_id:
            try:
                payload = _fetch_graph_channel_payload(
                    page,
                    team_id=team_id,
                    channel_id=conversation_id,
                    title=conversation_title,
                    stop_controller=stop_controller,
                )
                payload["conversationMeta"] = meta
                payload["conversation"] = meta.get("raw")
                return build_channel_export_document(payload)
            except Exception as exc:
                errors.append(f"graph: {exc}")
        return build_channel_fallback_document(
            meta,
            title=conversation_title,
            warning=" | ".join(errors) or "Live channel history endpoint was not accessible.",
        )

    payload = (
        _fetch_chat_messages_paged(
            page,
            conversation_id=conversation_id,
            title=conversation_title,
            stop_controller=stop_controller,
        )
        if stop_controller is not None
        else run_api_action(
            page,
            "conversation-messages",
            conversationId=conversation_id,
            title=conversation_title,
        )
    )
    if not payload.get("ok", False):
        raise RuntimeError(str(payload.get("error", "Conversation export failed.")))
    payload["conversationMeta"] = meta
    payload["conversation"] = meta.get("raw")
    return build_export_document(payload)


def build_export_document(scraped: dict[str, Any]) -> dict[str, Any]:
    raw_messages = scraped.get("rawMessages", [])
    normalized_raw_messages = [message for message in raw_messages if isinstance(message, dict)]
    conversation_meta = _ensure_dict(scraped.get("conversationMeta"))
    if conversation_meta:
        conversation = conversation_meta
    else:
        conversation_raw = _ensure_dict(scraped.get("conversation"))
        conversation = normalize_conversation(conversation_raw) if conversation_raw else None
    conversation_id = _optional_str(scraped.get("conversationId")) or _optional_str(
        conversation.get("id") if conversation else None
    )
    messages = convert_api_messages(
        normalized_raw_messages,
        conversation_id=conversation_id,
    )
    title = (
        _optional_str(scraped.get("title"))
        or _optional_str(conversation.get("title") if conversation else None)
        or "Teams Export"
    )
    start_at = _first_timestamp(messages)
    end_at = _last_timestamp(messages)
    exported_at = datetime.now(UTC).isoformat().replace("+00:00", "Z")
    meta = {
        "title": title,
        "count": len(messages),
        "startAt": start_at,
        "endAt": end_at,
        "timeRange": _format_time_range(start_at, end_at),
        "conversationId": conversation_id,
        "userRegion": scraped.get("userRegion"),
        "source": scraped.get("source") or "teams-web-api",
        "exportTarget": scraped.get("exportTarget") or "chat",
        "exportedAt": exported_at,
        "rawCount": scraped.get("rawCount", len(normalized_raw_messages)),
    }
    if conversation is not None:
        meta["conversation"] = conversation
    return {"meta": meta, "messages": messages}


def build_channel_export_document(scraped: dict[str, Any]) -> dict[str, Any]:
    raw_messages = scraped.get("rawMessages", [])
    normalized_raw_messages = [message for message in raw_messages if isinstance(message, dict)]
    conversation_meta = _ensure_dict(scraped.get("conversationMeta"))
    if conversation_meta:
        conversation = conversation_meta
    else:
        conversation_raw = _ensure_dict(scraped.get("conversation"))
        conversation = normalize_conversation(conversation_raw) if conversation_raw else None
    conversation_id = _optional_str(scraped.get("conversationId")) or _optional_str(
        conversation.get("id") if conversation else None
    )
    messages = convert_graph_channel_messages(normalized_raw_messages, conversation_id=conversation_id)
    title = (
        _optional_str(scraped.get("title"))
        or _optional_str(conversation.get("title") if conversation else None)
        or "Teams Channel Export"
    )
    start_at = _first_timestamp(messages)
    end_at = _last_timestamp(messages)
    exported_at = datetime.now(UTC).isoformat().replace("+00:00", "Z")
    meta = {
        "title": title,
        "count": len(messages),
        "startAt": start_at,
        "endAt": end_at,
        "timeRange": _format_time_range(start_at, end_at),
        "conversationId": conversation_id,
        "userRegion": scraped.get("userRegion"),
        "source": scraped.get("source") or "graph-teams-channel",
        "exportTarget": scraped.get("exportTarget") or "channel",
        "exportedAt": exported_at,
        "rawCount": scraped.get("rawCount", len(normalized_raw_messages)),
    }
    if conversation is not None:
        meta["conversation"] = conversation
    return {"meta": meta, "messages": messages}


def build_team_space_document(conversation_meta: dict[str, Any], *, title: str | None = None) -> dict[str, Any]:
    conversation = conversation_meta or {}
    exported_at = datetime.now(UTC).isoformat().replace("+00:00", "Z")
    meta = {
        "title": title or _optional_str(conversation.get("title")) or "Teams Team",
        "count": 0,
        "startAt": None,
        "endAt": None,
        "timeRange": None,
        "conversationId": _optional_str(conversation.get("id")),
        "userRegion": None,
        "source": "teams-team-metadata",
        "exportTarget": "team-space",
        "exportedAt": exported_at,
        "rawCount": 0,
        "partial": True,
        "warning": "Metadata-only export. Team space was discovered, but full channel history is not currently accessible from this session.",
        "teamChannels": _team_topics(conversation),
    }
    if conversation:
        meta["conversation"] = conversation
    return {"meta": meta, "messages": []}


def build_community_document(conversation_meta: dict[str, Any], *, title: str | None = None) -> dict[str, Any]:
    conversation = conversation_meta or {}
    raw = _ensure_dict(conversation.get("raw"))
    thread_properties = _ensure_dict(raw.get("threadProperties"))
    exported_at = datetime.now(UTC).isoformat().replace("+00:00", "Z")
    meta = {
        "title": title or _optional_str(conversation.get("title")) or "Engage Community",
        "count": 0,
        "startAt": None,
        "endAt": None,
        "timeRange": None,
        "conversationId": _optional_str(conversation.get("id")),
        "userRegion": None,
        "source": "teams-community-metadata",
        "exportTarget": "community",
        "exportedAt": exported_at,
        "rawCount": 0,
        "partial": True,
        "warning": "Metadata-only export. Community metadata was discovered, but full post history is not currently accessible from this session.",
        "engageCommunityId": _optional_str(thread_properties.get("engageCommunityId")),
    }
    if conversation:
        meta["conversation"] = conversation
    return {"meta": meta, "messages": []}


def build_channel_fallback_document(
    conversation_meta: dict[str, Any],
    *,
    title: str | None = None,
    warning: str,
) -> dict[str, Any]:
    conversation = conversation_meta or {}
    raw = _ensure_dict(conversation.get("raw"))
    last_message = _ensure_dict(raw.get("lastMessage"))
    conversation_id = _optional_str(conversation.get("id"))
    messages = convert_api_messages([last_message], conversation_id=conversation_id) if last_message else []
    exported_at = datetime.now(UTC).isoformat().replace("+00:00", "Z")
    start_at = _first_timestamp(messages)
    end_at = _last_timestamp(messages)
    meta = {
        "title": title or _optional_str(conversation.get("title")) or "Teams Channel",
        "count": len(messages),
        "startAt": start_at,
        "endAt": end_at,
        "timeRange": _format_time_range(start_at, end_at),
        "conversationId": conversation_id,
        "userRegion": None,
        "source": "teams-channel-fallback",
        "exportTarget": "channel-fallback",
        "exportedAt": exported_at,
        "rawCount": len(messages),
        "partial": True,
        "warning": warning,
    }
    if conversation:
        meta["conversation"] = conversation
    return {"meta": meta, "messages": messages}


def _fetch_chat_messages_paged(
    page: Any,
    *,
    conversation_id: str,
    title: str,
    stop_controller: Any = None,
) -> dict[str, Any]:
    all_messages: list[dict[str, Any]] = []
    next_url: str | None = None
    page_count = 0
    user_region: str | None = None

    while True:
        if stop_controller is not None and getattr(stop_controller, "stop_after_current_page", False) and next_url is None:
            raise ExportInterrupted(
                "Quit requested before starting the next page. Current chat was not written.",
                fetched_messages=0,
            )

        payload = run_api_action(
            page,
            "conversation-messages-page",
            conversationId=conversation_id,
            title=title,
            nextUrl=next_url,
        )
        if not payload.get("ok", False):
            raise RuntimeError(str(payload.get("error", "Conversation export failed.")))

        page_count += 1
        user_region = _optional_str(payload.get("userRegion")) or user_region
        raw_messages = [item for item in payload.get("rawMessages", []) if isinstance(item, dict)]
        all_messages.extend(raw_messages)
        next_url = _optional_str(payload.get("nextUrl"))

        if not next_url:
            return {
                "ok": True,
                "userRegion": user_region,
                "conversationId": conversation_id,
                "title": title,
                "rawMessages": all_messages,
                "rawCount": len(all_messages),
                "pageCount": page_count,
            }

        if stop_controller is not None and getattr(stop_controller, "stop_after_current_page", False):
            raise ExportInterrupted(
                "Force-quit requested. Stopped after the current Teams page and wrote a partial index.",
                fetched_messages=len(all_messages),
            )


def _fetch_thread_messages_payload(
    page: Any,
    *,
    thread_url: str,
    conversation_id: str,
    title: str,
    stop_controller: Any = None,
) -> dict[str, Any]:
    all_messages: list[dict[str, Any]] = []
    next_url: str | None = None
    page_count = 0
    user_region: str | None = None

    while True:
        payload = run_api_action(
            page,
            "thread-messages-page",
            threadUrl=thread_url,
            nextUrl=next_url,
        )
        if not payload.get("ok", False):
            raise RuntimeError(str(payload.get("error", "Thread export failed.")))

        page_count += 1
        user_region = _optional_str(payload.get("userRegion")) or user_region
        raw_messages = [item for item in payload.get("rawMessages", []) if isinstance(item, dict)]
        all_messages.extend(raw_messages)
        next_url = _optional_str(payload.get("nextUrl"))

        if not next_url:
            return {
                "ok": True,
                "userRegion": user_region,
                "conversationId": conversation_id,
                "title": title,
                "rawMessages": all_messages,
                "rawCount": len(all_messages),
                "pageCount": page_count,
            }

        if stop_controller is not None and getattr(stop_controller, "stop_after_current_page", False):
            raise ExportInterrupted(
                "Force-quit requested. Stopped after the current Teams thread page and wrote a partial index.",
                fetched_messages=len(all_messages),
            )


def _fetch_graph_channel_payload(
    page: Any,
    *,
    team_id: str,
    channel_id: str,
    title: str,
    stop_controller: Any = None,
) -> dict[str, Any]:
    all_messages: list[dict[str, Any]] = []
    next_url: str | None = None
    page_count = 0
    user_region: str | None = None

    while True:
        payload = run_api_action(
            page,
            "graph-channel-messages-page",
            teamId=team_id,
            channelId=channel_id,
            nextUrl=next_url,
        )
        if not payload.get("ok", False):
            raise RuntimeError(str(payload.get("error", "Channel export failed.")))

        page_count += 1
        user_region = _optional_str(payload.get("userRegion")) or user_region
        root_messages = [item for item in payload.get("rawMessages", []) if isinstance(item, dict)]
        for root in root_messages:
            root["replies"] = _fetch_graph_channel_replies(
                page,
                team_id=team_id,
                channel_id=channel_id,
                message_id=str(root.get("id") or ""),
                stop_controller=stop_controller,
            )
            all_messages.append(root)
        next_url = _optional_str(payload.get("nextUrl"))

        if not next_url:
            return {
                "ok": True,
                "userRegion": user_region,
                "conversationId": channel_id,
                "title": title,
                "rawMessages": all_messages,
                "rawCount": len(all_messages),
                "pageCount": page_count,
            }

        if stop_controller is not None and getattr(stop_controller, "stop_after_current_page", False):
            raise ExportInterrupted(
                "Force-quit requested. Stopped after the current Graph page and wrote a partial index.",
                fetched_messages=len(all_messages),
            )


def _fetch_graph_channel_replies(
    page: Any,
    *,
    team_id: str,
    channel_id: str,
    message_id: str,
    stop_controller: Any = None,
) -> list[dict[str, Any]]:
    if not message_id:
        return []
    replies: list[dict[str, Any]] = []
    next_url: str | None = None
    while True:
        payload = run_api_action(
            page,
            "graph-channel-message-replies-page",
            teamId=team_id,
            channelId=channel_id,
            messageId=message_id,
            nextUrl=next_url,
        )
        if not payload.get("ok", False):
            raise RuntimeError(str(payload.get("error", "Channel reply export failed.")))
        replies.extend(item for item in payload.get("rawReplies", []) if isinstance(item, dict))
        next_url = _optional_str(payload.get("nextUrl"))
        if not next_url:
            return replies
        if stop_controller is not None and getattr(stop_controller, "stop_after_current_page", False):
            raise ExportInterrupted(
                "Force-quit requested. Stopped after the current Graph reply page and wrote a partial index.",
                fetched_messages=len(replies),
            )


def _first_timestamp(messages: list[dict[str, Any]]) -> str | None:
    for message in messages:
        timestamp = message.get("timestamp")
        if isinstance(timestamp, str) and timestamp:
            return timestamp
    return None


def _last_timestamp(messages: list[dict[str, Any]]) -> str | None:
    for message in reversed(messages):
        timestamp = message.get("timestamp")
        if isinstance(timestamp, str) and timestamp:
            return timestamp
    return None


def _format_time_range(start_at: str | None, end_at: str | None) -> str | None:
    if start_at and end_at:
        return f"{start_at} -> {end_at}"
    if start_at:
        return f"from {start_at}"
    if end_at:
        return f"until {end_at}"
    return None


def _is_channel(conversation_meta: dict[str, Any]) -> bool:
    product_thread_type = str(conversation_meta.get("productThreadType") or "")
    thread_type = str(conversation_meta.get("threadType") or "")
    return product_thread_type == "TeamsStandardChannel" or thread_type == "topic"


def _is_team_space(conversation_meta: dict[str, Any]) -> bool:
    product_thread_type = str(conversation_meta.get("productThreadType") or "")
    thread_type = str(conversation_meta.get("threadType") or "")
    return product_thread_type == "TeamsTeam" or thread_type == "space"


def _is_community(conversation_meta: dict[str, Any]) -> bool:
    thread_type = str(conversation_meta.get("threadType") or "")
    return thread_type == "engagecommunity"


def _team_group_id(conversation_meta: dict[str, Any]) -> str | None:
    raw = _ensure_dict(conversation_meta.get("raw"))
    thread_properties = _ensure_dict(raw.get("threadProperties"))
    return _optional_str(thread_properties.get("groupId")) or _optional_str(raw.get("teamId"))


def _thread_url(conversation_meta: dict[str, Any]) -> str | None:
    raw = _ensure_dict(conversation_meta.get("raw"))
    return _optional_str(raw.get("targetLink"))


def _team_topics(conversation_meta: dict[str, Any]) -> list[dict[str, Any]]:
    raw = _ensure_dict(conversation_meta.get("raw"))
    thread_properties = _ensure_dict(raw.get("threadProperties"))
    raw_topics = thread_properties.get("topics")
    if not raw_topics:
        return []
    try:
        parsed = json.loads(raw_topics) if isinstance(raw_topics, str) else raw_topics
    except json.JSONDecodeError:
        return []
    if not isinstance(parsed, list):
        return []
    topics: list[dict[str, Any]] = []
    for item in parsed:
        if not isinstance(item, dict):
            continue
        topics.append(
            {
                "id": _optional_str(item.get("id")),
                "name": _optional_str(item.get("name")) or "Channel",
                "createdAt": _optional_str(item.get("createdat")),
                "deleted": bool(item.get("isdeleted")),
            }
        )
    return topics


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value)
    return text if text else None


def _ensure_dict(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value
    return {}
