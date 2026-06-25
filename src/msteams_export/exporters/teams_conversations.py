from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import UTC, datetime
import json
import os
from pathlib import Path
import time
from typing import Any

from msteams_export.attachment_stats import summarize_payload_attachments
from msteams_export.browser.session import DEFAULT_TEAMS_URL
from msteams_export.exporters.teams_browser import (
    TeamsBrowserRequest,
    open_teams_page,
    resolve_browser_target,
    run_api_action,
)
from msteams_export.polite_mode import DEFAULT_POLITE_MODE
from msteams_export.exporters.teams_chat import ExportInterrupted, export_live_conversation_document
from msteams_export.parsing.teams_conversations import conversation_filename, merge_conversation_sources


@dataclass(slots=True)
class ConversationListRequest:
    browser_name: str = "auto"
    profile_path: Path | None = None
    teams_url: str = DEFAULT_TEAMS_URL
    headless: bool = True
    timeout_ms: int = 30_000
    output: Path | None = None


@dataclass(slots=True)
class ConversationListResult:
    ok: bool
    message: str
    browser_name: str | None = None
    executable_path: Path | None = None
    profile_path: Path | None = None
    output_path: Path | None = None
    conversation_count: int = 0
    hidden_count: int = 0
    meeting_count: int = 0


@dataclass(slots=True)
class ExportAllRequest:
    outdir: Path
    browser_name: str = "auto"
    profile_path: Path | None = None
    teams_url: str = DEFAULT_TEAMS_URL
    headless: bool = True
    timeout_ms: int = 30_000
    polite_mode: bool = True
    max_chats: int | None = None
    skip_existing: bool = False
    progress: Callable[["ExportProgress"], None] | None = None
    stop_controller: "ExportStopController | None" = None


@dataclass(slots=True)
class ExportStopController:
    interrupt_count: int = 0

    def request_interrupt(self) -> int:
        self.interrupt_count += 1
        return self.interrupt_count

    @property
    def stop_after_current_chat(self) -> bool:
        return self.interrupt_count >= 1

    @property
    def stop_after_current_page(self) -> bool:
        return self.interrupt_count >= 2


@dataclass(slots=True)
class ExportProgress:
    phase: str
    total_conversations: int = 0
    total_hidden: int = 0
    total_meeting: int = 0
    processed_conversations: int = 0
    processed_hidden: int = 0
    processed_meeting: int = 0
    exported_conversations: int = 0
    skipped_conversations: int = 0
    failed_conversations: int = 0
    exported_messages: int = 0
    current_conversation_id: str | None = None
    current_title: str | None = None
    current_hidden: bool = False
    current_meeting: bool = False
    current_message_count: int = 0
    current_status: str | None = None
    note: str | None = None


@dataclass(slots=True)
class ExportAllResult:
    ok: bool
    message: str
    browser_name: str | None = None
    executable_path: Path | None = None
    profile_path: Path | None = None
    outdir: Path | None = None
    index_path: Path | None = None
    conversation_count: int = 0
    exported_count: int = 0
    failed_count: int = 0
    hidden_count: int = 0
    meeting_count: int = 0
    message_count: int = 0
    interrupted: bool = False

def list_conversations(request: ConversationListRequest) -> ConversationListResult:
    try:
        target = resolve_browser_target(
            TeamsBrowserRequest(
                browser_name=request.browser_name,
                profile_path=request.profile_path,
                teams_url=request.teams_url,
                headless=request.headless,
                timeout_ms=request.timeout_ms,
                polite_mode=request.polite_mode,
            )
        )
    except Exception as exc:
        return ConversationListResult(ok=False, message=str(exc))

    try:
        with open_teams_page(target) as page:
            scraped = run_api_action(page, "conversation-list")
    except Exception as exc:
        return ConversationListResult(
            ok=False,
            message=f"Conversation discovery failed: {exc}",
            browser_name=target.browser_name,
            executable_path=target.executable_path,
            profile_path=target.profile_path,
        )

    if not scraped.get("ok", False):
        return ConversationListResult(
            ok=False,
            message=str(scraped.get("error", "Conversation discovery failed.")),
            browser_name=target.browser_name,
            executable_path=target.executable_path,
            profile_path=target.profile_path,
        )

    document = build_conversation_list_document(scraped)
    output_path = None
    if request.output is not None:
        output_path = request.output.expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(document, ensure_ascii=False, indent=2), encoding="utf-8")

    conversations = document["conversations"]
    return ConversationListResult(
        ok=True,
        message=f"Discovered {len(conversations)} conversations.",
        browser_name=target.browser_name,
        executable_path=target.executable_path,
        profile_path=target.profile_path,
        output_path=output_path,
        conversation_count=len(conversations),
        hidden_count=sum(1 for item in conversations if item.get("hidden")),
        meeting_count=sum(1 for item in conversations if item.get("meeting")),
    )


def export_all_conversations(request: ExportAllRequest) -> ExportAllResult:
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
        return ExportAllResult(ok=False, message=str(exc))

    outdir = request.outdir.expanduser().resolve()
    chats_dir = outdir / "chats"
    chats_dir.mkdir(parents=True, exist_ok=True)
    index_path = outdir / "index.json"
    records: list[dict[str, Any]] = []
    total_conversations = 0
    total_hidden = 0
    total_meeting = 0
    processed_conversations = 0
    processed_hidden = 0
    processed_meeting = 0
    exported_count = 0
    skipped_count = 0
    failed_count = 0
    message_count = 0
    interrupted = False
    interruption_note: str | None = None
    scraped: dict[str, Any] = {}

    def emit_progress(
        *,
        phase: str,
        conversation: dict[str, Any] | None = None,
        current_status: str | None = None,
        current_message_count: int = 0,
        note: str | None = None,
    ) -> None:
        if request.progress is None:
            return
        current = conversation or {}
        request.progress(
            ExportProgress(
                phase=phase,
                total_conversations=total_conversations,
                total_hidden=total_hidden,
                total_meeting=total_meeting,
                processed_conversations=processed_conversations,
                processed_hidden=processed_hidden,
                processed_meeting=processed_meeting,
                exported_conversations=exported_count,
                skipped_conversations=skipped_count,
                failed_conversations=failed_count,
                exported_messages=message_count,
                current_conversation_id=_optional_str(current.get("id")),
                current_title=_optional_str(current.get("title")),
                current_hidden=bool(current.get("hidden")),
                current_meeting=bool(current.get("meeting")),
                current_message_count=current_message_count,
                current_status=current_status,
                note=note,
            )
        )

    try:
        with open_teams_page(target) as page:
            scraped = run_api_action(page, "conversation-list")
            if not scraped.get("ok", False):
                return ExportAllResult(
                    ok=False,
                    message=str(scraped.get("error", "Conversation discovery failed.")),
                    browser_name=target.browser_name,
                    executable_path=target.executable_path,
                    profile_path=target.profile_path,
                    outdir=outdir,
                )

            normalized = merge_conversation_sources(
                [item for item in scraped.get("rawConversations", []) if isinstance(item, dict)],
                [item for item in scraped.get("rawCachedConversations", []) if isinstance(item, dict)],
            )
            if request.max_chats is not None:
                normalized = normalized[: max(0, request.max_chats)]

            total_conversations = len(normalized)
            total_hidden = sum(1 for item in normalized if item.get("hidden"))
            total_meeting = sum(1 for item in normalized if item.get("meeting"))
            emit_progress(
                phase="discovering",
                note=(
                    f"Discovered {total_conversations} conversations "
                    f"({total_hidden} hidden, {total_meeting} meeting)."
                ),
            )

            for conversation in normalized:
                if request.stop_controller is not None and request.stop_controller.stop_after_current_chat:
                    interrupted = True
                    interruption_note = "Quit requested. Stopped before starting the next chat."
                    emit_progress(
                        phase="interrupting",
                        conversation=conversation,
                        current_status="interrupting",
                        note=interruption_note,
                    )
                    break

                conversation_id = str(conversation.get("id") or "")
                conversation_hidden = bool(conversation.get("hidden"))
                conversation_meeting = bool(conversation.get("meeting"))
                if not conversation_id:
                    processed_conversations += 1
                    processed_hidden += int(conversation_hidden)
                    processed_meeting += int(conversation_meeting)
                    failed_count += 1
                    records.append(
                        {
                            **conversation,
                            "exported": False,
                            "exportPath": None,
                            "messageCount": 0,
                            "rawCount": 0,
                            "startAt": None,
                            "endAt": None,
                            "timeRange": None,
                            "exportedAt": None,
                            "error": "Conversation record missing id.",
                        }
                    )
                    emit_progress(
                        phase="exporting",
                        conversation=conversation,
                        current_status="failed",
                        note="Conversation record missing id.",
                    )
                    continue

                relative_export_path = Path("chats") / conversation_filename(conversation_id)
                output_path = outdir / relative_export_path
                existing_attachment_stats = _load_attachment_stats(output_path)
                if request.skip_existing and output_path.exists():
                    existing_meta = _load_existing_meta_validated(output_path)
                    existing_count = int(existing_meta.get("count", 0) or 0)
                    existing_end_at = _optional_str(existing_meta.get("endAt"))
                    remote_last_at = _optional_str(conversation.get("lastMessageAt"))
                    # Re-fetch when remote has newer activity than what we already stored.
                    # Compare as datetimes to tolerate fractional-second differences
                    # (e.g. ``.4470000Z`` vs ``.447Z`` represent the same instant).
                    existing_dt = _parse_iso_timestamp(existing_end_at)
                    remote_dt = _parse_iso_timestamp(remote_last_at)
                    if remote_dt is not None and existing_dt is not None:
                        has_new_activity = remote_dt > existing_dt
                    elif remote_dt is not None and existing_dt is None:
                        # We have remote activity but no stored timestamp — re-fetch to be safe.
                        has_new_activity = True
                    else:
                        # No comparable remote timestamp; trust the existing export.
                        has_new_activity = False
                    if has_new_activity:
                        emit_progress(
                            phase="exporting",
                            conversation=conversation,
                            current_status="updating",
                            note="New messages detected — re-fetching.",
                        )
                    else:
                        existing_source = _optional_str(existing_meta.get("source"))
                        existing_target = _optional_str(existing_meta.get("exportTarget"))
                        existing_partial = bool(existing_meta.get("partial")) or existing_target in {
                            "team-space",
                            "community",
                        }
                        existing_warning = _optional_str(existing_meta.get("warning")) or _inferred_warning(
                            source=existing_source,
                            export_target=existing_target,
                        )
                        processed_conversations += 1
                        processed_hidden += int(conversation_hidden)
                        processed_meeting += int(conversation_meeting)
                        records.append(
                            {
                                **conversation,
                                "exported": True,
                                "exportPath": relative_export_path.as_posix(),
                                "messageCount": existing_count,
                                "rawCount": int(existing_meta.get("rawCount", 0) or 0),
                                "startAt": existing_meta.get("startAt"),
                                "endAt": existing_meta.get("endAt"),
                                "timeRange": existing_meta.get("timeRange"),
                                "exportedAt": existing_meta.get("exportedAt"),
                                "source": existing_source,
                                "exportTarget": existing_target,
                                "error": None,
                                "skipped": True,
                                "partial": existing_partial,
                                "warning": existing_warning,
                                **existing_attachment_stats,
                            }
                        )
                        exported_count += 1
                        skipped_count += 1
                        message_count += existing_count
                        emit_progress(
                            phase="exporting",
                            conversation=conversation,
                            current_status="skipped",
                            current_message_count=existing_count,
                            note="Reused existing export file.",
                        )
                        continue

                try:
                    document = export_live_conversation_document(
                        page,
                        conversation_id=conversation_id,
                        conversation_title=str(conversation.get("title") or conversation_id),
                        conversation_meta=conversation,
                        stop_controller=request.stop_controller,
                    )
                    output_path.parent.mkdir(parents=True, exist_ok=True)
                    chat_content = json.dumps(document, ensure_ascii=False, indent=2)
                    chat_temp = output_path.with_suffix(".tmp.json")
                    chat_temp.write_text(chat_content, encoding="utf-8")
                    os.replace(str(chat_temp), str(output_path))
                    meta = document["meta"]
                    exported_messages = int(meta.get("count", 0) or 0)
                    export_source = _optional_str(meta.get("source"))
                    export_target = _optional_str(meta.get("exportTarget"))
                    export_partial = bool(meta.get("partial")) or export_target in {"team-space", "community"}
                    export_warning = _optional_str(meta.get("warning")) or _inferred_warning(
                        source=export_source,
                        export_target=export_target,
                    )
                    attachment_stats = summarize_payload_attachments(document)
                    processed_conversations += 1
                    processed_hidden += int(conversation_hidden)
                    processed_meeting += int(conversation_meeting)
                    records.append(
                        {
                            **conversation,
                            "exported": True,
                            "exportPath": relative_export_path.as_posix(),
                            "messageCount": exported_messages,
                            "rawCount": int(meta.get("rawCount", 0) or 0),
                            "startAt": meta.get("startAt"),
                            "endAt": meta.get("endAt"),
                            "timeRange": meta.get("timeRange"),
                            "exportedAt": meta.get("exportedAt"),
                            "source": export_source,
                            "exportTarget": export_target,
                            "error": None,
                            "partial": export_partial,
                            "warning": export_warning,
                            **attachment_stats,
                        }
                    )
                    exported_count += 1
                    message_count += exported_messages
                    emit_progress(
                        phase="exporting",
                        conversation=conversation,
                        current_status="exported",
                        current_message_count=exported_messages,
                    )
                    if request.polite_mode:
                        time.sleep(DEFAULT_POLITE_MODE.conversation_spacing_ms / 1000)
                except ExportInterrupted as exc:
                    interrupted = True
                    interruption_note = str(exc)
                    failed_count += 1
                    records.append(
                        {
                            **conversation,
                            "exported": False,
                            "exportPath": relative_export_path.as_posix(),
                            "messageCount": 0,
                            "rawCount": exc.fetched_messages,
                            "startAt": None,
                            "endAt": None,
                            "timeRange": None,
                            "exportedAt": None,
                            "error": str(exc),
                            "interrupted": True,
                            **existing_attachment_stats,
                        }
                    )
                    _write_index_atomic(index_path, records, user_region=_optional_str(scraped.get("userRegion")))
                    emit_progress(
                        phase="interrupting",
                        conversation=conversation,
                        current_status="interrupted",
                        current_message_count=exc.fetched_messages,
                        note=str(exc),
                    )
                    break
                except Exception as exc:
                    processed_conversations += 1
                    processed_hidden += int(conversation_hidden)
                    processed_meeting += int(conversation_meeting)
                    failed_count += 1
                    records.append(
                        {
                            **conversation,
                            "exported": False,
                            "exportPath": relative_export_path.as_posix(),
                            "messageCount": 0,
                            "rawCount": 0,
                            "startAt": None,
                            "endAt": None,
                            "timeRange": None,
                            "exportedAt": None,
                            "error": str(exc),
                            **existing_attachment_stats,
                        }
                    )
                    emit_progress(
                        phase="exporting",
                        conversation=conversation,
                        current_status="failed",
                        note=str(exc),
                    )
                    if request.polite_mode:
                        time.sleep(DEFAULT_POLITE_MODE.conversation_spacing_ms / 1000)

    except Exception as exc:
        if records:
            try:
                _write_index_atomic(index_path, records, user_region=_optional_str(scraped.get("userRegion")))
            except Exception:
                pass  # Best-effort — don't mask original error
        return ExportAllResult(
            ok=False,
            message=f"Export all failed: {exc}",
            browser_name=target.browser_name,
            executable_path=target.executable_path,
            profile_path=target.profile_path,
            outdir=outdir,
        )

    _write_index_atomic(index_path, records, user_region=_optional_str(scraped.get("userRegion")))
    emit_progress(
        phase="interrupted" if interrupted else "finished",
        current_status="interrupted" if interrupted else "finished",
        note=interruption_note or "Export index written.",
    )

    if interrupted:
        detail = f" {interruption_note}" if interruption_note else ""
        message = (
            f"Export interrupted after {processed_conversations} conversations. "
            f"Partial index written to {index_path}.{detail}"
        )
    else:
        message = f"Exported {exported_count} conversations to {outdir}"

    return ExportAllResult(
        ok=not interrupted,
        message=message,
        browser_name=target.browser_name,
        executable_path=target.executable_path,
        profile_path=target.profile_path,
        outdir=outdir,
        index_path=index_path,
        conversation_count=len(records),
        exported_count=exported_count,
        failed_count=failed_count,
        hidden_count=sum(1 for item in records if item.get("hidden")),
        meeting_count=sum(1 for item in records if item.get("meeting")),
        message_count=message_count,
        interrupted=interrupted,
    )


def build_conversation_list_document(scraped: dict[str, Any]) -> dict[str, Any]:
    normalized = merge_conversation_sources(
        [item for item in scraped.get("rawConversations", []) if isinstance(item, dict)],
        [item for item in scraped.get("rawCachedConversations", []) if isinstance(item, dict)],
    )
    exported_at = datetime.now(UTC).isoformat().replace("+00:00", "Z")
    return {
        "meta": {
            "title": "Teams conversations index",
            "conversationCount": len(normalized),
            "userRegion": scraped.get("userRegion"),
            "source": "teams-web-api",
            "exportTarget": "conversation-list",
            "exportedAt": exported_at,
            "rawCount": scraped.get("rawCount", len(normalized)),
            "pageCount": scraped.get("pageCount"),
            "cacheCount": scraped.get("cacheCount"),
            "cacheHiddenCount": scraped.get("cacheHiddenCount"),
            "hiddenCount": sum(1 for item in normalized if item.get("hidden")),
            "meetingCount": sum(1 for item in normalized if item.get("meeting")),
        },
        "conversations": normalized,
    }


def build_export_index_document(
    records: list[dict[str, Any]],
    *,
    user_region: str | None = None,
) -> dict[str, Any]:
    exported_at = datetime.now(UTC).isoformat().replace("+00:00", "Z")
    start_at = _first_non_empty(record.get("startAt") for record in records)
    end_at = _last_non_empty(record.get("endAt") for record in records)
    return {
        "meta": {
            "title": "Teams chat export index",
            "conversationCount": len(records),
            "exportedCount": sum(1 for item in records if item.get("exported")),
            "failedCount": sum(1 for item in records if item.get("error")),
            "partialCount": sum(1 for item in records if item.get("partial")),
            "metadataOnlyCount": sum(
                1
                for item in records
                if item.get("exportTarget") in {"team-space", "community"}
            ),
            "hiddenCount": sum(1 for item in records if item.get("hidden")),
            "meetingCount": sum(1 for item in records if item.get("meeting")),
            "messageCount": sum(int(item.get("messageCount", 0) or 0) for item in records),
            "attachmentCount": sum(int(item.get("assetCount", 0) or 0) for item in records),
            "mirroredAttachmentCount": sum(int(item.get("mirroredAssetCount", 0) or 0) for item in records),
            "attachmentFailureCount": sum(int(item.get("assetFailureCount", 0) or 0) for item in records),
            "unauthorizedAttachmentCount": sum(int(item.get("unauthorizedAssetCount", 0) or 0) for item in records),
            "offlineReadyConversationCount": sum(1 for item in records if item.get("offlineReady")),
            "startAt": start_at,
            "endAt": end_at,
            "timeRange": _format_time_range(start_at, end_at),
            "userRegion": user_region,
            "source": "teams-web-api",
            "exportTarget": "all-chats",
            "exportedAt": exported_at,
        },
        "conversations": records,
    }


def _load_existing_meta(path: Path) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(payload, dict):
        return {}
    meta = payload.get("meta")
    if not isinstance(meta, dict):
        return {}
    return meta


def _load_attachment_stats(path: Path) -> dict[str, int | bool]:
    if not path.is_file():
        return {
            "assetCount": 0,
            "mirroredAssetCount": 0,
            "assetFailureCount": 0,
            "offlineReady": True,
        }
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {
            "assetCount": 0,
            "mirroredAssetCount": 0,
            "assetFailureCount": 0,
            "offlineReady": True,
        }
    if not isinstance(payload, dict):
        return {
            "assetCount": 0,
            "mirroredAssetCount": 0,
            "assetFailureCount": 0,
            "offlineReady": True,
        }
    return summarize_payload_attachments(payload)


def _inferred_warning(*, source: str | None, export_target: str | None) -> str | None:
    if export_target == "team-space" or source == "teams-team-metadata":
        return "Metadata-only export. Team space was discovered, but full channel history is not currently accessible from this session."
    if export_target == "community" or source == "teams-community-metadata":
        return "Metadata-only export. Community metadata was discovered, but full post history is not currently accessible from this session."
    return None


def _first_non_empty(values: Any) -> str | None:
    filtered = [value for value in values if isinstance(value, str) and value]
    return min(filtered) if filtered else None


def _last_non_empty(values: Any) -> str | None:
    filtered = [value for value in values if isinstance(value, str) and value]
    return max(filtered) if filtered else None


def _format_time_range(start_at: str | None, end_at: str | None) -> str | None:
    if start_at and end_at:
        return f"{start_at} -> {end_at}"
    if start_at:
        return f"from {start_at}"
    if end_at:
        return f"until {end_at}"
    return None


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value)
    return text if text else None


def _parse_iso_timestamp(value: str | None) -> datetime | None:
    """Parse an ISO-8601 timestamp, tolerating varying fractional-second precision.

    Teams returns fractional seconds with anywhere from 0 to 7 digits depending on
    the API surface (Skype IC3 vs. graph). Plain string comparison gives wrong
    results across surfaces (e.g. ``.4470000Z`` vs ``.447Z``), so normalize to a
    datetime before comparing.
    """
    if not value:
        return None
    text = value.strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    if "." in text:
        head, rest = text.split(".", 1)
        # Split fractional digits from any trailing timezone offset.
        tz_index = len(rest)
        for marker in ("+", "-"):
            idx = rest.find(marker)
            if idx != -1 and idx < tz_index:
                tz_index = idx
        frac = rest[:tz_index]
        tail = rest[tz_index:]
        frac_digits = "".join(ch for ch in frac if ch.isdigit())[:6]
        if frac_digits:
            text = f"{head}.{frac_digits}{tail}"
        else:
            text = f"{head}{tail}"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=UTC)
    return parsed


def _write_index_atomic(
    index_path: Path,
    records: list[dict[str, Any]],
    *,
    user_region: str | None = None,
) -> None:
    """Write index.json atomically via temp file + rename."""
    index_document = build_export_index_document(records, user_region=user_region)
    content = json.dumps(index_document, ensure_ascii=False, indent=2)
    temp_path = index_path.with_suffix(".tmp.json")
    temp_path.write_text(content, encoding="utf-8")
    os.replace(str(temp_path), str(index_path))


def _load_existing_meta_validated(path: Path) -> dict[str, Any]:
    """Load and validate an existing export file. Returns empty dict if invalid."""
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(payload, dict):
        return {}
    meta = payload.get("meta")
    if not isinstance(meta, dict):
        return {}
    messages = payload.get("messages")
    if not isinstance(messages, list):
        return {}
    count = int(meta.get("count", 0) or 0)
    if count > 0 and len(messages) == 0:
        return {}
    return meta
