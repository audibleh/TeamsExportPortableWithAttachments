from __future__ import annotations

import csv
from dataclasses import dataclass
import hashlib
import html
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import io
import json
import mimetypes
from pathlib import Path
from pathlib import PurePosixPath
import re
import secrets
import threading
import time
from typing import Any, Callable
import unicodedata
from urllib.parse import parse_qs, quote, urlparse
import webbrowser
import zipfile

from msteams_export import __version__
from msteams_export.attachment_policy import filter_attachment_dicts
from msteams_export.attachment_stats import summarize_payload_attachments
from msteams_export.browser.session import DEFAULT_TEAMS_URL
from msteams_export.bundle_paths import resolve_bundle_relative_path
from msteams_export.export_bundle import resolve_export_root
from msteams_export.exporters.attachment_mirror import (
    AttachmentMirrorProgress,
    MirrorAttachmentsRequest,
    MirrorAttachmentsResult,
    MirrorStopController,
    mirror_bundle_attachments,
)
from msteams_export.exporters.teams_conversations import (
    ExportAllRequest,
    ExportAllResult,
    ExportProgress,
    ExportStopController,
    export_all_conversations,
)
from msteams_export.parsing.teams_api import (
    is_system_message_type,
    merge_embedded_html_attachments,
    parse_system_content,
)
from msteams_export.webapp.attachments import (
    build_viewer_attachment_href,
    fetch_attachment,
    is_inline_content_type,
    normalize_attachment_url,
)

GLOBAL_SEARCH_RESULT_LIMIT = 200


@dataclass(slots=True)
class ViewerServeRequest:
    target: Path
    host: str = "127.0.0.1"
    port: int = 8765
    open_browser: bool = False
    browser_name: str = "auto"
    profile_path: Path | None = None
    teams_url: str = DEFAULT_TEAMS_URL
    timeout_ms: int = 30_000


@dataclass(slots=True)
class ViewerServeResult:
    ok: bool
    message: str
    url: str | None = None
    index_path: Path | None = None
    chats_dir: Path | None = None


class ExportRepository:
    def __init__(self, index_path: Path, chats_dir: Path) -> None:
        self._lock = threading.RLock()
        self.index_path = index_path
        self.chats_dir = chats_dir
        self.reload()

    def reload(self) -> None:
        with self._lock:
            payload = json.loads(self.index_path.read_text(encoding="utf-8"))
            if not isinstance(payload, dict):
                raise ValueError(f"Expected JSON object in {self.index_path}")
            self.index_payload = payload
            self.meta = payload.get("meta", {}) if isinstance(payload.get("meta"), dict) else {}
            conversations = payload.get("conversations", [])
            if not isinstance(conversations, list):
                raise ValueError("Expected 'conversations' to be an array in index.json")
            self._conversations = [item for item in conversations if isinstance(item, dict)]
            self._by_id = {
                str(item.get("id")): item
                for item in self._conversations
                if item.get("id") is not None and str(item.get("id"))
            }
            self._attachment_summary_cache: dict[str, int] | None = None
            self._space_title_by_id = _build_space_title_index(self._conversations)
            self._channel_parent_by_id = _build_channel_parent_index(
                self._conversations,
                self._space_title_by_id,
            )

    def summary(self) -> dict[str, Any]:
        with self._lock:
            failed = [item for item in self._conversations if item.get("error")]
            partial = [item for item in self._conversations if item.get("partial")]
            metadata_only = [
                item for item in self._conversations if item.get("exportTarget") in {"team-space", "community"}
            ]
            message_count = _meta_or_sum(self.meta.get("messageCount"), self._conversations, key="messageCount")
            derived = self._derive_attachment_summary_from_chats()
            attachment_count = derived["attachmentCount"]
            mirrored_attachment_count = derived["mirroredAttachmentCount"]
            attachment_failure_count = derived["attachmentFailureCount"]
            offline_ready_count = derived["offlineReadyConversationCount"]
            unauthorized_attachment_count = derived["unauthorizedAttachmentCount"]
            summary_meta = dict(self.meta)
            summary_meta["messageCount"] = message_count
            summary_meta["attachmentCount"] = attachment_count
            summary_meta["mirroredAttachmentCount"] = mirrored_attachment_count
            summary_meta["attachmentFailureCount"] = attachment_failure_count
            summary_meta["offlineReadyConversationCount"] = offline_ready_count
            summary_meta["unauthorizedAttachmentCount"] = unauthorized_attachment_count
            self._persist_attachment_summary(summary_meta)
            return {
                "meta": summary_meta,
                "conversationCount": len(self._conversations),
                "hiddenCount": sum(1 for item in self._conversations if item.get("hidden")),
                "meetingCount": sum(1 for item in self._conversations if item.get("meeting")),
                "messageCount": message_count,
                "failedCount": len(failed),
                "partialCount": len(partial),
                "metadataOnlyCount": len(metadata_only),
                "attachmentCount": attachment_count,
                "mirroredAttachmentCount": mirrored_attachment_count,
                "attachmentFailureCount": attachment_failure_count,
                "offlineReadyConversationCount": offline_ready_count,
                "unauthorizedAttachmentCount": unauthorized_attachment_count,
                "kindBreakdown": _count_by(self._conversations, key="kind"),
                "failedKindBreakdown": _count_by(failed, key="kind"),
                "bundleRoot": str(self.index_path.parent.resolve()),
                "indexPath": str(self.index_path),
                "chatsDir": str(self.chats_dir),
            }

    def _derive_attachment_summary_from_chats(self) -> dict[str, int]:
        with self._lock:
            if self._attachment_summary_cache is not None:
                return dict(self._attachment_summary_cache)
            bundle_root = self.index_path.parent.resolve()
            totals = {
                "attachmentCount": 0,
                "mirroredAttachmentCount": 0,
                "attachmentFailureCount": 0,
                "offlineReadyConversationCount": 0,
                "unauthorizedAttachmentCount": 0,
            }
            for record in self._conversations:
                export_path = record.get("exportPath")
                if not isinstance(export_path, str) or not export_path:
                    continue
                chat_path = resolve_bundle_relative_path(bundle_root, export_path)
                if chat_path is None or not chat_path.is_file():
                    continue
                try:
                    payload = json.loads(chat_path.read_text(encoding="utf-8"))
                except Exception:
                    continue
                if not isinstance(payload, dict):
                    continue
                stats = summarize_payload_attachments(payload)
                totals["attachmentCount"] += int(stats.get("assetCount", 0) or 0)
                totals["mirroredAttachmentCount"] += int(stats.get("mirroredAssetCount", 0) or 0)
                totals["attachmentFailureCount"] += int(stats.get("assetFailureCount", 0) or 0)
                totals["offlineReadyConversationCount"] += int(bool(stats.get("offlineReady")))
                totals["unauthorizedAttachmentCount"] += int(stats.get("unauthorizedAssetCount", 0) or 0)
            self._attachment_summary_cache = totals
            return dict(totals)

    def _persist_attachment_summary(self, summary_meta: dict[str, Any]) -> None:
        payload = dict(self.index_payload)
        payload["meta"] = dict(summary_meta)
        current_meta = self.index_payload.get("meta", {})
        if current_meta == payload["meta"]:
            return
        self.index_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.index_payload = payload
        self.meta = payload["meta"]

    def _enrich_conversation_summary(self, summary: dict[str, Any], item: dict[str, Any] | None) -> dict[str, Any]:
        record = item if isinstance(item, dict) else {}
        enriched = dict(summary)
        display_title = _conversation_display_title(record)
        if display_title:
            enriched["displayTitle"] = display_title
        channel_title = _conversation_channel_title(record)
        if channel_title and channel_title != display_title:
            enriched["channelTitle"] = channel_title
        conversation_id = _optional_str(record.get("id"))
        if conversation_id:
            parent = self._channel_parent_by_id.get(conversation_id)
            if isinstance(parent, dict):
                parent_title = _optional_str(parent.get("title"))
                parent_id = _optional_str(parent.get("id"))
                if parent_title:
                    enriched["parentSpaceTitle"] = parent_title
                if parent_id:
                    enriched["parentSpaceId"] = parent_id
        return enriched

    def list_conversations(
        self,
        *,
        query: str = "",
        case_sensitive: bool = False,
        hidden: str = "any",
        meeting: str = "any",
        status: str = "any",
        kind: str = "any",
        exported_only: bool = False,
    ) -> list[dict[str, Any]]:
        with self._lock:
            query_matcher = _build_text_matcher(query, case_sensitive=case_sensitive)
            results: list[dict[str, Any]] = []
            for item in self._conversations:
                if exported_only and not item.get("exportPath"):
                    continue
                if hidden != "any" and bool(item.get("hidden")) != (hidden == "true"):
                    continue
                if meeting != "any" and bool(item.get("meeting")) != (meeting == "true"):
                    continue
                summary = self._enrich_conversation_summary(_conversation_summary(item), item)
                if status != "any" and summary["status"] != status:
                    continue
                if kind != "any" and summary["kind"] != kind:
                    continue
                haystack = " ".join(
                    [
                        str(summary.get("displayTitle") or ""),
                        str(summary.get("channelTitle") or ""),
                        str(summary.get("parentSpaceTitle") or ""),
                        str(item.get("title") or ""),
                        str(item.get("id") or ""),
                        str(item.get("threadType") or ""),
                        str(item.get("productThreadType") or ""),
                        str(item.get("error") or ""),
                    ]
                )
                if query_matcher is not None and not query_matcher(haystack):
                    continue
                results.append(summary)
            return results

    def search_messages(
        self,
        *,
        query: str,
        conversation_query: str = "",
        case_sensitive: bool = False,
        hidden: str = "any",
        meeting: str = "any",
        status: str = "any",
        kind: str = "any",
        exported_only: bool = False,
        hide_system: bool = False,
        limit: int | None = None,
    ) -> list[dict[str, Any]]:
        content_matcher = _build_text_matcher(query, case_sensitive=case_sensitive)
        if content_matcher is None:
            return self.list_conversations(
                query=conversation_query,
                case_sensitive=case_sensitive,
                hidden=hidden,
                meeting=meeting,
                status=status,
                kind=kind,
                exported_only=exported_only,
            )
        base = self.list_conversations(
            query=conversation_query,
            case_sensitive=case_sensitive,
            hidden=hidden,
            meeting=meeting,
            status=status,
            kind=kind,
            exported_only=exported_only,
        )
        results: list[dict[str, Any]] = []
        for summary in base:
            conversation_id = _optional_str(summary.get("id"))
            if conversation_id is None:
                continue
            try:
                payload = self.get_chat(conversation_id)
            except Exception:
                continue
            messages = payload.get("messages", [])
            if not isinstance(messages, list):
                continue
            match_count = 0
            latest_match_at = ""
            preview_text = ""
            preview_author = ""
            preview_message_id = ""
            for item in messages:
                if not isinstance(item, dict):
                    continue
                presentation = _message_presentation(item)
                if hide_system and presentation["system"]:
                    continue
                text = str(presentation["text"] or "")
                if not content_matcher(text):
                    continue
                match_count += 1
                timestamp = str(item.get("timestamp") or "")
                if timestamp and timestamp >= latest_match_at:
                    latest_match_at = timestamp
                if not preview_text:
                    preview_text = text
                    preview_author = str(presentation["author"] or "")
                    preview_message_id = str(item.get("id") or "")
            if match_count <= 0:
                continue
            record = dict(summary)
            record["matchCount"] = match_count
            record["latestMatchAt"] = latest_match_at
            record["matchPreview"] = preview_text
            record["matchPreviewAuthor"] = preview_author
            record["matchPreviewMessageId"] = preview_message_id
            results.append(record)
        results.sort(
            key=lambda item: (
                int(item.get("matchCount", 0) or 0),
                str(item.get("latestMatchAt") or ""),
                str(item.get("displayTitle") or item.get("title") or item.get("id") or ""),
            ),
            reverse=True,
        )
        if limit is not None and limit >= 0:
            return results[:limit]
        return results

    def get_chat(self, conversation_id: str) -> dict[str, Any]:
        with self._lock:
            record = self._by_id.get(conversation_id)
            if record is None:
                raise KeyError(f"Unknown conversation id: {conversation_id}")
            chat_path = self._conversation_path(record)
        payload = json.loads(chat_path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise ValueError(f"Expected JSON object in {chat_path}")
        messages = payload.get("messages")
        if isinstance(messages, list):
            sanitized_messages: list[dict[str, Any]] = []
            for item in messages:
                if not isinstance(item, dict):
                    continue
                message = dict(item)
                message["attachments"] = filter_attachment_dicts(
                    merge_embedded_html_attachments(item.get("attachments"), _optional_str(item.get("contentHtml")))
                )
                sanitized_messages.append(message)
            payload["messages"] = sanitized_messages
        return payload

    def get_chat_payload(
        self,
        conversation_id: str,
        *,
        query: str = "",
        author: str = "",
        case_sensitive: bool = False,
        hide_system: bool = False,
        limit: int = 300,
        offset: int = 0,
    ) -> dict[str, Any]:
        payload = self.get_chat(conversation_id)
        messages = payload.get("messages", [])
        if not isinstance(messages, list):
            raise ValueError("Expected 'messages' array in chat export.")

        query_matcher = _build_text_matcher(query, case_sensitive=case_sensitive)
        author_matcher = _build_text_matcher(author, case_sensitive=case_sensitive)
        filtered = []
        for item in messages:
            if not isinstance(item, dict):
                continue
            presentation = _message_presentation(item)
            if hide_system and presentation["system"]:
                continue
            author_text = str(presentation["author"] or "")
            if author_matcher is not None and not author_matcher(author_text):
                continue
            text = str(presentation["text"] or "")
            if query_matcher is not None and not query_matcher(text):
                continue
            filtered.append(item)

        start = max(0, offset)
        stop = start + max(1, limit)
        window = [_decorate_message(item, conversation_id=conversation_id) for item in filtered[start:stop]]
        return {
            "meta": payload.get("meta", {}),
            "conversation": self._enrich_conversation_summary(
                _conversation_summary(self._by_id.get(conversation_id, {})),
                self._by_id.get(conversation_id, {}),
            ),
            "messages": window,
            "totalMatches": len(filtered),
            "returned": len(window),
            "offset": start,
            "limit": limit,
        }

    def export_conversation(self, conversation_id: str, fmt: str) -> tuple[bytes, str, str]:
        payload = self.get_chat(conversation_id)
        with self._lock:
            conversation = self._by_id.get(conversation_id, {})
        title = str(payload.get("meta", {}).get("title") or conversation.get("title") or conversation_id)
        stem = _safe_stem(title)
        export_payload, export_assets = _prepare_export_payload(self, payload, conversation_id=conversation_id)
        if fmt == "csv":
            content = _chat_to_csv(export_payload)
            return content.encode("utf-8"), "text/csv; charset=utf-8", f"{stem}.csv"
        if fmt == "md":
            content = _chat_to_markdown(export_payload)
            return content.encode("utf-8"), "text/markdown; charset=utf-8", f"{stem}.md"
        if fmt == "html":
            bundle_name = f"{stem}-html-bundle"
            html_payload, html_assets = _prepare_html_bundle_payload(export_payload, export_assets)
            content = _chat_to_html_bundle(
                html_payload,
                archive_root=bundle_name,
                bundled_assets=html_assets,
            )
            return content, "application/zip", f"{bundle_name}.zip"
        raise ValueError(f"Unsupported export format: {fmt}")

    def get_attachment(self, conversation_id: str, message_id: str, attachment_index: int) -> dict[str, Any]:
        payload = self.get_chat(conversation_id)
        messages = payload.get("messages", [])
        if not isinstance(messages, list):
            raise ValueError("Expected 'messages' array in chat export.")
        for item in messages:
            if not isinstance(item, dict):
                continue
            if str(item.get("id") or "") != message_id:
                continue
            attachments = item.get("attachments", [])
            if not isinstance(attachments, list):
                break
            if attachment_index < 0 or attachment_index >= len(attachments):
                raise IndexError(f"Attachment index {attachment_index} is out of range for message {message_id}.")
            attachment = attachments[attachment_index]
            if not isinstance(attachment, dict):
                raise ValueError(f"Attachment {attachment_index} in message {message_id} is not a JSON object.")
            href = str(attachment.get("href") or "").strip()
            if not href:
                raise ValueError(f"Attachment {attachment_index} in message {message_id} does not have a URL.")
            return attachment
        raise KeyError(f"Unknown message id: {message_id}")

    def resolve_local_attachment(self, attachment: dict[str, Any]) -> Path | None:
        local_path = attachment.get("localPath")
        if not isinstance(local_path, str) or not local_path.strip():
            return None
        bundle_root = self.index_path.parent.resolve()
        candidate = resolve_bundle_relative_path(bundle_root, local_path)
        if candidate is None:
            return None
        return candidate if candidate.is_file() else None

    def _conversation_path(self, record: dict[str, Any]) -> Path:
        export_path = record.get("exportPath")
        if isinstance(export_path, str) and export_path:
            candidate = resolve_bundle_relative_path(self.index_path.parent.resolve(), export_path)
            if candidate is None:
                raise FileNotFoundError(f"Conversation file missing for {record.get('id')}")
            if candidate.is_file():
                return candidate
        raise FileNotFoundError(f"Conversation file missing for {record.get('id')}")


class AttachmentMirrorJobManager:
    def __init__(self, repo: ExportRepository, request: ViewerServeRequest) -> None:
        self.repo = repo
        self.request = request
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._snapshot: AttachmentMirrorProgress | None = None
        self._result: MirrorAttachmentsResult | None = None
        self._stop_controller: MirrorStopController | None = None
        self._started_at: float | None = None

    def status(self) -> dict[str, Any]:
        with self._lock:
            running = self._thread is not None and self._thread.is_alive()
            snapshot = self._snapshot
            result = self._result
            started_at = self._started_at
        return {
            "running": running,
            "startedAt": started_at,
            "snapshot": _serialize_attachment_mirror_progress(snapshot),
            "result": _serialize_attachment_mirror_result(result),
        }

    def start(self, *, min_free_gb: float = 30.0) -> dict[str, Any]:
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return {"ok": False, "message": "Attachment mirror is already running.", "started": False}
            self._snapshot = None
            self._result = None
            self._stop_controller = MirrorStopController()
            self._started_at = time.time()
            thread = threading.Thread(
                target=self._run_job,
                args=(max(0, int(float(min_free_gb) * 1024 * 1024 * 1024)),),
                daemon=True,
                name="attachment-mirror-job",
            )
            self._thread = thread
            thread.start()
        return {"ok": True, "message": "Attachment mirror started.", "started": True}

    def stop(self, *, force: bool = False) -> dict[str, Any]:
        with self._lock:
            controller = self._stop_controller
            running = self._thread is not None and self._thread.is_alive()
        if not running or controller is None:
            return {"ok": False, "message": "Attachment mirror is not running."}
        level = controller.request_interrupt()
        if force:
            level = controller.request_interrupt()
        if level <= 1:
            return {"ok": True, "message": "Stop requested. The mirror will finish the current chat first."}
        return {"ok": True, "message": "Force-stop requested. The mirror will stop after the current attachment."}

    def _run_job(self, min_free_bytes: int) -> None:
        result = mirror_bundle_attachments(
            MirrorAttachmentsRequest(
                target=self.repo.index_path,
                browser_name=self.request.browser_name,
                profile_path=self.request.profile_path,
                teams_url=self.request.teams_url,
                timeout_ms=self.request.timeout_ms,
                min_free_bytes=min_free_bytes,
                progress=self._on_progress,
                stop_controller=self._stop_controller,
            )
        )
        self.repo.reload()
        with self._lock:
            self._result = result
            self._thread = None
            self._stop_controller = None

    def _on_progress(self, snapshot: AttachmentMirrorProgress) -> None:
        with self._lock:
            self._snapshot = snapshot


class ExportAllJobManager:
    def __init__(self, repo: ExportRepository, request: ViewerServeRequest) -> None:
        self.repo = repo
        self.request = request
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._snapshot: ExportProgress | None = None
        self._result: ExportAllResult | None = None
        self._stop_controller: ExportStopController | None = None
        self._started_at: float | None = None

    def status(self) -> dict[str, Any]:
        with self._lock:
            running = self._thread is not None and self._thread.is_alive()
            snapshot = self._snapshot
            result = self._result
            started_at = self._started_at
        return {
            "running": running,
            "startedAt": started_at,
            "snapshot": _serialize_export_progress(snapshot),
            "result": _serialize_export_result(result),
        }

    def start(self, *, skip_existing: bool = True, max_chats: int | None = None) -> dict[str, Any]:
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return {"ok": False, "message": "Export is already running.", "started": False}
            self._snapshot = None
            self._result = None
            self._stop_controller = ExportStopController()
            self._started_at = time.time()
            thread = threading.Thread(
                target=self._run_job,
                args=(skip_existing, max_chats),
                daemon=True,
                name="export-all-job",
            )
            self._thread = thread
            thread.start()
        return {"ok": True, "message": "Export started.", "started": True}

    def stop(self, *, force: bool = False) -> dict[str, Any]:
        with self._lock:
            controller = self._stop_controller
            running = self._thread is not None and self._thread.is_alive()
        if not running or controller is None:
            return {"ok": False, "message": "Export is not running."}
        level = controller.request_interrupt()
        if force:
            level = controller.request_interrupt()
        if level <= 1:
            return {"ok": True, "message": "Stop requested. The export will finish the current chat first."}
        return {"ok": True, "message": "Force-stop requested. The export will stop after the current Teams page."}

    def _run_job(self, skip_existing: bool, max_chats: int | None) -> None:
        result = export_all_conversations(
            ExportAllRequest(
                outdir=self.repo.index_path.parent,
                browser_name=self.request.browser_name,
                profile_path=self.request.profile_path,
                teams_url=self.request.teams_url,
                headless=True,
                timeout_ms=self.request.timeout_ms,
                max_chats=max_chats,
                skip_existing=skip_existing,
                progress=self._on_progress,
                stop_controller=self._stop_controller,
            )
        )
        self.repo.reload()
        with self._lock:
            self._result = result
            self._thread = None
            self._stop_controller = None

    def _on_progress(self, snapshot: ExportProgress) -> None:
        with self._lock:
            self._snapshot = snapshot


class ViewerHTTPServer(ThreadingHTTPServer):
    def __init__(
        self,
        server_address: tuple[str, int],
        handler_class: type[BaseHTTPRequestHandler],
        repo: ExportRepository,
        request: ViewerServeRequest,
    ) -> None:
        self.repo = repo
        self.viewer_request = request
        self.control_token = secrets.token_urlsafe(24)
        self.build_label = __version__
        self.export_jobs = ExportAllJobManager(repo, request)
        self.mirror_jobs = AttachmentMirrorJobManager(repo, request)
        super().__init__(server_address, handler_class)


class ViewerRequestHandler(BaseHTTPRequestHandler):
    server: ViewerHTTPServer

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._write_html(_app_html(self.server.control_token, self.server.build_label))
            return
        if parsed.path == "/api/meta":
            self._write_json(self.server.repo.summary())
            return
        if parsed.path == "/api/control/export/status":
            self._write_json(self.server.export_jobs.status())
            return
        if parsed.path == "/api/control/mirror/status":
            self._write_json(self.server.mirror_jobs.status())
            return
        if parsed.path == "/api/conversations":
            params = parse_qs(parsed.query)
            try:
                items = self.server.repo.list_conversations(
                    query=_first_param(params, "query"),
                    case_sensitive=_first_param(params, "caseSensitive", "false") == "true",
                    hidden=_first_param(params, "hidden", "any"),
                    meeting=_first_param(params, "meeting", "any"),
                    status=_first_param(params, "status", "any"),
                    kind=_first_param(params, "kind", "any"),
                    exported_only=_first_param(params, "exported", "true") == "true",
                )
            except ValueError as exc:
                self._write_json({"error": str(exc), "conversations": [], "count": 0}, status=HTTPStatus.BAD_REQUEST)
                return
            self._write_json({"conversations": items, "count": len(items)})
            return
        if parsed.path == "/api/search":
            params = parse_qs(parsed.query)
            requested_limit = _int_param(params, "limit", GLOBAL_SEARCH_RESULT_LIMIT)
            try:
                items = self.server.repo.search_messages(
                    query=_first_param(params, "query"),
                    conversation_query=_first_param(params, "conversationQuery"),
                    case_sensitive=_first_param(params, "caseSensitive", "false") == "true",
                    hidden=_first_param(params, "hidden", "any"),
                    meeting=_first_param(params, "meeting", "any"),
                    status=_first_param(params, "status", "any"),
                    kind=_first_param(params, "kind", "any"),
                    exported_only=_first_param(params, "exported", "true") == "true",
                    hide_system=_first_param(params, "hideSystem", "false") == "true",
                    limit=requested_limit,
                )
            except ValueError as exc:
                self._write_json({"error": str(exc), "conversations": [], "count": 0}, status=HTTPStatus.BAD_REQUEST)
                return
            self._write_json(
                {
                    "conversations": items,
                    "count": len(items),
                    "searchMode": "content",
                    "limit": requested_limit,
                    "truncated": requested_limit >= 0 and len(items) >= requested_limit,
                }
            )
            return
        if parsed.path == "/api/chat":
            params = parse_qs(parsed.query)
            conversation_id = _first_param(params, "conversationId")
            if not conversation_id:
                self._write_json({"error": "conversationId is required"}, status=HTTPStatus.BAD_REQUEST)
                return
            try:
                payload = self.server.repo.get_chat_payload(
                    conversation_id,
                    query=_first_param(params, "query"),
                    author=_first_param(params, "author"),
                    case_sensitive=_first_param(params, "caseSensitive", "false") == "true",
                    hide_system=_first_param(params, "hideSystem", "false") == "true",
                    limit=_int_param(params, "limit", 300),
                    offset=_int_param(params, "offset", 0),
                )
            except Exception as exc:
                self._write_json({"error": str(exc)}, status=HTTPStatus.NOT_FOUND)
                return
            self._write_json(payload)
            return
        if parsed.path == "/api/export":
            params = parse_qs(parsed.query)
            conversation_id = _first_param(params, "conversationId")
            fmt = _first_param(params, "format", "md")
            if not conversation_id:
                self._write_json({"error": "conversationId is required"}, status=HTTPStatus.BAD_REQUEST)
                return
            try:
                body, content_type, filename = self.server.repo.export_conversation(conversation_id, fmt)
            except Exception as exc:
                self._write_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
                return
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(filename)}")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path == "/api/attachment":
            params = parse_qs(parsed.query)
            conversation_id = _first_param(params, "conversationId")
            message_id = _first_param(params, "messageId")
            attachment_index = _int_param(params, "attachmentIndex", -1)
            if not conversation_id or not message_id or attachment_index < 0:
                self._write_text(
                    "conversationId, messageId and attachmentIndex are required.",
                    status=HTTPStatus.BAD_REQUEST,
                )
                return
            try:
                attachment = self.server.repo.get_attachment(conversation_id, message_id, attachment_index)
                local_path = self.server.repo.resolve_local_attachment(attachment)
                if local_path is not None:
                    body = local_path.read_bytes()
                    content_type = (
                        mimetypes.guess_type(local_path.name)[0]
                        or str(attachment.get("localContentType") or "")
                        or "application/octet-stream"
                    )
                    disposition = "inline" if is_inline_content_type(content_type) else "attachment"
                    self.send_response(HTTPStatus.OK)
                    self.send_header("Content-Type", content_type)
                    self.send_header("Content-Disposition", f"{disposition}; filename*=UTF-8''{quote(local_path.name)}")
                    self.send_header("Content-Length", str(len(body)))
                    self.end_headers()
                    self.wfile.write(body)
                    return
                result = fetch_attachment(
                    url=str(attachment.get("href") or ""),
                    label=str(attachment.get("label") or "") or None,
                    browser_name=self.server.viewer_request.browser_name,
                    profile_path=self.server.viewer_request.profile_path,
                    teams_url=self.server.viewer_request.teams_url,
                    timeout_ms=self.server.viewer_request.timeout_ms,
                )
            except Exception as exc:
                self._write_text(f"Could not fetch attachment: {exc}", status=HTTPStatus.BAD_GATEWAY)
                return
            disposition = "inline" if is_inline_content_type(result.content_type) else "attachment"
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", result.content_type)
            self.send_header("Content-Disposition", f"{disposition}; filename*=UTF-8''{quote(result.filename)}")
            self.send_header("Content-Length", str(len(result.body)))
            self.end_headers()
            self.wfile.write(result.body)
            return
        self._write_json({"error": "Not found"}, status=HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if not self._is_control_request_authorized():
            self._write_json({"error": "Forbidden"}, status=HTTPStatus.FORBIDDEN)
            return
        payload = self._read_json_body()

        if parsed.path == "/api/control/export/start":
            if self.server.mirror_jobs.status().get("running"):
                self._write_json(
                    {"ok": False, "message": "Attachment mirror is running. Wait for it to finish before starting export."},
                    status=HTTPStatus.CONFLICT,
                )
                return
            skip_existing = bool(payload.get("skipExisting", True))
            max_chats = payload.get("maxChats")
            if max_chats is not None:
                try:
                    max_chats = max(0, int(max_chats))
                except (TypeError, ValueError):
                    max_chats = None
            result = self.server.export_jobs.start(skip_existing=skip_existing, max_chats=max_chats)
            self._write_json(result, status=HTTPStatus.ACCEPTED if result.get("ok") else HTTPStatus.CONFLICT)
            return
        if parsed.path == "/api/control/export/stop":
            result = self.server.export_jobs.stop(force=bool(payload.get("force", False)))
            self._write_json(result, status=HTTPStatus.ACCEPTED if result.get("ok") else HTTPStatus.CONFLICT)
            return
        if parsed.path == "/api/control/mirror/start":
            if self.server.export_jobs.status().get("running"):
                self._write_json(
                    {"ok": False, "message": "Export is running. Wait for it to finish before starting attachment mirror."},
                    status=HTTPStatus.CONFLICT,
                )
                return
            min_free_gb = payload.get("minFreeGb", 30.0)
            try:
                min_free_gb = max(0.0, float(min_free_gb))
            except (TypeError, ValueError):
                min_free_gb = 30.0
            result = self.server.mirror_jobs.start(min_free_gb=min_free_gb)
            self._write_json(result, status=HTTPStatus.ACCEPTED if result.get("ok") else HTTPStatus.CONFLICT)
            return
        if parsed.path == "/api/control/mirror/stop":
            result = self.server.mirror_jobs.stop(force=bool(payload.get("force", False)))
            self._write_json(result, status=HTTPStatus.ACCEPTED if result.get("ok") else HTTPStatus.CONFLICT)
            return
        self._write_json({"error": "Not found"}, status=HTTPStatus.NOT_FOUND)

    def log_message(self, format: str, *args: Any) -> None:
        return

    def _write_json(self, payload: dict[str, Any], *, status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _write_html(self, body: str) -> None:
        encoded = body.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def _write_text(self, body: str, *, status: HTTPStatus = HTTPStatus.OK) -> None:
        encoded = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Cache-Control", "no-store, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def _read_json_body(self) -> dict[str, Any]:
        try:
            length = int(self.headers.get("Content-Length", "0") or "0")
        except ValueError:
            length = 0
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        try:
            payload = json.loads(raw.decode("utf-8"))
        except Exception:
            return {}
        return payload if isinstance(payload, dict) else {}

    def _is_control_request_authorized(self) -> bool:
        token = self.headers.get("X-Control-Token", "")
        return bool(token) and token == self.server.control_token and self._is_same_origin_control_request()

    def _is_same_origin_control_request(self) -> bool:
        host = _optional_str(self.headers.get("Host"))
        if host is None:
            server_host, server_port = self.server.server_address[:2]
            host = f"{server_host}:{server_port}"
        origin = _optional_str(self.headers.get("Origin"))
        referer = _optional_str(self.headers.get("Referer"))
        if origin and not _origin_matches_host(host, origin):
            return False
        if referer and not _origin_matches_host(host, referer):
            return False
        return True


def serve_viewer(request: ViewerServeRequest) -> ViewerServeResult:
    index_path, chats_dir = resolve_export_root(request.target)
    repo = ExportRepository(index_path=index_path, chats_dir=chats_dir)
    server = ViewerHTTPServer((request.host, request.port), ViewerRequestHandler, repo, request)
    host, port = server.server_address
    url = f"http://{host}:{port}/"
    if request.open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return ViewerServeResult(
        ok=True,
        message="Viewer server stopped.",
        url=url,
        index_path=index_path,
        chats_dir=chats_dir,
    )


def _conversation_summary(item: dict[str, Any]) -> dict[str, Any]:
    status = (
        "failed" if item.get("error")
        else "metadata-only" if item.get("exportTarget") in {"team-space", "community"}
        else "partial" if item.get("partial")
        else "skipped" if item.get("skipped")
        else "exported" if item.get("exported", True)
        else "pending"
    )
    kind = _conversation_kind(item)
    return {
        "id": item.get("id"),
        "title": item.get("title"),
        "kind": kind,
        "status": status,
        "warning": item.get("warning"),
        "hidden": bool(item.get("hidden")),
        "meeting": bool(item.get("meeting")),
        "messageCount": int(item.get("messageCount", 0) or 0),
        "startAt": item.get("startAt"),
        "endAt": item.get("endAt"),
        "exportPath": item.get("exportPath"),
        "threadType": item.get("threadType"),
        "productThreadType": item.get("productThreadType"),
        "source": item.get("source"),
        "exportTarget": item.get("exportTarget"),
        "error": item.get("error"),
        "discoverySources": item.get("discoverySources", []),
    }


def _conversation_display_title(item: dict[str, Any]) -> str | None:
    kind = _conversation_kind(item)
    thread_properties = _raw_thread_properties(item)
    if kind == "team":
        return (
            _optional_str(thread_properties.get("spaceThreadTopic"))
            or _optional_str(item.get("title"))
            or _optional_str(item.get("id"))
        )
    return _optional_str(item.get("title")) or _optional_str(item.get("id"))


def _conversation_channel_title(item: dict[str, Any]) -> str | None:
    if _conversation_kind(item) != "team":
        return None
    thread_properties = _raw_thread_properties(item)
    return (
        _optional_str(item.get("title"))
        or _optional_str(thread_properties.get("topic"))
        or _optional_str(item.get("id"))
    )


def _build_text_matcher(value: str, *, case_sensitive: bool) -> Callable[[str], bool] | None:
    text = value.strip()
    if not text:
        return None
    regex = _parse_regex_literal(text, case_sensitive=case_sensitive)
    if regex is not None:
        return lambda candidate: bool(regex.search(candidate))
    if case_sensitive:
        return lambda candidate: text in candidate
    folded = text.casefold()
    return lambda candidate: folded in candidate.casefold()


def _parse_regex_literal(value: str, *, case_sensitive: bool) -> re.Pattern[str] | None:
    if len(value) < 2 or not value.startswith("/"):
        return None
    closing_index = _find_regex_closing_slash(value)
    if closing_index <= 0:
        return None
    pattern = value[1:closing_index]
    flags_text = value[closing_index + 1 :]
    if "/" in flags_text:
        return None
    flags = 0 if case_sensitive else re.IGNORECASE
    for char in flags_text:
        if char == "i":
            flags &= ~re.IGNORECASE
            flags |= re.IGNORECASE
            continue
        if char == "m":
            flags |= re.MULTILINE
            continue
        if char == "s":
            flags |= re.DOTALL
            continue
        raise ValueError(f"Unsupported regex flag: {char}")
    try:
        return re.compile(pattern, flags)
    except re.error as exc:
        raise ValueError(f"Invalid regex: {exc}") from exc


def _find_regex_closing_slash(value: str) -> int:
    escaped = False
    for index in range(1, len(value)):
        char = value[index]
        if escaped:
            escaped = False
            continue
        if char == "\\":
            escaped = True
            continue
        if char == "/":
            return index
    return -1


def _conversation_kind(item: dict[str, Any]) -> str:
    product_thread_type = str(item.get("productThreadType") or "")
    thread_type = str(item.get("threadType") or "")
    if product_thread_type == "TeamsStandardChannel" or thread_type == "topic":
        return "channel"
    if product_thread_type == "TeamsTeam" or thread_type == "space":
        return "team"
    if thread_type == "engagecommunity":
        return "community"
    if bool(item.get("meeting")):
        return "meeting-chat"
    return "chat"


def _count_by(items: list[dict[str, Any]], *, key: str) -> dict[str, int]:
    counts: dict[str, int] = {}
    for item in items:
        value = _conversation_summary(item).get(key) if key in {"kind", "status"} else item.get(key)
        label = str(value or "unknown")
        counts[label] = counts.get(label, 0) + 1
    return counts


def _build_space_title_index(conversations: list[dict[str, Any]]) -> dict[str, str]:
    titles: dict[str, str] = {}
    for item in conversations:
        if _conversation_kind(item) != "team":
            continue
        conversation_id = _optional_str(item.get("id"))
        title = _conversation_display_title(item)
        if conversation_id and title:
            titles[conversation_id] = title
    return titles


def _build_channel_parent_index(
    conversations: list[dict[str, Any]],
    space_title_by_id: dict[str, str],
) -> dict[str, dict[str, str]]:
    parent_by_channel_id: dict[str, dict[str, str]] = {}

    for item in conversations:
        if _conversation_kind(item) != "team":
            continue
        parent_id = _optional_str(item.get("id"))
        parent_title = space_title_by_id.get(parent_id or "")
        if not parent_id or not parent_title:
            continue
        for topic in _thread_topics(item):
            topic_id = _optional_str(topic.get("id"))
            if topic_id:
                parent_by_channel_id[topic_id] = {"id": parent_id, "title": parent_title}

    for item in conversations:
        if _conversation_kind(item) != "channel":
            continue
        channel_id = _optional_str(item.get("id"))
        if not channel_id or channel_id in parent_by_channel_id:
            continue
        raw = _raw_conversation(item)
        thread_properties = _raw_thread_properties(item)
        parent_id = (
            _optional_str(raw.get("teamId"))
            or _optional_str(thread_properties.get("spaceId"))
            or _optional_str(thread_properties.get("teamId"))
        )
        parent_title = space_title_by_id.get(parent_id or "")
        if parent_id and parent_title:
            parent_by_channel_id[channel_id] = {"id": parent_id, "title": parent_title}

    return parent_by_channel_id


def _raw_conversation(item: dict[str, Any]) -> dict[str, Any]:
    raw = item.get("raw")
    return raw if isinstance(raw, dict) else {}


def _raw_thread_properties(item: dict[str, Any]) -> dict[str, Any]:
    thread_properties = _raw_conversation(item).get("threadProperties")
    return thread_properties if isinstance(thread_properties, dict) else {}


def _thread_topics(item: dict[str, Any]) -> list[dict[str, Any]]:
    raw_topics = _raw_thread_properties(item).get("topics")
    if isinstance(raw_topics, list):
        return [topic for topic in raw_topics if isinstance(topic, dict)]
    if isinstance(raw_topics, str):
        try:
            parsed = json.loads(raw_topics)
        except json.JSONDecodeError:
            return []
        if isinstance(parsed, list):
            return [topic for topic in parsed if isinstance(topic, dict)]
    return []


def _meta_or_sum(meta_value: Any, items: list[dict[str, Any]], *, key: str, truthy: bool = False) -> int:
    coerced = _safe_int(meta_value)
    if coerced is not None:
        return coerced
    if truthy:
        return sum(1 for item in items if item.get(key))
    return sum(int(item.get(key, 0) or 0) for item in items)


def _safe_int(value: Any) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text if text else None


def _message_presentation(message: dict[str, Any]) -> dict[str, Any]:
    message_type = _optional_str(message.get("messageType")) or ""
    system = bool(message.get("system")) or is_system_message_type(message_type)
    raw_text = str(message.get("text") or "")
    content = _optional_str(message.get("contentHtml")) or raw_text
    display_text = parse_system_content(content, message_type) if system else raw_text
    author = _optional_str(message.get("author")) or ("[system]" if system else "[unknown]")
    return {
        "author": author,
        "text": display_text or raw_text or "[no text]",
        "system": system,
    }


def _decorate_message(message: dict[str, Any], *, conversation_id: str) -> dict[str, Any]:
    decorated = dict(message)
    presentation = _message_presentation(message)
    decorated["author"] = presentation["author"]
    decorated["text"] = presentation["text"]
    decorated["system"] = presentation["system"]
    attachments = message.get("attachments", [])
    message_id = str(message.get("id") or "")
    decorated_attachments: list[dict[str, Any]] = []
    if isinstance(attachments, list):
        for index, attachment in enumerate(attachments):
            if not isinstance(attachment, dict):
                continue
            item = dict(attachment)
            href = str(item.get("href") or "").strip()
            local_path = str(item.get("localPath") or "").strip()
            if local_path:
                item["offlineReady"] = True
                if message_id:
                    item["viewerHref"] = build_viewer_attachment_href(conversation_id, message_id, index)
            if href:
                item["downloadHref"] = normalize_attachment_url(href)
                if message_id and not item.get("viewerHref"):
                    item["viewerHref"] = build_viewer_attachment_href(conversation_id, message_id, index)
            decorated_attachments.append(item)
    decorated["attachments"] = decorated_attachments
    return decorated


def _normalize_export_local_path(value: Any) -> str | None:
    text = _optional_str(value)
    if text is None:
        return None
    normalized = text.replace("\\", "/")
    candidate = PurePosixPath(normalized)
    if candidate.is_absolute() or ".." in candidate.parts:
        return None
    path_text = candidate.as_posix().lstrip("/")
    return path_text or None


def _quoted_export_href(path: str) -> str:
    # Keep bundle paths portable under file:// URLs. Percent signs in actual
    # filenames (for example "19%3Aalpha%40thread.v2") must be re-encoded to
    # "%25", otherwise browsers decode them to ":" / "@" and point at a
    # different on-disk path than the mirrored asset actually uses.
    return quote(path, safe="/:@-._~+()[]")


def _html_bundle_segment(value: str, *, fallback: str) -> str:
    text = unicodedata.normalize("NFKD", value)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Za-z0-9._-]+", "-", text).strip("-._").lower()
    return text or fallback


def _html_bundle_filename(name: str, *, key: str) -> str:
    candidate = PurePosixPath(name).name or "attachment"
    if "." in candidate:
        stem, suffix = candidate.rsplit(".", 1)
        ext = "." + re.sub(r"[^A-Za-z0-9]+", "", suffix).lower()
    else:
        stem, ext = candidate, ""
    safe_stem = _html_bundle_segment(stem, fallback="attachment")
    # Non-security digest used only to disambiguate bundled attachment filenames.
    digest = hashlib.sha1(key.encode("utf-8"), usedforsecurity=False).hexdigest()[:12]
    return f"{safe_stem}-{digest}{ext}"


def _html_bundle_dir_token(value: str, *, prefix: str, fallback: str) -> str:
    candidate = _html_bundle_segment(value, fallback=fallback)
    shortened = candidate[:24].strip("-._") or fallback
    # Non-security digest used only to disambiguate bundled directory names.
    digest = hashlib.sha1(value.encode("utf-8"), usedforsecurity=False).hexdigest()[:10]
    return f"{prefix}-{shortened}-{digest}"


def _html_bundle_asset_path(relative_path: str) -> str:
    # Use browser-friendly archive paths for HTML bundles so file:// rendering
    # does not depend on OS-specific percent-decoding behavior.
    parts = PurePosixPath(relative_path).parts
    conversation_part = parts[1] if len(parts) > 1 else ""
    message_part = parts[2] if len(parts) > 2 else ""
    filename_part = parts[-1] if parts else "attachment"
    conversation_dir = _html_bundle_dir_token(conversation_part, prefix="c", fallback="conversation")
    message_dir = _html_bundle_dir_token(message_part, prefix="m", fallback="message")
    filename = _html_bundle_filename(filename_part, key=relative_path)
    return PurePosixPath("assets", conversation_dir, message_dir, filename).as_posix()


def _attachment_label(attachment: dict[str, Any]) -> str:
    return str(attachment.get("label") or attachment.get("href") or "attachment")


def _attachment_is_image(attachment: dict[str, Any]) -> bool:
    attachment_type = str(attachment.get("type") or "").strip().casefold()
    if attachment_type.startswith("image"):
        return True
    for key in ("exportLocalPath", "localPath", "href", "label"):
        candidate = _optional_str(attachment.get(key))
        if candidate is None:
            continue
        guessed, _ = mimetypes.guess_type(candidate)
        if guessed and guessed.startswith("image/"):
            return True
    return False


def _prepare_export_payload(
    repo: ExportRepository,
    payload: dict[str, Any],
    *,
    conversation_id: str,
) -> tuple[dict[str, Any], dict[str, Path]]:
    prepared = dict(payload)
    bundled_assets: dict[str, Path] = {}
    messages = payload.get("messages", [])
    prepared_messages: list[dict[str, Any]] = []
    if not isinstance(messages, list):
        prepared["messages"] = prepared_messages
        return prepared, bundled_assets
    for item in messages:
        if not isinstance(item, dict):
            continue
        decorated = _decorate_message(item, conversation_id=conversation_id)
        prepared_attachments: list[dict[str, Any]] = []
        for attachment in decorated.get("attachments", []) or []:
            if not isinstance(attachment, dict):
                continue
            export_item = dict(attachment)
            normalized_remote_href = normalize_attachment_url(str(export_item.get("href") or ""))
            if normalized_remote_href:
                export_item["exportRemoteHref"] = normalized_remote_href
            normalized_local_path = _normalize_export_local_path(export_item.get("localPath"))
            resolved_local_path = repo.resolve_local_attachment(attachment)
            if normalized_local_path and resolved_local_path is not None:
                export_item["exportLocalPath"] = normalized_local_path
                export_item["exportHref"] = _quoted_export_href(normalized_local_path)
                bundled_assets.setdefault(normalized_local_path, resolved_local_path)
            elif normalized_remote_href:
                export_item["exportHref"] = normalized_remote_href
            export_item["isImage"] = _attachment_is_image(export_item)
            prepared_attachments.append(export_item)
        decorated["attachments"] = prepared_attachments
        prepared_messages.append(decorated)
    prepared["messages"] = prepared_messages
    return prepared, bundled_assets


def _prepare_html_bundle_payload(
    payload: dict[str, Any],
    bundled_assets: dict[str, Path],
) -> tuple[dict[str, Any], dict[str, Path]]:
    prepared = dict(payload)
    asset_path_map: dict[str, str] = {}
    html_bundle_assets: dict[str, Path] = {}
    prepared_messages: list[dict[str, Any]] = []
    for item in payload.get("messages", []) or []:
        if not isinstance(item, dict):
            continue
        message = dict(item)
        prepared_attachments: list[dict[str, Any]] = []
        for attachment in item.get("attachments", []) or []:
            if not isinstance(attachment, dict):
                continue
            export_item = dict(attachment)
            local_path = _normalize_export_local_path(export_item.get("exportLocalPath"))
            if local_path and local_path in bundled_assets:
                html_bundle_path = asset_path_map.setdefault(local_path, _html_bundle_asset_path(local_path))
                export_item["exportHref"] = quote(html_bundle_path, safe="/-._~+()[]")
                html_bundle_assets.setdefault(html_bundle_path, bundled_assets[local_path])
            prepared_attachments.append(export_item)
        message["attachments"] = prepared_attachments
        prepared_messages.append(message)
    prepared["messages"] = prepared_messages
    return prepared, html_bundle_assets


def _chat_to_csv(payload: dict[str, Any]) -> str:
    stream = io.StringIO()
    writer = csv.writer(stream)
    writer.writerow(
        [
            "timestamp",
            "author",
            "text",
            "system",
            "edited",
            "reactions",
            "attachments",
            "attachment_local_paths",
            "attachment_remote_urls",
            "reply_author",
            "reply_text",
        ]
    )
    for item in payload.get("messages", []):
        if not isinstance(item, dict):
            continue
        reactions = ", ".join(
            f"{reaction.get('emoji')}x{reaction.get('count')}"
            for reaction in item.get("reactions", []) or []
            if isinstance(reaction, dict)
        )
        attachment_items = [
            attachment for attachment in item.get("attachments", []) or [] if isinstance(attachment, dict)
        ]
        attachments = ", ".join(_attachment_label(attachment) for attachment in attachment_items)
        attachment_local_paths = ", ".join(
            str(attachment.get("exportLocalPath") or "")
            for attachment in attachment_items
            if str(attachment.get("exportLocalPath") or "").strip()
        )
        attachment_remote_urls = ", ".join(
            str(attachment.get("exportRemoteHref") or "")
            for attachment in attachment_items
            if str(attachment.get("exportRemoteHref") or "").strip()
        )
        reply = item.get("replyTo") if isinstance(item.get("replyTo"), dict) else {}
        writer.writerow(
            [
                item.get("timestamp", ""),
                item.get("author", ""),
                item.get("text", ""),
                item.get("system", False),
                item.get("edited", False),
                reactions,
                attachments,
                attachment_local_paths,
                attachment_remote_urls,
                reply.get("author", ""),
                reply.get("text", ""),
            ]
        )
    return stream.getvalue()


def _chat_to_markdown(payload: dict[str, Any]) -> str:
    meta = payload.get("meta", {}) if isinstance(payload.get("meta"), dict) else {}
    lines = [f"# {meta.get('title', 'Teams Export')}", ""]
    if meta.get("timeRange"):
        lines.append(f"_Range: {meta['timeRange']}_")
        lines.append("")
    for item in payload.get("messages", []):
        if not isinstance(item, dict):
            continue
        author = item.get("author", "[unknown]")
        timestamp = item.get("timestamp", "")
        lines.append(f"## {timestamp} - {author}")
        if item.get("system"):
            lines.append("_system message_")
        lines.append("")
        lines.append(str(item.get("text", "") or "[no text]"))
        reply = item.get("replyTo") if isinstance(item.get("replyTo"), dict) else {}
        if reply.get("text"):
            lines.append("")
            lines.append(f"> Reply to {reply.get('author', '[unknown]')}: {reply.get('text', '')}")
        attachments = item.get("attachments", []) or []
        if attachments:
            lines.append("")
            lines.append("Attachments:")
            for attachment in attachments:
                if not isinstance(attachment, dict):
                    continue
                label = _attachment_label(attachment)
                href = str(attachment.get("exportHref") or attachment.get("exportRemoteHref") or "").strip()
                if attachment.get("isImage") and href:
                    alt = label.replace("[", "\\[").replace("]", "\\]")
                    lines.append(f"![{alt}]({href})")
                elif href:
                    lines.append(f"- [{label}]({href})")
                else:
                    lines.append(f"- {label}")
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def _chat_to_html(payload: dict[str, Any]) -> str:
    meta = payload.get("meta", {}) if isinstance(payload.get("meta"), dict) else {}
    title = html.escape(str(meta.get("title", "Teams Export")))
    items: list[str] = []
    for message in payload.get("messages", []):
        if not isinstance(message, dict):
            continue
        attachments = message.get("attachments", []) or []
        attachment_html = ""
        if attachments:
            image_blocks: list[str] = []
            link_rows: list[str] = []
            for attachment in attachments:
                if not isinstance(attachment, dict):
                    continue
                label = html.escape(_attachment_label(attachment))
                href = str(attachment.get("exportHref") or attachment.get("exportRemoteHref") or "").strip()
                escaped_href = html.escape(href)
                if attachment.get("isImage") and href:
                    image_blocks.append(
                        "".join(
                            [
                                "<figure class='attachment attachment-image'>",
                                f"<a href=\"{escaped_href}\"><img src=\"{escaped_href}\" alt=\"{label}\"></a>",
                                f"<figcaption><a href=\"{escaped_href}\">{label}</a></figcaption>",
                                "</figure>",
                            ]
                        )
                    )
                    continue
                if href:
                    link_rows.append(f"<li><a href=\"{escaped_href}\">{label}</a></li>")
                else:
                    link_rows.append(f"<li>{label}</li>")
            attachment_parts = list(image_blocks)
            if link_rows:
                attachment_parts.append(f"<ul>{''.join(link_rows)}</ul>")
            attachment_html = f"<div class='attachments'>{''.join(attachment_parts)}</div>"
        items.append(
            "".join(
                [
                    "<article class='msg'>",
                    f"<header><strong>{html.escape(str(message.get('author') or '[unknown]'))}</strong>",
                    f"<span>{html.escape(str(message.get('timestamp') or ''))}</span></header>",
                    f"<p>{html.escape(str(message.get('text') or '[no text]'))}</p>",
                    attachment_html,
                    "</article>",
                ]
            )
        )
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title}</title>
  <style>
    body {{ font-family: ui-sans-serif, system-ui, sans-serif; background: #09111d; color: #e6f6ff; margin: 0; padding: 24px; }}
    h1 {{ margin-top: 0; color: #6cf0ff; }}
    article.msg {{ border: 1px solid #183247; border-radius: 14px; padding: 14px; margin: 0 0 12px; background: rgba(5, 15, 28, 0.9); }}
    article.msg header {{ display: flex; justify-content: space-between; gap: 12px; color: #9fe6ff; }}
    .attachments {{ margin-top: 12px; }}
    .attachments ul {{ margin: 0; padding-left: 18px; }}
    .attachment-image {{ margin: 12px 0; }}
    .attachment-image img {{ display: block; max-width: min(100%, 960px); border-radius: 10px; border: 1px solid #183247; }}
    .attachment-image figcaption {{ margin-top: 8px; }}
    a {{ color: #ffb347; }}
  </style>
</head>
<body>
  <h1>{title}</h1>
  <p>{html.escape(str(meta.get("timeRange") or ""))}</p>
  {''.join(items)}
</body>
</html>"""


def _chat_to_html_bundle(
    payload: dict[str, Any],
    *,
    archive_root: str,
    bundled_assets: dict[str, Path],
) -> bytes:
    html_content = _chat_to_html(payload)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"{archive_root}/index.html", html_content)
        for relative_path, source_path in sorted(bundled_assets.items()):
            archive.write(source_path, arcname=f"{archive_root}/{relative_path}")
    return buffer.getvalue()


def _serialize_export_progress(snapshot: ExportProgress | None) -> dict[str, Any] | None:
    if snapshot is None:
        return None
    return {
        "phase": snapshot.phase,
        "totalConversations": snapshot.total_conversations,
        "totalHidden": snapshot.total_hidden,
        "totalMeeting": snapshot.total_meeting,
        "processedConversations": snapshot.processed_conversations,
        "processedHidden": snapshot.processed_hidden,
        "processedMeeting": snapshot.processed_meeting,
        "exportedConversations": snapshot.exported_conversations,
        "skippedConversations": snapshot.skipped_conversations,
        "failedConversations": snapshot.failed_conversations,
        "exportedMessages": snapshot.exported_messages,
        "currentConversationId": snapshot.current_conversation_id,
        "currentTitle": snapshot.current_title,
        "currentHidden": snapshot.current_hidden,
        "currentMeeting": snapshot.current_meeting,
        "currentMessageCount": snapshot.current_message_count,
        "currentStatus": snapshot.current_status,
        "note": snapshot.note,
    }


def _serialize_export_result(result: ExportAllResult | None) -> dict[str, Any] | None:
    if result is None:
        return None
    return {
        "ok": result.ok,
        "message": result.message,
        "conversationCount": result.conversation_count,
        "exportedCount": result.exported_count,
        "failedCount": result.failed_count,
        "hiddenCount": result.hidden_count,
        "meetingCount": result.meeting_count,
        "messageCount": result.message_count,
        "interrupted": result.interrupted,
        "outdir": str(result.outdir) if result.outdir else None,
        "indexPath": str(result.index_path) if result.index_path else None,
    }


def _serialize_attachment_mirror_progress(snapshot: AttachmentMirrorProgress | None) -> dict[str, Any] | None:
    if snapshot is None:
        return None
    return {
        "phase": snapshot.phase,
        "totalChats": snapshot.total_chats,
        "totalAssets": snapshot.total_assets,
        "processedChats": snapshot.processed_chats,
        "processedAssets": snapshot.processed_assets,
        "mirroredAssets": snapshot.mirrored_assets,
        "reusedAssets": snapshot.reused_assets,
        "failedAssets": snapshot.failed_assets,
        "currentChatTitle": snapshot.current_chat_title,
        "currentAssetLabel": snapshot.current_asset_label,
        "currentStatus": snapshot.current_status,
        "elapsedSeconds": snapshot.elapsed_seconds,
        "etaSeconds": snapshot.eta_seconds,
        "freeBytes": snapshot.free_bytes,
        "minFreeBytes": snapshot.min_free_bytes,
        "bytesDownloaded": snapshot.bytes_downloaded,
        "note": snapshot.note,
    }


def _serialize_attachment_mirror_result(result: MirrorAttachmentsResult | None) -> dict[str, Any] | None:
    if result is None:
        return None
    return {
        "ok": result.ok,
        "message": result.message,
        "chatCount": result.chat_count,
        "attachmentCount": result.attachment_count,
        "mirroredCount": result.mirrored_count,
        "reusedCount": result.reused_count,
        "failedCount": result.failed_count,
        "interrupted": result.interrupted,
        "lowDisk": result.low_disk,
        "freeBytes": result.free_bytes,
        "minFreeBytes": result.min_free_bytes,
        "bundleRoot": str(result.bundle_root) if result.bundle_root else None,
        "assetsDir": str(result.assets_dir) if result.assets_dir else None,
    }


def _safe_stem(value: str) -> str:
    cleaned = "".join(char if char.isalnum() or char in {"-", "_", "."} else "-" for char in value).strip("-")
    return cleaned or "teams-export"


def _first_param(params: dict[str, list[str]], key: str, default: str = "") -> str:
    values = params.get(key)
    if not values:
        return default
    return values[0]


def _origin_matches_host(host: str, header_value: str) -> bool:
    normalized_host = str(host or "").strip().lower()
    if not normalized_host:
        return False
    parsed = urlparse(header_value)
    if parsed.scheme not in {"http", "https"}:
        return False
    netloc = str(parsed.netloc or "").strip().lower()
    if not netloc:
        return False
    return netloc == normalized_host


def _int_param(params: dict[str, list[str]], key: str, default: int) -> int:
    try:
        return int(_first_param(params, key, str(default)))
    except ValueError:
        return default


def _app_html(control_token: str, build_label: str = "dev") -> str:
    html = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MS Teams Exporter/Viewer v__BUILD_LABEL__</title>
  <style>
    :root {
      --bg: #050913;
      --panel: rgba(10, 21, 35, 0.88);
      --panel-2: rgba(14, 28, 46, 0.92);
      --line: #173452;
      --accent: #66f7ff;
      --accent-2: #ffb347;
      --text: #ebfbff;
      --muted: #90bfd0;
      --warn: #ff8f70;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: "SF Mono", "JetBrains Mono", "IBM Plex Mono", monospace;
      color: var(--text);
      background:
        radial-gradient(circle at top left, rgba(102, 247, 255, 0.16), transparent 28%),
        radial-gradient(circle at top right, rgba(255, 179, 71, 0.14), transparent 22%),
        linear-gradient(160deg, #03060d 0%, #07101c 48%, #04070c 100%);
    }
    .shell {
      display: flex;
      min-height: 100vh;
      gap: 20px;
      align-items: stretch;
    }
    .sidebar, .main {
      padding: 20px;
    }
    .sidebar {
      width: 360px;
      min-width: 280px;
      max-width: 760px;
      flex: 0 0 auto;
      resize: horizontal;
      overflow: auto;
      position: relative;
      isolation: isolate;
      border-right: 1px solid var(--line);
      background: linear-gradient(180deg, rgba(6, 14, 24, 0.92), rgba(6, 14, 24, 0.78));
    }
    .sidebar::after {
      content: "";
      position: absolute;
      inset: 0;
      background:
        linear-gradient(180deg, rgba(5, 9, 19, 0.04), rgba(5, 9, 19, 0.16)),
        linear-gradient(90deg, rgba(6, 14, 24, 0.08), rgba(6, 14, 24, 0.01));
      pointer-events: none;
      z-index: 0;
    }
    .sidebar > :not(.matrix-rain) {
      position: relative;
      z-index: 2;
    }
    .main {
      flex: 1 1 auto;
      min-width: 0;
      display: grid;
      grid-template-rows: auto auto auto 1fr;
      gap: 16px;
    }
    h1, h2, h3 { margin: 0; font-weight: 600; letter-spacing: 0.04em; }
    .title { color: var(--accent); text-transform: uppercase; font-size: 14px; }
    .title-wrap {
      display: flex;
      align-items: center;
      gap: 10px;
      min-width: 0;
    }
    .build-chip {
      border: 1px solid rgba(102, 247, 255, 0.2);
      border-radius: 999px;
      padding: 4px 8px;
      color: var(--muted);
      font-size: 10px;
      letter-spacing: 0.08em;
      white-space: nowrap;
    }
    .subtitle { color: var(--muted); font-size: 12px; margin-top: 6px; }
    .sidebar-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }
    .help-toggle {
      padding: 7px 12px;
      font-size: 12px;
      letter-spacing: 0.08em;
    }
    .help-panel {
      margin-top: 12px;
      padding: 12px 14px;
      border: 1px solid rgba(102, 247, 255, 0.18);
      background: rgba(8, 18, 29, 0.72);
      border-radius: 14px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.5;
      display: none;
      white-space: pre-line;
    }
    .help-panel.open {
      display: block;
    }
    .summary-block {
      color: var(--muted);
      font-size: 12px;
      margin-top: 10px;
      line-height: 1.45;
      white-space: pre-line;
      overflow-wrap: anywhere;
    }
    .matrix-rain {
      position: absolute;
      inset: 0;
      overflow: hidden;
      pointer-events: none;
      z-index: 1;
      opacity: 0.52;
    }
    .matrix-rain canvas {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      display: block;
      filter: blur(0.15px);
    }
    .panel {
      border: 1px solid var(--line);
      background: var(--panel);
      border-radius: 18px;
      box-shadow: 0 24px 80px rgba(0, 0, 0, 0.32);
      min-width: 0;
    }
    .sidebar .panel {
      background: rgba(10, 21, 35, 0.32);
      backdrop-filter: blur(1.5px);
    }
    .control-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(280px, 1fr));
      gap: 16px;
    }
    .control-shell {
      display: grid;
      gap: 10px;
    }
    .control-shell-head {
      display: flex;
      justify-content: flex-end;
      align-items: center;
      min-height: 22px;
    }
    .control-panels-toggle {
      width: 34px;
      height: 34px;
      padding: 0;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      font-size: 18px;
      line-height: 1;
      border-radius: 999px;
      color: var(--accent);
    }
    .control-grid.collapsed .control-card-body {
      display: none;
    }
    .control-grid.collapsed .control-card {
      padding-top: 14px;
      padding-bottom: 14px;
    }
    .control-card {
      padding: 16px;
      display: grid;
      gap: 12px;
    }
    .control-card-head {
      display: grid;
      gap: 4px;
    }
    .toolbar, .meta, .chat {
      padding: 16px;
    }
    .toolbar { display: grid; gap: 12px; }
    .filters, .chat-filters, .export-actions, .control-actions, .control-inputs {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
    }
    .sidebar-search-row {
      display: grid;
      gap: 10px;
    }
    .search-with-toggle {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
    }
    .search-with-toggle input {
      flex: 1 1 220px;
      min-width: 0;
    }
    .filters-toggle-row {
      display: flex;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
    }
    .filters-toggle-button {
      padding: 7px 12px;
      font-size: 12px;
      letter-spacing: 0.05em;
    }
    .sidebar-advanced-filters {
      display: none;
      gap: 10px;
      flex-wrap: wrap;
      align-items: center;
    }
    .sidebar-advanced-filters.open {
      display: flex;
    }
    .field-stack {
      display: grid;
      gap: 6px;
      min-width: 240px;
    }
    .statline {
      font-size: 12px;
      color: var(--muted);
      line-height: 1.45;
    }
    .bar-stack {
      display: grid;
      gap: 8px;
    }
    .bar-row {
      display: grid;
      gap: 4px;
    }
    .bar-label {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      font-size: 12px;
      color: var(--muted);
    }
    .bar {
      height: 10px;
      border-radius: 999px;
      overflow: hidden;
      border: 1px solid var(--line);
      background: rgba(8, 18, 29, 0.8);
    }
    .bar-fill {
      height: 100%;
      width: 0%;
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
      transition: width 160ms ease;
    }
    .tiny { font-size: 11px; color: var(--muted); }
    input, select, button {
      border: 1px solid var(--line);
      background: var(--panel-2);
      color: var(--text);
      border-radius: 12px;
      padding: 10px 12px;
      font: inherit;
    }
    button {
      cursor: pointer;
      transition: transform 120ms ease, border-color 120ms ease, color 120ms ease;
    }
    button:hover {
      transform: translateY(-1px);
      border-color: var(--accent);
      color: var(--accent);
    }
    button:disabled {
      opacity: 0.45;
      cursor: not-allowed;
      transform: none;
      color: var(--muted);
    }
    button.danger {
      border-color: rgba(255, 143, 112, 0.45);
      color: var(--warn);
    }
    .conversation-list {
      margin-top: 16px;
      display: grid;
      gap: 10px;
      max-height: calc(100vh - 170px);
      overflow: auto;
      padding-right: 4px;
    }
    .empty-state {
      border: 1px dashed rgba(102, 247, 255, 0.22);
      border-radius: 14px;
      padding: 14px;
      background: rgba(8, 18, 29, 0.26);
      color: var(--muted);
      line-height: 1.5;
    }
    .empty-state strong {
      display: block;
      margin-bottom: 6px;
      color: var(--text);
    }
    .conversation-group {
      display: grid;
      gap: 8px;
    }
    .conversation-group-title {
      font-size: 11px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      color: var(--accent);
      padding: 2px 4px 0;
    }
    .conversation-group-items {
      display: grid;
      gap: 10px;
    }
    .conversation {
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 12px;
      background: rgba(8, 18, 29, 0.38);
      cursor: pointer;
      width: 100%;
      text-align: left;
    }
    .conversation.active { border-color: var(--accent); box-shadow: 0 0 0 1px rgba(102, 247, 255, 0.18); }
    .conversation .name { color: var(--text); display: block; text-align: left; }
    .conversation .meta-line { color: var(--muted); font-size: 12px; margin-top: 6px; }
    .conversation .relation-line {
      display: flex;
      align-items: center;
      gap: 8px;
      min-width: 0;
      color: var(--accent-2);
    }
    .conversation .relation-name {
      min-width: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .flags { display: flex; gap: 8px; margin-top: 8px; flex-wrap: wrap; }
    .flag {
      font-size: 11px;
      padding: 3px 8px;
      border-radius: 999px;
      border: 1px solid var(--line);
      color: var(--accent-2);
    }
    .flag.relation-flag {
      color: var(--accent);
      border-color: rgba(102, 247, 255, 0.24);
      background: rgba(10, 28, 40, 0.48);
      flex: 0 0 auto;
    }
    .sidebar-toggle {
      font-size: 11px;
      color: var(--muted);
      display: inline-flex;
      align-items: center;
      gap: 8px;
    }
    .chat {
      overflow-x: hidden;
      overflow-y: auto;
      max-height: calc(100vh - 300px);
      display: flex;
      flex-direction: column;
      gap: 12px;
      min-width: 0;
    }
    .msg {
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px;
      background: rgba(8, 18, 29, 0.86);
      min-width: 0;
      overflow: hidden;
      flex: 0 0 auto;
    }
    .msg.system { border-color: rgba(255, 179, 71, 0.3); }
    .msg header {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      align-items: start;
      gap: 12px;
      margin-bottom: 8px;
      color: var(--accent);
      font-size: 12px;
      min-width: 0;
    }
    .msg header span {
      min-width: 0;
      overflow-wrap: anywhere;
    }
    .msg header span:last-child {
      margin-left: 0;
      text-align: right;
      justify-self: end;
    }
    .msg .body {
      white-space: pre-wrap;
      line-height: 1.45;
      overflow-wrap: anywhere;
      word-break: break-word;
      min-width: 0;
      color: var(--text);
    }
    .msg .extra {
      margin-top: 8px;
      color: var(--muted);
      font-size: 12px;
      overflow-wrap: anywhere;
      word-break: break-word;
      min-width: 0;
      display: grid;
      gap: 6px;
    }
    .msg .extra a { color: var(--accent-2); }
    .muted { color: var(--muted); }
    .warn { color: var(--warn); }
    @media (max-width: 980px) {
      .shell { display: block; }
      .sidebar { border-right: 0; border-bottom: 1px solid var(--line); width: auto; min-width: 0; max-width: none; resize: none; }
      .conversation-list { max-height: 40vh; }
      .chat { max-height: none; }
      .control-grid { grid-template-columns: 1fr; }
      .msg header { grid-template-columns: 1fr; }
      .msg header span:last-child { justify-self: start; text-align: left; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <aside class="sidebar">
      <div class="matrix-rain" id="matrixRain" aria-hidden="true"></div>
      <div class="sidebar-head">
        <div class="title-wrap">
          <div class="title">MS Teams Exporter/Viewer</div>
          <span class="build-chip">v__BUILD_LABEL__</span>
        </div>
        <button id="helpToggle" class="help-toggle" type="button" aria-expanded="false" aria-controls="helpPanel">HELP</button>
      </div>
      <div class="help-panel" id="helpPanel">
This is cakeware application. If you like it, send to Martin Ryšavý a small cake :)

This app is both an exporter and a viewer.

Use Export Control to build or resume the JSON archive.
Use Attachment Mirror to download useful files into `assets/` for offline viewing.

Control-panel POST actions are local-only by design:
- the server binds to `127.0.0.1` by default
- state-changing POST calls need a random session `X-Control-Token`
- control POSTs also check same-origin `Origin` / `Referer` when the browser sends them
- plain external form posts or blind cross-site POSTs should not pass these checks

This is still a localhost tool, not something to expose to a LAN or the internet.

After export, the left panel lists stored conversations. Click any conversation to inspect it.

Search notes for advanced users:
- plain text search is supported by default
- `/pattern/flags` regex syntax is also supported in search fields
- the `case sensitive` toggle affects normal text search and regex matching unless regex flags override it
- `Search all chat messages` scans exported message content across chats and narrows the sidebar to matching conversations
- Global content search waits for at least 3 characters, uses a short debounce, and returns a capped result set to stay polite on large exports
- Enable `full search` when you intentionally want the uncapped cross-chat result set
- regex examples for server names:
  - `/\\bAABB\\d{2}\\b/` -> matches `AABB01`, `AABB99`
  - `/\\bAABB[- ]?\\d{2}\\b/` -> matches `AABB01`, `AABB-01`, `AABB 01`
  - `/\\b(?:APP|DB|WEB)\\d{2}\\b/i` -> matches multiple server prefixes with two-digit indexes
  - `/\\bSRV-[A-Z]{3}-\\d{2}\\b/` -> matches names like `SRV-PRD-01`

Common flags:
- `chat`: a regular one-to-one or group chat
- `meeting-chat`: a meeting thread
- `channel`, `team`, `community`: non-chat conversation spaces
- `hidden`: a conversation that was hidden in Teams
- `skipped`: the JSON file already existed and was reused instead of exported again
- `partial`: only part of the history or metadata could be recovered
  For regular chats, a partial export can often be retried by running export again with `skip existing` turned off.
- `metadata-only`: conversation shell found, but not the full message history
      </div>
      <div class="summary-block" id="summary">Loading export index...</div>
      <div class="panel toolbar">
        <div class="sidebar-search-row">
          <input id="conversationQuery" placeholder="Search conversations">
          <div class="search-with-toggle">
            <input id="globalSearchQuery" placeholder="Search all chat messages">
            <label class="sidebar-toggle"><input id="globalSearchFull" type="checkbox"> full search</label>
          </div>
          <div class="filters-toggle-row">
            <label class="sidebar-toggle"><input id="groupByTeam" type="checkbox"> group by team</label>
            <label class="sidebar-toggle"><input id="caseSensitive" type="checkbox"> case sensitive</label>
            <button id="sidebarFiltersToggle" class="filters-toggle-button" type="button" aria-expanded="false" aria-controls="sidebarAdvancedFilters">filters ▸</button>
          </div>
        </div>
        <div class="sidebar-advanced-filters" id="sidebarAdvancedFilters">
          <select id="hiddenFilter">
            <option value="any">all visibility</option>
            <option value="true">hidden only</option>
            <option value="false">visible only</option>
          </select>
          <select id="meetingFilter">
            <option value="any">all types</option>
            <option value="true">meeting only</option>
            <option value="false">non-meeting</option>
          </select>
          <select id="statusFilter">
            <option value="any">all status</option>
            <option value="failed">failed only</option>
            <option value="metadata-only">metadata only</option>
            <option value="partial">partial only</option>
            <option value="exported">exported only</option>
            <option value="skipped">reused only</option>
          </select>
          <select id="kindFilter">
            <option value="any">all kinds</option>
            <option value="chat">chat</option>
            <option value="meeting-chat">meeting chat</option>
            <option value="channel">channel</option>
            <option value="team">team</option>
            <option value="community">community</option>
          </select>
        </div>
      </div>
      <div class="conversation-list" id="conversationList"></div>
    </aside>
    <main class="main">
      <section class="control-shell">
        <div class="control-shell-head">
          <button
            id="controlPanelsToggle"
            class="control-panels-toggle"
            type="button"
            aria-expanded="true"
            aria-controls="controlGrid"
            title="collapse or expand export panels"
          >▾</button>
        </div>
        <section class="control-grid" id="controlGrid">
        <section class="panel control-card">
          <div class="control-card-head">
            <div class="title">Export Control</div>
            <div class="subtitle">Run or resume the full chat export for this bundle.</div>
          </div>
          <div class="control-card-body">
          <div class="control-inputs">
            <label><input id="exportSkipExisting" type="checkbox" checked> skip existing</label>
            <input id="exportMaxChats" type="number" min="0" step="1" placeholder="max chats (optional)">
          </div>
          <div class="control-actions">
            <button id="exportStartButton" type="button">start or resume export</button>
            <button id="exportStopButton" type="button">stop gracefully</button>
            <button id="exportForceStopButton" class="danger" type="button">stop after current page</button>
          </div>
          <div class="subtitle" id="exportStatusText">idle</div>
          <div class="statline" id="exportStatusMeta">No export job running.</div>
          <div class="bar-stack">
            <div class="bar-row">
              <div class="bar-label"><span>all chats</span><span id="exportAllCount">0 / 0</span></div>
              <div class="bar"><div class="bar-fill" id="exportAllBar"></div></div>
            </div>
            <div class="bar-row">
              <div class="bar-label"><span>hidden</span><span id="exportHiddenCount">0 / 0</span></div>
              <div class="bar"><div class="bar-fill" id="exportHiddenBar"></div></div>
            </div>
            <div class="bar-row">
              <div class="bar-label"><span>meeting</span><span id="exportMeetingCount">0 / 0</span></div>
              <div class="bar"><div class="bar-fill" id="exportMeetingBar"></div></div>
            </div>
          </div>
          </div>
        </section>
        <section class="panel control-card">
          <div class="control-card-head">
            <div class="title">Attachment Mirror</div>
            <div class="subtitle">Mirror useful attachments into `assets/` for offline use.</div>
          </div>
          <div class="control-card-body">
          <div class="control-inputs">
            <div class="field-stack">
              <label class="tiny" for="mirrorMinFreeGb">Minimum free disk space to keep (GB)</label>
              <input id="mirrorMinFreeGb" type="number" min="0" step="1" value="30" placeholder="minimum free GB">
            </div>
          </div>
          <div class="tiny">Animated GIF and video attachments stay excluded from mirroring.</div>
          <div class="control-actions">
            <button id="mirrorStartButton" type="button">start or resume mirror</button>
            <button id="mirrorStopButton" type="button">stop gracefully</button>
            <button id="mirrorForceStopButton" class="danger" type="button">stop after current file</button>
          </div>
          <div class="subtitle" id="mirrorStatusText">idle</div>
          <div class="statline" id="mirrorStatusMeta">No mirror job running.</div>
          <div class="bar-stack">
            <div class="bar-row">
              <div class="bar-label"><span>assets</span><span id="mirrorAssetCount">0 / 0</span></div>
              <div class="bar"><div class="bar-fill" id="mirrorAssetBar"></div></div>
            </div>
            <div class="bar-row">
              <div class="bar-label"><span>chats</span><span id="mirrorChatCount">0 / 0</span></div>
              <div class="bar"><div class="bar-fill" id="mirrorChatBar"></div></div>
            </div>
          </div>
          </div>
        </section>
        </section>
      </section>
      <section class="panel toolbar">
        <div class="chat-filters">
          <input id="messageQuery" placeholder="Search inside selected chat">
          <input id="authorFilter" placeholder="Author filter">
          <label><input id="hideSystem" type="checkbox"> hide system</label>
          <button id="reloadButton" type="button">reload</button>
        </div>
        <div class="export-actions">
          <button data-export-format="csv" type="button">export csv</button>
          <button data-export-format="md" type="button">export markdown</button>
          <button data-export-format="html" type="button">export html</button>
        </div>
      </section>
      <section class="panel meta">
        <div class="title" id="chatTitle">No chat selected</div>
        <div class="subtitle" id="chatMeta">Choose a conversation on the left.</div>
      </section>
      <section class="panel chat" id="chat"></section>
    </main>
  </div>
  <script>
    const CONTROL_TOKEN = "__CONTROL_TOKEN__";
    const SEARCH_DEBOUNCE_MS = 320;
    const CHAT_FILTER_DEBOUNCE_MS = 220;
    const GLOBAL_SEARCH_MIN_CHARS = 3;
    const GLOBAL_SEARCH_LIMIT = __GLOBAL_SEARCH_LIMIT__;
    const state = {
      selectedConversationId: null,
      conversations: [],
      exportCompletionKey: '',
      mirrorCompletionKey: '',
      repositoryRefreshPromise: null,
      sidebarEmptyTitle: 'No exported conversations yet.',
      sidebarEmptyHint: 'The sidebar will refresh automatically after export completes.',
      globalSearchQuery: '',
      conversationSearchMode: 'titles',
      conversationSearchTimer: null,
      chatSearchTimer: null,
      conversationsAbortController: null,
      chatAbortController: null,
      conversationRequestToken: 0,
      chatRequestToken: 0,
      globalSearchFullMode: false,
    };

    const summaryEl = document.getElementById('summary');
    const conversationListEl = document.getElementById('conversationList');
    const chatTitleEl = document.getElementById('chatTitle');
    const chatMetaEl = document.getElementById('chatMeta');
    const chatEl = document.getElementById('chat');
    let statusPoller = null;
    let matrixAnimationFrame = null;
    let matrixCanvasState = null;
    let matrixResizeHandler = null;
    let matrixResizeObserver = null;

    async function loadMeta() {
      const response = await fetch('/api/meta');
      const payload = await response.json();
      const mirrored = payload.mirroredAttachmentCount ?? payload.meta?.mirroredAttachmentCount ?? 0;
      const attachmentCount = payload.attachmentCount ?? payload.meta?.attachmentCount ?? 0;
      const unauthorizedCount = payload.unauthorizedAttachmentCount ?? payload.meta?.unauthorizedAttachmentCount ?? 0;
      const totalMessages = payload.messageCount ?? payload.meta?.messageCount ?? 0;
      const bundleRoot = payload.bundleRoot || '';
      summaryEl.textContent = [
        `stored conversations: ${payload.conversationCount} | hidden: ${payload.hiddenCount} | meeting: ${payload.meetingCount}`,
        `total messages: ${totalMessages} | partial: ${payload.partialCount} | metadata-only: ${payload.metadataOnlyCount} | failed: ${payload.failedCount}`,
        `assets mirrored: ${mirrored} / ${attachmentCount} | unauthorized: ${unauthorizedCount}`,
        `export root: ${bundleRoot}`,
      ].join('\\n');
    }

    async function loadConversations() {
      const query = document.getElementById('conversationQuery').value.trim();
      const globalQuery = document.getElementById('globalSearchQuery').value.trim();
      const globalSearchFullToggle = document.getElementById('globalSearchFull');
      const caseSensitive = document.getElementById('caseSensitive').checked ? 'true' : 'false';
      const hidden = document.getElementById('hiddenFilter').value;
      const meeting = document.getElementById('meetingFilter').value;
      const status = document.getElementById('statusFilter').value;
      const kind = document.getElementById('kindFilter').value;
      state.globalSearchQuery = globalQuery;
      globalSearchFullToggle.disabled = !globalQuery || globalQuery.length < GLOBAL_SEARCH_MIN_CHARS;
      if (!state.globalSearchFullMode) {
        globalSearchFullToggle.checked = false;
      }
      const requestToken = ++state.conversationRequestToken;
      if (state.conversationsAbortController) {
        state.conversationsAbortController.abort();
      }
      state.conversationsAbortController = new AbortController();
      const params = new URLSearchParams({
        caseSensitive,
        hidden,
        meeting,
        status,
        kind,
        exported: 'true',
      });
      let endpoint = '/api/conversations';
      if (globalQuery) {
        if (globalQuery.length < GLOBAL_SEARCH_MIN_CHARS) {
          state.conversations = [];
          state.conversationSearchMode = 'content';
          state.globalSearchFullMode = false;
          globalSearchFullToggle.checked = false;
          state.sidebarEmptyTitle = `Global content search needs at least ${GLOBAL_SEARCH_MIN_CHARS} characters.`;
          state.sidebarEmptyHint = 'Keep typing before scanning all exported chat messages.';
          renderConversationList();
          state.selectedConversationId = null;
          renderNoConversationSelected(state.sidebarEmptyTitle, state.sidebarEmptyHint);
          return;
        }
        endpoint = '/api/search';
        params.set('query', globalQuery);
        params.set('conversationQuery', query);
        params.set('hideSystem', document.getElementById('hideSystem').checked ? 'true' : 'false');
        params.set('limit', state.globalSearchFullMode ? '-1' : String(GLOBAL_SEARCH_LIMIT));
      } else {
        state.globalSearchFullMode = false;
        globalSearchFullToggle.checked = false;
        params.set('query', query);
      }
      let payload;
      try {
        const response = await fetch(`${endpoint}?${params.toString()}`, {
          signal: state.conversationsAbortController.signal,
        });
        payload = await response.json();
      } catch (error) {
        if (error && error.name === 'AbortError') {
          return;
        }
        throw error;
      }
      if (requestToken !== state.conversationRequestToken) {
        return;
      }
      if (payload.error) {
        state.conversations = [];
        state.sidebarEmptyTitle = 'Conversation filter error.';
        state.sidebarEmptyHint = payload.error;
        renderConversationList();
        state.selectedConversationId = null;
        renderNoConversationSelected('Conversation filter error.', payload.error);
        return;
      }
      state.conversations = payload.conversations || [];
      state.conversationSearchMode = payload.searchMode || (globalQuery ? 'content' : 'titles');
      const filtersActive = Boolean(query) || hidden !== 'any' || meeting !== 'any' || status !== 'any' || kind !== 'any';
      if (globalQuery) {
        state.sidebarEmptyTitle = payload.truncated
          ? `Showing the top ${payload.count} global matches.`
          : 'No chats matched the global content search.';
        state.sidebarEmptyHint = payload.truncated && !state.globalSearchFullMode
          ? 'Refine the message search or sidebar filters to reduce load and narrow the result set.'
          : state.globalSearchFullMode
            ? 'Full cross-chat search mode is active.'
            : 'Adjust the global message search, case sensitivity, or the sidebar filters.';
      } else {
        state.sidebarEmptyTitle = filtersActive ? 'No conversations match the current filters.' : 'No exported conversations yet.';
        state.sidebarEmptyHint = filtersActive
          ? 'Clear or loosen the filters on the left to see more results.'
          : 'The sidebar will refresh automatically after export completes.';
      }
      renderConversationList();
      if (!state.conversations.length) {
        state.selectedConversationId = null;
        renderNoConversationSelected(
          filtersActive ? 'No conversations match the current sidebar filters.' : 'No exported conversations yet.',
          filtersActive
            ? 'Clear or loosen the filters on the left to see more results.'
            : 'Run Export Control. The sidebar refreshes automatically when the export finishes.',
        );
        return;
      }
      if ((!state.selectedConversationId || !state.conversations.find((item) => item.id === state.selectedConversationId)) && state.conversations.length) {
        state.selectedConversationId = state.conversations[0].id;
      }
      if (state.selectedConversationId) {
        await loadChat();
      }
    }

    function scheduleLoadConversations() {
      if (state.conversationSearchTimer) {
        window.clearTimeout(state.conversationSearchTimer);
      }
      state.conversationSearchTimer = window.setTimeout(() => {
        state.conversationSearchTimer = null;
        void loadConversations();
      }, SEARCH_DEBOUNCE_MS);
    }

    function resetGlobalSearchModeAndSchedule() {
      state.globalSearchFullMode = false;
      const toggle = document.getElementById('globalSearchFull');
      const globalQuery = document.getElementById('globalSearchQuery').value.trim();
      toggle.checked = false;
      toggle.disabled = !globalQuery || globalQuery.length < GLOBAL_SEARCH_MIN_CHARS;
      scheduleLoadConversations();
    }

    function renderConversationList() {
      conversationListEl.innerHTML = '';
      if (!state.conversations.length) {
        const empty = document.createElement('div');
        empty.className = 'empty-state';
        empty.innerHTML = `
          <strong>${escapeHtml(state.sidebarEmptyTitle)}</strong>
          <div>${escapeHtml(state.sidebarEmptyHint)}</div>
        `;
        conversationListEl.appendChild(empty);
        queueMatrixResize();
        return;
      }
      const grouped = document.getElementById('groupByTeam').checked;
      if (!grouped) {
        for (const conversation of state.conversations) {
          conversationListEl.appendChild(buildConversationButton(conversation));
        }
        queueMatrixResize();
        return;
      }

      const groups = new Map();
      for (const conversation of state.conversations) {
        const key = conversation.parentSpaceTitle || (conversation.kind === 'team' ? (conversation.displayTitle || conversation.title || conversation.id) : 'other');
        if (!groups.has(key)) {
          groups.set(key, []);
        }
        groups.get(key).push(conversation);
      }

      for (const [groupTitle, conversations] of groups.entries()) {
        const group = document.createElement('section');
        group.className = 'conversation-group';
        const title = document.createElement('div');
        title.className = 'conversation-group-title';
        title.textContent = groupTitle === 'other' ? 'Other conversations' : groupTitle;
        group.appendChild(title);
        const items = document.createElement('div');
        items.className = 'conversation-group-items';
        for (const conversation of conversations) {
          items.appendChild(buildConversationButton(conversation));
        }
        group.appendChild(items);
        conversationListEl.appendChild(group);
      }
      queueMatrixResize();
    }

    function buildConversationButton(conversation) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = `conversation ${conversation.id === state.selectedConversationId ? 'active' : ''}`;
      const displayTitle = conversation.displayTitle || conversation.title || conversation.id;
      const metaParts = [];
      if (conversation.matchCount) {
        metaParts.push(`${conversation.matchCount} matches`);
      }
      if (conversation.channelTitle && conversation.channelTitle !== displayTitle) {
        metaParts.push(conversation.channelTitle);
      }
      metaParts.push(`${conversation.messageCount || 0} messages`);
      if (conversation.productThreadType || conversation.threadType) {
        metaParts.push(conversation.productThreadType || conversation.threadType || '');
      }
      let relationLine = '';
      if (conversation.parentSpaceTitle) {
        relationLine = `
          <div class="meta-line relation-line" title="Parent team: ${escapeHtml(conversation.parentSpaceTitle)}">
            <span class="flag relation-flag">team</span>
            <span class="relation-name">${escapeHtml(conversation.parentSpaceTitle)}</span>
          </div>
        `;
      } else if (conversation.channelTitle && conversation.channelTitle !== displayTitle && conversation.kind === 'team') {
        relationLine = `
          <div class="meta-line relation-line" title="Primary channel: ${escapeHtml(conversation.channelTitle)}">
            <span class="flag relation-flag">channel</span>
            <span class="relation-name">${escapeHtml(conversation.channelTitle)}</span>
          </div>
        `;
      }
      const flags = [];
      if (conversation.hidden) flags.push('<span class="flag">hidden</span>');
      if (conversation.meeting) flags.push('<span class="flag">meeting</span>');
      if (conversation.kind) flags.push(`<span class="flag">${escapeHtml(conversation.kind)}</span>`);
      if (conversation.status) flags.push(`<span class="flag">${escapeHtml(conversation.status)}</span>`);
      if (conversation.error) flags.push('<span class="flag warn">error</span>');
      button.innerHTML = `
        <div class="name">${escapeHtml(displayTitle)}</div>
        <div class="meta-line">${escapeHtml(conversation.id || '')}</div>
        ${relationLine}
        <div class="meta-line">${escapeHtml(metaParts.join(' | '))}</div>
        ${conversation.matchPreview ? `<div class="meta-line muted">${escapeHtml(String(conversation.matchPreviewAuthor || '[unknown]'))}: ${escapeHtml(String(conversation.matchPreview).slice(0, 140))}</div>` : ''}
        ${conversation.warning ? `<div class="meta-line muted">${escapeHtml(String(conversation.warning).slice(0, 140))}</div>` : ''}
        ${conversation.error ? `<div class="meta-line warn">${escapeHtml(String(conversation.error).slice(0, 140))}</div>` : ''}
        <div class="flags">${flags.join('')}</div>
      `;
      button.addEventListener('click', async () => {
        state.selectedConversationId = conversation.id;
        renderConversationList();
        await loadChat();
      });
      return button;
    }

    async function loadChat() {
      if (!state.selectedConversationId) return;
      const requestToken = ++state.chatRequestToken;
      if (state.chatAbortController) {
        state.chatAbortController.abort();
      }
      state.chatAbortController = new AbortController();
      const selected = state.conversations.find((item) => item.id === state.selectedConversationId) || null;
      if (selected && selected.error) {
        chatTitleEl.textContent = selected.title || selected.id || 'Failed export';
        chatMetaEl.textContent = `${selected.kind || 'conversation'} | ${selected.productThreadType || selected.threadType || 'unknown type'} | status: ${selected.status}`;
        chatEl.innerHTML = `<article class="msg system"><header><span>export error</span><span>${escapeHtml(selected.id || '')}</span></header><div class="body">${escapeHtml(selected.error || 'Unknown error')}</div></article>`;
        setExportButtonsDisabled(true);
        return;
      }
      const basePayload = {
        conversationId: state.selectedConversationId,
        query: document.getElementById('messageQuery').value.trim() || state.globalSearchQuery,
        author: document.getElementById('authorFilter').value.trim(),
        caseSensitive: document.getElementById('caseSensitive').checked ? 'true' : 'false',
        hideSystem: document.getElementById('hideSystem').checked ? 'true' : 'false',
      };
      const pageSize = 400;
      let offset = 0;
      let payload = null;
      const allMessages = [];
      while (true) {
        const params = new URLSearchParams({
          ...basePayload,
          limit: String(pageSize),
          offset: String(offset),
        });
        try {
          const response = await fetch(`/api/chat?${params.toString()}`, {
            signal: state.chatAbortController.signal,
          });
          payload = await response.json();
        } catch (error) {
          if (error && error.name === 'AbortError') {
            return;
          }
          throw error;
        }
        if (requestToken !== state.chatRequestToken) {
          return;
        }
        if (payload.error) {
          break;
        }
        const chunk = Array.isArray(payload.messages) ? payload.messages : [];
        allMessages.push(...chunk);
        const returned = Number(payload.returned || chunk.length || 0);
        const totalMatches = Number(payload.totalMatches || allMessages.length || 0);
        if (returned <= 0 || allMessages.length >= totalMatches || returned < pageSize) {
          payload.messages = allMessages;
          break;
        }
        offset += returned;
      }
      if (payload.error) {
        chatTitleEl.textContent = 'Load failed';
        chatMetaEl.textContent = payload.error;
        chatEl.innerHTML = '';
        setExportButtonsDisabled(true);
        return;
      }
      const conversation = payload.conversation || {};
      const displayTitle = conversation.displayTitle || conversation.title || payload.meta?.title || 'Teams chat';
      chatTitleEl.textContent = displayTitle;
      const flags = [];
      if (conversation.hidden) flags.push('hidden');
      if (conversation.meeting) flags.push('meeting');
      if (conversation.kind) flags.push(conversation.kind);
      if (conversation.status) flags.push(conversation.status);
      if (conversation.parentSpaceTitle) flags.push(`team: ${conversation.parentSpaceTitle}`);
      if (conversation.kind === 'team' && conversation.channelTitle && conversation.channelTitle !== displayTitle) {
        flags.push(`channel: ${conversation.channelTitle}`);
      }
      const warnings = [];
      if (conversation.warning) warnings.push(conversation.warning);
      if (payload.meta?.warning) warnings.push(payload.meta.warning);
      const globalSearchNote = state.globalSearchQuery && !document.getElementById('messageQuery').value.trim()
        ? ` | global search: ${state.globalSearchQuery}`
        : '';
      chatMetaEl.textContent = `${payload.totalMatches} matching messages${flags.length ? ' | ' + flags.join(', ') : ''}${globalSearchNote}${warnings.length ? ' | warning: ' + warnings[0] : ''}`;
      renderMessages(sortMessagesNewestFirst(payload.messages || []));
      setExportButtonsDisabled(false);
    }

    function scheduleLoadChat() {
      if (state.chatSearchTimer) {
        window.clearTimeout(state.chatSearchTimer);
      }
      state.chatSearchTimer = window.setTimeout(() => {
        state.chatSearchTimer = null;
        void loadChat();
      }, CHAT_FILTER_DEBOUNCE_MS);
    }

    function renderNoConversationSelected(title, subtitle) {
      chatTitleEl.textContent = title;
      chatMetaEl.textContent = subtitle;
      chatEl.innerHTML = '<div class="muted">Choose a conversation on the left after export finishes.</div>';
      setExportButtonsDisabled(true);
      queueMatrixResize();
    }

    async function refreshRepositoryData() {
      if (state.repositoryRefreshPromise) {
        return state.repositoryRefreshPromise;
      }
      state.repositoryRefreshPromise = (async () => {
        await loadMeta();
        await loadConversations();
      })().finally(() => {
        state.repositoryRefreshPromise = null;
      });
      return state.repositoryRefreshPromise;
    }

    function renderMessages(messages) {
      chatEl.innerHTML = '';
      if (!messages.length) {
        chatEl.innerHTML = '<div class="muted">No messages matched the current filters.</div>';
        return;
      }
      for (const message of messages) {
        const card = document.createElement('article');
        card.className = `msg ${message.system ? 'system' : ''}`;
        const attachments = (message.attachments || []).map((attachment) => {
          const label = escapeHtml(attachment.label || attachment.href || 'attachment');
          const href = attachment.viewerHref || attachment.downloadHref || attachment.href || '';
          if (href) {
            return `<a href="${escapeHtml(href)}" target="_blank" rel="noreferrer">${label}</a>`;
          }
          return label;
        });
        const reactions = (message.reactions || []).map((reaction) => `${escapeHtml(reaction.emoji || '')}x${reaction.count || 0}`);
        const mentions = (message.mentions || []).map((mention) => escapeHtml(mention.name || mention)).filter(Boolean);
        const extraLines = [];
        if (message.replyTo && message.replyTo.text) {
          extraLines.push(`reply: ${escapeHtml(message.replyTo.author || '[unknown]')} :: ${escapeHtml(message.replyTo.text)}`);
        }
        if (reactions.length) {
          extraLines.push(`reactions: ${reactions.join(' ')}`);
        }
        if (attachments.length) {
          extraLines.push(`attachments: ${attachments.join(', ')}`);
        }
        if (mentions.length) {
          extraLines.push(`mentions: ${mentions.join(', ')}`);
        }
        card.innerHTML = `
          <header>
            <span>${escapeHtml(message.author || '[unknown]')}</span>
            <span>${escapeHtml(message.timestamp || '')}</span>
          </header>
          <div class="body">${escapeHtml(message.text || '[no text]')}</div>
          ${extraLines.length ? `<div class="extra">${extraLines.map((line) => `<div>${line}</div>`).join('')}</div>` : ''}
        `;
        chatEl.appendChild(card);
      }
      queueMatrixResize();
    }

    function sortMessagesNewestFirst(messages) {
      return [...messages]
        .map((message, index) => ({ message, index }))
        .sort((left, right) => {
          const leftTimestamp = Date.parse(left.message?.timestamp || '');
          const rightTimestamp = Date.parse(right.message?.timestamp || '');
          const leftValid = Number.isFinite(leftTimestamp);
          const rightValid = Number.isFinite(rightTimestamp);
          if (leftValid && rightValid && leftTimestamp !== rightTimestamp) {
            return rightTimestamp - leftTimestamp;
          }
          if (leftValid !== rightValid) {
            return leftValid ? -1 : 1;
          }
          return right.index - left.index;
        })
        .map((item) => item.message);
    }

    function queueMatrixResize() {
      if (!matrixCanvasState) return;
      window.requestAnimationFrame(() => resizeMatrixRain());
    }

    function escapeHtml(value) {
      return String(value)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;');
    }

    function setExportButtonsDisabled(disabled) {
      document.querySelectorAll('[data-export-format]').forEach((button) => {
        button.disabled = disabled;
      });
    }

    function initHelpToggle() {
      const toggle = document.getElementById('helpToggle');
      const panel = document.getElementById('helpPanel');
      toggle.addEventListener('click', () => {
        const open = panel.classList.toggle('open');
        toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
      });
    }

    function initControlPanelsToggle() {
      const toggle = document.getElementById('controlPanelsToggle');
      const grid = document.getElementById('controlGrid');
      if (!toggle || !grid) return;
      toggle.addEventListener('click', () => {
        const collapsed = grid.classList.toggle('collapsed');
        toggle.textContent = collapsed ? '▸' : '▾';
        toggle.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
        toggle.setAttribute('title', collapsed ? 'expand export panels' : 'collapse export panels');
      });
    }

    function initSidebarFiltersToggle() {
      const toggle = document.getElementById('sidebarFiltersToggle');
      const panel = document.getElementById('sidebarAdvancedFilters');
      if (!toggle || !panel) return;
      toggle.addEventListener('click', () => {
        const open = panel.classList.toggle('open');
        toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
        toggle.textContent = open ? 'filters ▾' : 'filters ▸';
      });
    }

    function initMatrixRain() {
      const rain = document.getElementById('matrixRain');
      if (!rain) return;
      rain.innerHTML = '';
      if (matrixAnimationFrame) {
        window.cancelAnimationFrame(matrixAnimationFrame);
      }
      if (matrixResizeHandler) {
        window.removeEventListener('resize', matrixResizeHandler);
      }
      if (matrixResizeObserver) {
        matrixResizeObserver.disconnect();
        matrixResizeObserver = null;
      }
      const canvas = document.createElement('canvas');
      rain.appendChild(canvas);
      const context = canvas.getContext('2d');
      if (!context) return;
      matrixCanvasState = {
        host: rain,
        container: rain.parentElement || rain,
        canvas,
        context,
        glyphs: '01<>[]{}()/\\\\|+-*MS_TEAMS_EXPORT',
        columns: [],
        width: 0,
        height: 0,
        fontSize: 14,
      };
      matrixResizeHandler = () => resizeMatrixRain();
      window.addEventListener('resize', matrixResizeHandler, { passive: true });
      if (window.ResizeObserver) {
        matrixResizeObserver = new window.ResizeObserver(() => resizeMatrixRain());
        matrixResizeObserver.observe(matrixCanvasState.container);
      }
      resizeMatrixRain();

      const tick = () => {
        refreshMatrixRain();
        matrixAnimationFrame = window.requestAnimationFrame(tick);
      };
      tick();
    }

    function buildMatrixColumnText(glyphs, length) {
      let text = '';
      for (let index = 0; index < length; index += 1) {
        const char = glyphs[Math.floor(Math.random() * glyphs.length)];
        text += char + (index === length - 1 ? '' : '\\n');
      }
      return text;
    }

    function resizeMatrixRain() {
      const state = matrixCanvasState;
      if (!state) return;
      const rect = state.container.getBoundingClientRect();
      const dpr = window.devicePixelRatio || 1;
      state.width = Math.max(1, Math.floor(rect.width));
      state.height = Math.max(1, Math.floor(rect.height));
      state.canvas.width = Math.max(1, Math.floor(rect.width * dpr));
      state.canvas.height = Math.max(1, Math.floor(rect.height * dpr));
      state.canvas.style.width = `${rect.width}px`;
      state.canvas.style.height = `${rect.height}px`;
      state.context.setTransform(dpr, 0, 0, dpr, 0, 0);
      state.fontSize = rect.width > 420 ? 14 : 12;
      state.columns = buildMatrixColumns(state.width, state.height, state.fontSize);
      state.context.clearRect(0, 0, state.width, state.height);
    }

    function buildMatrixColumns(width, height, fontSize) {
      const spacing = Math.max(14, Math.floor(fontSize * 1.15));
      const count = Math.max(12, Math.floor(width / spacing));
      return Array.from({ length: count }, (_, index) => ({
        x: index * spacing + Math.random() * 4,
        y: Math.random() * height - height,
        speed: 0.42 + Math.random() * 0.7,
        trail: 10 + Math.floor(Math.random() * 16),
        drift: Math.random() * 0.12 - 0.06,
      }));
    }

    function refreshMatrixRain() {
      const state = matrixCanvasState;
      if (!state) return;
      const ctx = state.context;
      const glyphs = state.glyphs;
      ctx.fillStyle = 'rgba(3, 10, 14, 0.10)';
      ctx.fillRect(0, 0, state.width, state.height);
      ctx.font = `${state.fontSize}px "IBM Plex Mono", monospace`;
      ctx.textBaseline = 'top';

      state.columns.forEach((column) => {
        for (let step = 0; step < column.trail; step += 1) {
          const y = column.y - step * (state.fontSize * 0.92);
          if (y < -state.fontSize || y > state.height + state.fontSize) {
            continue;
          }
          const alpha = Math.max(0.05, 1 - (step / column.trail));
          const glyph = glyphs[Math.floor(Math.random() * glyphs.length)];
          if (step === 0) {
            ctx.fillStyle = 'rgba(204, 245, 222, 0.7)';
            ctx.shadowColor = 'rgba(150, 240, 198, 0.42)';
            ctx.shadowBlur = 10;
          } else {
            ctx.fillStyle = `rgba(108, 232, 155, ${0.04 + alpha * 0.2})`;
            ctx.shadowColor = 'rgba(108, 232, 155, 0.14)';
            ctx.shadowBlur = 5 * alpha;
          }
          ctx.fillText(glyph, column.x, y);
        }

        column.y += column.speed;
        column.x += column.drift;
        if (column.y - column.trail * state.fontSize > state.height + 28 || column.x < -20 || column.x > state.width + 20) {
          column.y = -20 - Math.random() * state.height * 0.3;
          column.x = Math.max(0, Math.min(state.width - state.fontSize, column.x + (Math.random() * 18 - 9)));
          column.speed = 0.42 + Math.random() * 0.7;
          column.trail = 10 + Math.floor(Math.random() * 16);
          column.drift = Math.random() * 0.12 - 0.06;
        }
      });
      ctx.shadowBlur = 0;
    }

    async function postControl(url, payload = {}) {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Control-Token': CONTROL_TOKEN,
        },
        body: JSON.stringify(payload),
      });
      return await response.json();
    }

    async function loadControlStatus() {
      const [exportResponse, mirrorResponse] = await Promise.all([
        fetch('/api/control/export/status'),
        fetch('/api/control/mirror/status'),
      ]);
      renderExportStatus(await exportResponse.json());
      renderMirrorStatus(await mirrorResponse.json());
    }

    function renderExportStatus(payload) {
      const snapshot = payload.snapshot || null;
      const result = payload.result || null;
      const running = Boolean(payload.running);
      const completionKey = !running && result
        ? JSON.stringify([
            result.indexPath || '',
            result.outdir || '',
            result.exportedCount || 0,
            result.failedCount || 0,
            result.messageCount || 0,
            Boolean(result.interrupted),
            result.message || '',
          ])
        : '';
      document.getElementById('exportStatusText').textContent = running
        ? (snapshot?.currentStatus || snapshot?.phase || 'running')
        : (result?.message || 'idle');
      document.getElementById('exportStatusMeta').textContent = snapshot
        ? [
            `ok:${snapshot.exportedConversations}`,
            `skip:${snapshot.skippedConversations}`,
            `fail:${snapshot.failedConversations}`,
            `msgs:${snapshot.exportedMessages}`,
            snapshot.currentTitle || snapshot.note || '',
          ].filter(Boolean).join(' | ')
        : (result ? [
            `conversations:${result.conversationCount || 0}`,
            `exported:${result.exportedCount || 0}`,
            `failed:${result.failedCount || 0}`,
            `messages:${result.messageCount || 0}`,
          ].join(' | ') : 'No export job running.');
      setBar('exportAllBar', snapshot?.processedConversations || 0, snapshot?.totalConversations || 0);
      setBar('exportHiddenBar', snapshot?.processedHidden || 0, snapshot?.totalHidden || 0);
      setBar('exportMeetingBar', snapshot?.processedMeeting || 0, snapshot?.totalMeeting || 0);
      document.getElementById('exportAllCount').textContent = `${snapshot?.processedConversations || 0} / ${snapshot?.totalConversations || 0}`;
      document.getElementById('exportHiddenCount').textContent = `${snapshot?.processedHidden || 0} / ${snapshot?.totalHidden || 0}`;
      document.getElementById('exportMeetingCount').textContent = `${snapshot?.processedMeeting || 0} / ${snapshot?.totalMeeting || 0}`;
      document.getElementById('exportStartButton').disabled = running;
      document.getElementById('exportStopButton').disabled = !running;
      document.getElementById('exportForceStopButton').disabled = !running;
      if (running) {
        state.exportCompletionKey = '';
      } else if (completionKey && completionKey !== state.exportCompletionKey) {
        state.exportCompletionKey = completionKey;
        void refreshRepositoryData();
      }
    }

    function renderMirrorStatus(payload) {
      const snapshot = payload.snapshot || null;
      const result = payload.result || null;
      const running = Boolean(payload.running);
      const completionKey = !running && result
        ? JSON.stringify([
            result.bundleRoot || '',
            result.attachmentCount || 0,
            result.mirroredCount || 0,
            result.reusedCount || 0,
            result.failedCount || 0,
            Boolean(result.interrupted),
            Boolean(result.lowDisk),
            result.message || '',
          ])
        : '';
      document.getElementById('mirrorStatusText').textContent = running
        ? (snapshot?.currentStatus || snapshot?.phase || 'running')
        : (result?.message || 'idle');
      document.getElementById('mirrorStatusMeta').textContent = snapshot
        ? [
            `mirrored:${snapshot.mirroredAssets}`,
            `reused:${snapshot.reusedAssets}`,
            `fail:${snapshot.failedAssets}`,
            `data:${formatBytes(snapshot.bytesDownloaded)}`,
            `free:${formatBytes(snapshot.freeBytes)}`,
            `eta:${formatDuration(snapshot.etaSeconds)}`,
            snapshot.currentAssetLabel || snapshot.note || '',
          ].filter(Boolean).join(' | ')
        : (result ? [
            `attachments:${result.attachmentCount || 0}`,
            `mirrored:${result.mirroredCount || 0}`,
            `reused:${result.reusedCount || 0}`,
            `failed:${result.failedCount || 0}`,
            result.freeBytes != null ? `free:${formatBytes(result.freeBytes)}` : '',
          ].filter(Boolean).join(' | ') : 'No mirror job running.');
      setBar('mirrorAssetBar', snapshot?.processedAssets || 0, snapshot?.totalAssets || 0);
      setBar('mirrorChatBar', snapshot?.processedChats || 0, snapshot?.totalChats || 0);
      document.getElementById('mirrorAssetCount').textContent = `${snapshot?.processedAssets || 0} / ${snapshot?.totalAssets || 0}`;
      document.getElementById('mirrorChatCount').textContent = `${snapshot?.processedChats || 0} / ${snapshot?.totalChats || 0}`;
      document.getElementById('mirrorStartButton').disabled = running;
      document.getElementById('mirrorStopButton').disabled = !running;
      document.getElementById('mirrorForceStopButton').disabled = !running;
      if (running) {
        state.mirrorCompletionKey = '';
      } else if (completionKey && completionKey !== state.mirrorCompletionKey) {
        state.mirrorCompletionKey = completionKey;
        void refreshRepositoryData();
      }
    }

    function setBar(id, done, total) {
      const node = document.getElementById(id);
      const percent = total > 0 ? Math.max(0, Math.min(100, (done / total) * 100)) : 0;
      node.style.width = `${percent}%`;
    }

    function formatBytes(value) {
      if (value == null || Number.isNaN(Number(value))) return '?';
      const units = ['B', 'KB', 'MB', 'GB', 'TB'];
      let size = Number(value);
      let unit = units[0];
      for (const candidate of units) {
        unit = candidate;
        if (Math.abs(size) < 1024 || candidate === units[units.length - 1]) break;
        size /= 1024;
      }
      if (unit === 'B') return `${Math.round(size)}${unit}`;
      return `${size.toFixed(1)}${unit}`;
    }

    function formatDuration(value) {
      if (value == null || Number.isNaN(Number(value))) return '?';
      const total = Math.max(0, Math.floor(Number(value)));
      const hours = Math.floor(total / 3600);
      const minutes = Math.floor((total % 3600) / 60);
      const seconds = total % 60;
      if (hours) return `${hours}h${String(minutes).padStart(2, '0')}m`;
      if (minutes) return `${minutes}m${String(seconds).padStart(2, '0')}s`;
      return `${seconds}s`;
    }

    document.getElementById('globalSearchQuery').addEventListener('input', resetGlobalSearchModeAndSchedule);
    document.getElementById('globalSearchQuery').addEventListener('change', resetGlobalSearchModeAndSchedule);
    document.getElementById('globalSearchFull').addEventListener('input', async (event) => {
      const globalQuery = document.getElementById('globalSearchQuery').value.trim();
      if (!globalQuery || globalQuery.length < GLOBAL_SEARCH_MIN_CHARS) {
        event.target.checked = false;
        return;
      }
      state.globalSearchFullMode = Boolean(event.target.checked);
      await loadConversations();
    });
    document.getElementById('globalSearchFull').addEventListener('change', async (event) => {
      const globalQuery = document.getElementById('globalSearchQuery').value.trim();
      if (!globalQuery || globalQuery.length < GLOBAL_SEARCH_MIN_CHARS) {
        event.target.checked = false;
        return;
      }
      state.globalSearchFullMode = Boolean(event.target.checked);
      await loadConversations();
    });
    for (const id of ['conversationQuery', 'caseSensitive', 'hiddenFilter', 'meetingFilter', 'statusFilter', 'kindFilter']) {
      document.getElementById(id).addEventListener('input', scheduleLoadConversations);
      document.getElementById(id).addEventListener('change', scheduleLoadConversations);
    }
    document.getElementById('groupByTeam').addEventListener('change', renderConversationList);
    for (const id of ['messageQuery', 'authorFilter']) {
      document.getElementById(id).addEventListener('input', scheduleLoadChat);
      document.getElementById(id).addEventListener('change', scheduleLoadChat);
    }
    document.getElementById('hideSystem').addEventListener('input', async () => {
      if (document.getElementById('globalSearchQuery').value.trim()) {
        await loadConversations();
        return;
      }
      await loadChat();
    });
    document.getElementById('hideSystem').addEventListener('change', async () => {
      if (document.getElementById('globalSearchQuery').value.trim()) {
        await loadConversations();
        return;
      }
      await loadChat();
    });
    document.getElementById('reloadButton').addEventListener('click', loadChat);
    document.getElementById('exportStartButton').addEventListener('click', async () => {
      const maxChatsRaw = document.getElementById('exportMaxChats').value.trim();
      const payload = { skipExisting: document.getElementById('exportSkipExisting').checked };
      if (maxChatsRaw) payload.maxChats = Number(maxChatsRaw);
      await postControl('/api/control/export/start', payload);
      await loadControlStatus();
    });
    document.getElementById('exportStopButton').addEventListener('click', async () => {
      await postControl('/api/control/export/stop', { force: false });
      await loadControlStatus();
    });
    document.getElementById('exportForceStopButton').addEventListener('click', async () => {
      await postControl('/api/control/export/stop', { force: true });
      await loadControlStatus();
    });
    document.getElementById('mirrorStartButton').addEventListener('click', async () => {
      const minFreeGb = Number(document.getElementById('mirrorMinFreeGb').value || '30');
      await postControl('/api/control/mirror/start', { minFreeGb });
      await loadControlStatus();
    });
    document.getElementById('mirrorStopButton').addEventListener('click', async () => {
      await postControl('/api/control/mirror/stop', { force: false });
      await loadControlStatus();
    });
    document.getElementById('mirrorForceStopButton').addEventListener('click', async () => {
      await postControl('/api/control/mirror/stop', { force: true });
      await loadControlStatus();
    });
    document.querySelectorAll('[data-export-format]').forEach((button) => {
      button.addEventListener('click', () => {
        if (!state.selectedConversationId) return;
        const format = button.getAttribute('data-export-format');
        window.location.href = `/api/export?conversationId=${encodeURIComponent(state.selectedConversationId)}&format=${encodeURIComponent(format)}`;
      });
    });

    async function bootstrap() {
      initHelpToggle();
      initControlPanelsToggle();
      initSidebarFiltersToggle();
      initMatrixRain();
      await loadMeta();
      await loadConversations();
      await loadControlStatus();
      statusPoller = window.setInterval(async () => {
        try {
          await loadControlStatus();
          await loadMeta();
        } catch (_error) {
        }
      }, 1500);
    }

    bootstrap();
  </script>
</body>
</html>"""
    html = html.replace('"__CONTROL_TOKEN__"', json.dumps(control_token))
    html = html.replace("__GLOBAL_SEARCH_LIMIT__", str(GLOBAL_SEARCH_RESULT_LIMIT))
    return html.replace("__BUILD_LABEL__", build_label)
