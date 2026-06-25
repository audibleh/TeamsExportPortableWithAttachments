from __future__ import annotations

import base64
from contextlib import contextmanager
from dataclasses import dataclass
from http import HTTPStatus
import mimetypes
from pathlib import Path
import re
import time
from typing import Any, Iterator
from urllib.parse import parse_qsl, quote, urlencode, unquote, urlparse, urlunparse

from msteams_export.browser.session import DEFAULT_TEAMS_URL
from msteams_export.exporters.teams_browser import TeamsBrowserRequest, resolve_browser_target
from msteams_export.polite_mode import DEFAULT_POLITE_MODE, PoliteModePolicy, apply_spacing, compute_retry_delay_ms


_DOWNLOAD_HINT_HOSTS = (
    "sharepoint.com",
    "sharepoint-df.com",
    "sharepoint.us",
    "onedrive.live.com",
    "1drv.ms",
)
_FILENAME_STAR_RE = re.compile(r"filename\*=UTF-8''(?P<name>[^;]+)", re.IGNORECASE)
_FILENAME_RE = re.compile(r'filename="?(?P<name>[^";]+)"?', re.IGNORECASE)


class AttachmentUnauthorizedError(RuntimeError):
    pass


@dataclass(slots=True)
class AttachmentDownload:
    body: bytes
    content_type: str
    filename: str
    source_url: str
    normalized_url: str


@dataclass(slots=True)
class AttachmentFetchSession:
    request_context: Any
    browser_context: Any
    timeout_ms: int
    polite_mode: PoliteModePolicy
    last_request_at: float | None = None

    def fetch(self, *, url: str, label: str | None = None) -> AttachmentDownload:
        result = _fetch_with_request_context(
            self.request_context,
            browser_context=self.browser_context,
            url=url,
            label=label,
            timeout_ms=self.timeout_ms,
            polite_mode=self.polite_mode,
            last_request_at=self.last_request_at,
        )
        self.last_request_at = time.monotonic()
        return result


def normalize_attachment_url(url: str) -> str:
    text = url.strip()
    parsed = urlparse(text)
    host = (parsed.hostname or "").lower()
    if not host or not _is_download_hint_host(host):
        return text

    query_items = parse_qsl(parsed.query, keep_blank_values=True)
    normalized_items: list[tuple[str, str]] = []
    has_download = False
    for key, value in query_items:
        key_lower = key.lower()
        if key_lower == "web" and value == "1":
            continue
        if key_lower == "download":
            has_download = True
            normalized_items.append((key, "1"))
            continue
        normalized_items.append((key, value))
    if not has_download:
        normalized_items.append(("download", "1"))
    return urlunparse(parsed._replace(query=urlencode(normalized_items, doseq=True)))


def build_viewer_attachment_href(conversation_id: str, message_id: str, attachment_index: int) -> str:
    params = urlencode(
        {
            "conversationId": conversation_id,
            "messageId": message_id,
            "attachmentIndex": str(max(0, attachment_index)),
        }
    )
    return f"/api/attachment?{params}"


def fetch_attachment(
    *,
    url: str,
    label: str | None = None,
    browser_name: str = "auto",
    profile_path: Path | None = None,
    teams_url: str = DEFAULT_TEAMS_URL,
    timeout_ms: int = 30_000,
) -> AttachmentDownload:
    with open_attachment_fetch_session(
        browser_name=browser_name,
        profile_path=profile_path,
        teams_url=teams_url,
        timeout_ms=timeout_ms,
    ) as session:
        return session.fetch(url=url, label=label)


@contextmanager
def open_attachment_fetch_session(
    *,
    browser_name: str = "auto",
    profile_path: Path | None = None,
    teams_url: str = DEFAULT_TEAMS_URL,
    timeout_ms: int = 30_000,
) -> Iterator[AttachmentFetchSession]:
    target = resolve_browser_target(
        TeamsBrowserRequest(
            browser_name=browser_name,
            profile_path=profile_path,
            teams_url=teams_url,
            headless=True,
            timeout_ms=timeout_ms,
        )
    )

    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        browser_context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(target.profile_path),
            executable_path=str(target.executable_path),
            headless=True,
            args=[
                "--no-first-run",
                "--no-default-browser-check",
            ],
        )
        try:
            bootstrap = browser_context.pages[0] if browser_context.pages else browser_context.new_page()
            bootstrap.goto(target.teams_url, wait_until="domcontentloaded", timeout=target.timeout_ms)
            bootstrap.wait_for_timeout(1_200)
            request_context = playwright.request.new_context(
                storage_state=browser_context.storage_state(),
                user_agent=bootstrap.evaluate("() => navigator.userAgent"),
                ignore_https_errors=True,
            )
            try:
                yield AttachmentFetchSession(
                    request_context=request_context,
                    browser_context=browser_context,
                    timeout_ms=target.timeout_ms,
                    polite_mode=DEFAULT_POLITE_MODE,
                )
            finally:
                request_context.dispose()
        finally:
            browser_context.close()


def is_inline_content_type(content_type: str) -> bool:
    normalized = (content_type or "").split(";", 1)[0].strip().lower()
    return normalized.startswith("image/") or normalized in {
        "application/pdf",
        "text/plain",
        "text/markdown",
    }


def _is_download_hint_host(host: str) -> bool:
    return any(host == suffix or host.endswith(f".{suffix}") for suffix in _DOWNLOAD_HINT_HOSTS)


def _candidate_urls(url: str) -> list[str]:
    normalized = normalize_attachment_url(url)
    parsed = urlparse(url)
    candidates = [url]
    if normalized not in candidates:
        candidates.append(normalized)
    if _is_download_hint_host((parsed.hostname or "").lower()) and parsed.scheme and parsed.netloc:
        layouts_full = (
            f"{parsed.scheme}://{parsed.netloc}/_layouts/15/download.aspx?SourceUrl={quote(url, safe='')}"
        )
        layouts_path = (
            f"{parsed.scheme}://{parsed.netloc}/_layouts/15/download.aspx?SourceUrl={quote(parsed.path, safe='/')}"
        )
        for candidate in [layouts_full, layouts_path]:
            if candidate not in candidates:
                candidates.append(candidate)
    return candidates


def _fetch_with_request_context(
    request_context: Any,
    *,
    browser_context: Any,
    url: str,
    label: str | None,
    timeout_ms: int,
    polite_mode: PoliteModePolicy,
    last_request_at: float | None,
) -> AttachmentDownload:
    errors: list[str] = []
    for candidate_url in _candidate_urls(url):
        for attempt in range(polite_mode.retry_limit + 1):
            last_request_at = apply_spacing(last_request_at, polite_mode.attachment_spacing_ms)
            response = request_context.get(
                candidate_url,
                timeout=timeout_ms,
                fail_on_status_code=False,
                headers={"Accept": "*/*"},
            )
            last_request_at = time.monotonic()
            headers = {str(key).lower(): str(value) for key, value in response.headers.items()}
            body = response.body()
            content_type = _resolve_content_type(headers.get("content-type"), label, candidate_url)
            if response.status == HTTPStatus.UNAUTHORIZED and _is_asm_storage_url(candidate_url):
                browser_result = _fetch_with_browser_page(
                    browser_context,
                    url=candidate_url,
                    label=label,
                    timeout_ms=timeout_ms,
                )
                if browser_result is not None:
                    return browser_result
            if response.status in polite_mode.retry_statuses:
                if attempt >= polite_mode.retry_limit:
                    errors.append(f"{candidate_url} => {response.status} {response.status_text} after retries".strip())
                    break
                time.sleep(
                    compute_retry_delay_ms(
                        attempt=attempt,
                        retry_after_header=headers.get("retry-after"),
                        policy=polite_mode,
                    )
                    / 1000
                )
                continue
            if response.status >= 400:
                if response.status == HTTPStatus.UNAUTHORIZED:
                    raise AttachmentUnauthorizedError(
                        "Attachment unauthorized. The current Teams session could not access this media object. "
                        f"Tried: {candidate_url} => {response.status} {response.status_text}".strip()
                    )
                errors.append(f"{candidate_url} => {response.status} {response.status_text}".strip())
                break
            if _looks_like_http_header_dump(body):
                errors.append(f"{candidate_url} => header-dump response")
                break
            if _looks_like_html_error_page(body, content_type):
                errors.append(f"{candidate_url} => HTML error page")
                break
            filename = _resolve_filename(
                label=label,
                url=candidate_url,
                content_disposition=headers.get("content-disposition"),
                content_type=content_type,
            )
            return AttachmentDownload(
                body=body,
                content_type=content_type,
                filename=filename,
                source_url=url,
                normalized_url=candidate_url,
            )
    detail = "; ".join(errors[:4]) or "No working download variant was found."
    raise RuntimeError(
        "Attachment could not be retrieved. The original Teams/SharePoint file may already be missing, "
        f"moved to a different tenant, or no longer accessible from this session. Tried: {detail}"
    )


def _fetch_with_browser_page(
    browser_context: Any,
    *,
    url: str,
    label: str | None,
    timeout_ms: int,
) -> AttachmentDownload | None:
    created_page = not bool(browser_context.pages)
    page = browser_context.pages[0] if browser_context.pages else browser_context.new_page()
    try:
        fetch_result = _fetch_with_page_context_fetch(page, url=url, label=label, timeout_ms=timeout_ms)
        if fetch_result is not None:
            return fetch_result
        try:
            response = page.goto(url, wait_until="commit", timeout=timeout_ms)
        except Exception as exc:
            if _is_asm_storage_url(url) and "ERR_HTTP_RESPONSE_CODE_FAILURE" in str(exc):
                raise AttachmentUnauthorizedError(
                    "Attachment access was blocked during browser fetch. "
                    "The current Teams session likely does not have permission to open this media object. "
                    f"Tried: {url} => browser navigation returned HTTP response failure"
                ) from exc
            raise
        if response is None:
            return None
        headers = {str(key).lower(): str(value) for key, value in response.headers.items()}
        if response.status == HTTPStatus.UNAUTHORIZED:
            raise AttachmentUnauthorizedError(
                "Attachment unauthorized. The current Teams session could not access this media object. "
                f"Tried: {url} => {response.status} {response.status_text}".strip()
            )
        if response.status >= 400:
            return None
        body = response.body()
        content_type = _resolve_content_type(headers.get("content-type"), label, url)
        if _looks_like_http_header_dump(body):
            return None
        if _looks_like_html_error_page(body, content_type):
            return None
        filename = _resolve_filename(
            label=label,
            url=url,
            content_disposition=headers.get("content-disposition"),
            content_type=content_type,
        )
        return AttachmentDownload(
            body=body,
            content_type=content_type,
            filename=filename,
            source_url=url,
            normalized_url=url,
        )
    finally:
        if created_page:
            page.close()


def _is_asm_storage_url(url: str) -> bool:
    host = (urlparse(url).hostname or "").lower()
    return host == "asm.skype.com" or host.endswith(".asm.skype.com")


def _fetch_with_page_context_fetch(
    page: Any,
    *,
    url: str,
    label: str | None,
    timeout_ms: int,
) -> AttachmentDownload | None:
    try:
        result = page.evaluate(
            """async ({ url }) => {
                try {
                    const response = await fetch(url, {
                        credentials: 'include',
                        headers: { Accept: '*/*' },
                    });
                    const headers = {};
                    for (const [key, value] of response.headers.entries()) {
                        headers[String(key).toLowerCase()] = String(value);
                    }
                    if (!response.ok) {
                        return {
                            ok: false,
                            status: response.status,
                            statusText: response.statusText,
                            headers,
                        };
                    }
                    const buffer = await response.arrayBuffer();
                    const bytes = new Uint8Array(buffer);
                    let binary = '';
                    const chunkSize = 0x8000;
                    for (let index = 0; index < bytes.length; index += chunkSize) {
                        binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
                    }
                    return {
                        ok: true,
                        status: response.status,
                        statusText: response.statusText,
                        headers,
                        bodyBase64: btoa(binary),
                    };
                } catch (error) {
                    return {
                        ok: false,
                        error: String(error),
                    };
                }
            }""",
            {"url": url},
        )
    except Exception:
        return None
    if not isinstance(result, dict):
        return None
    if result.get("error"):
        return None
    status = int(result.get("status", 0) or 0)
    headers = result.get("headers") if isinstance(result.get("headers"), dict) else {}
    status_text = str(result.get("statusText") or "")
    if status == HTTPStatus.UNAUTHORIZED:
        raise AttachmentUnauthorizedError(
            "Attachment unauthorized. The current Teams session could not access this media object. "
            f"Tried: {url} => {status} {status_text}".strip()
        )
    if not result.get("ok"):
        return None
    body_base64 = str(result.get("bodyBase64") or "")
    if not body_base64:
        return None
    body = base64.b64decode(body_base64)
    content_type = _resolve_content_type(str(headers.get("content-type") or ""), label, url)
    if _looks_like_http_header_dump(body):
        return None
    if _looks_like_html_error_page(body, content_type):
        return None
    filename = _resolve_filename(
        label=label,
        url=url,
        content_disposition=str(headers.get("content-disposition") or ""),
        content_type=content_type,
    )
    return AttachmentDownload(
        body=body,
        content_type=content_type,
        filename=filename,
        source_url=url,
        normalized_url=url,
    )


def _resolve_filename(
    *,
    label: str | None,
    url: str,
    content_disposition: str | None,
    content_type: str,
) -> str:
    header_filename = _filename_from_content_disposition(content_disposition)
    if header_filename:
        return header_filename

    cleaned_label = (label or "").strip()
    if cleaned_label:
        return cleaned_label

    parsed = urlparse(url)
    path_name = Path(unquote(parsed.path)).name
    if path_name:
        return path_name

    guessed_ext = mimetypes.guess_extension(content_type.split(";", 1)[0].strip()) or ""
    return f"attachment{guessed_ext}"


def _filename_from_content_disposition(value: str | None) -> str | None:
    if not value:
        return None
    star_match = _FILENAME_STAR_RE.search(value)
    if star_match:
        return unquote(star_match.group("name"))
    basic_match = _FILENAME_RE.search(value)
    if basic_match:
        return basic_match.group("name")
    return None


def _resolve_content_type(content_type: str | None, label: str | None, url: str) -> str:
    cleaned = (content_type or "").split(";", 1)[0].strip().lower()
    if cleaned and cleaned != "application/octet-stream":
        return cleaned

    guessed, _ = mimetypes.guess_type(label or url)
    if guessed:
        return guessed
    return "application/octet-stream"


def _looks_like_http_header_dump(body: bytes) -> bool:
    sample = body[:200].decode("utf-8", errors="ignore").strip()
    return sample.startswith("HTTP/1.") and "Server:" in sample


def _looks_like_html_error_page(body: bytes, content_type: str) -> bool:
    normalized = (content_type or "").split(";", 1)[0].strip().lower()
    if normalized not in {"text/html", "text/plain"}:
        return False
    sample = body[:4000].decode("utf-8", errors="ignore").lower()
    return "file not found" in sample or "sorry, something went wrong" in sample
