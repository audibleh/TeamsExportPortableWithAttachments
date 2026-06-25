from __future__ import annotations

from dataclasses import dataclass
import base64
import hashlib
import os
from pathlib import Path
import re
from typing import Any
from urllib.parse import urlparse

from msteams_export.browser.session import DEFAULT_TEAMS_URL
from msteams_export.config import detect_project_paths
from msteams_export.exporters.teams_browser import TeamsBrowserRequest, resolve_browser_target
from msteams_export.models import Attachment, ExportBundle, ExportMessage
from msteams_export.viewer.render import ViewOptions
from msteams_export.webapp.attachments import normalize_attachment_url


IMAGE_TYPES = {"png", "jpg", "jpeg", "gif", "bmp", "webp", "svg", "heic", "heif"}
RESET = "\033[0m"
NEON_CYAN = "\033[96m"
NEON_MAGENTA = "\033[95m"
NEON_YELLOW = "\033[93m"
DIM = "\033[2m"


@dataclass(slots=True)
class PreviewImagesRequest:
    input_path: Path
    limit: int = 3
    author: str | None = None
    query: str | None = None
    browser_name: str = "auto"
    profile_path: Path | None = None
    teams_url: str = DEFAULT_TEAMS_URL
    timeout_ms: int = 30_000
    mode: str = "auto"


@dataclass(slots=True)
class ImagePreviewItem:
    message_id: str
    author: str
    timestamp: str
    label: str
    href: str
    attachment_type: str | None


@dataclass(slots=True)
class PreviewImagesResult:
    ok: bool
    message: str
    rendered_output: str
    preview_dir: Path | None = None
    preview_count: int = 0
    browser_name: str | None = None
    executable_path: Path | None = None
    profile_path: Path | None = None


def preview_images_from_export(request: PreviewImagesRequest) -> PreviewImagesResult:
    bundle = ExportBundle.load(request.input_path)
    items = collect_image_preview_items(
        bundle,
        ViewOptions(limit=max(1, request.limit), author=request.author, query=request.query),
    )
    if not items:
        output = "\n".join(
            [
                _banner("Image Preview"),
                f"{NEON_YELLOW}No image attachments matched the current filters.{RESET}",
            ]
        )
        return PreviewImagesResult(ok=True, message="No image attachments found.", rendered_output=output)

    preview_dir = (detect_project_paths().root / ".state" / "previews").resolve()
    preview_dir.mkdir(parents=True, exist_ok=True)

    try:
        target = resolve_browser_target(
            TeamsBrowserRequest(
                browser_name=request.browser_name,
                profile_path=request.profile_path,
                teams_url=request.teams_url,
                headless=True,
                timeout_ms=request.timeout_ms,
            )
        )
    except Exception as exc:
        output = render_preview_listing(items, preview_dir=preview_dir, protocol=None, failures=[str(exc)] * len(items))
        return PreviewImagesResult(
            ok=False,
            message=str(exc),
            rendered_output=output,
            preview_dir=preview_dir,
        )

    protocol = _detect_terminal_protocol(request.mode)
    previews: list[Path | None] = []
    failures: list[str | None] = []
    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as playwright:
            context = playwright.chromium.launch_persistent_context(
                user_data_dir=str(target.profile_path),
                executable_path=str(target.executable_path),
                headless=True,
                args=[
                    "--no-first-run",
                    "--no-default-browser-check",
                ],
            )
            try:
                bootstrap = context.pages[0] if context.pages else context.new_page()
                bootstrap.goto(target.teams_url, wait_until="domcontentloaded", timeout=target.timeout_ms)
                bootstrap.wait_for_timeout(2_500)
                for item in items:
                    preview_path = preview_dir / _preview_filename(item)
                    try:
                        _capture_attachment_preview(context, item.href, preview_path, target.timeout_ms)
                        previews.append(preview_path)
                        failures.append(None)
                    except Exception as exc:
                        previews.append(None)
                        failures.append(str(exc))
            finally:
                context.close()
    except Exception as exc:
        failures = [str(exc)] * len(items)
        previews = [None] * len(items)

    output = render_preview_listing(items, preview_dir=preview_dir, protocol=protocol, previews=previews, failures=failures)
    preview_count = sum(1 for preview in previews if preview is not None and preview.exists())
    return PreviewImagesResult(
        ok=True,
        message=f"Prepared {preview_count} image previews in {preview_dir}",
        rendered_output=output,
        preview_dir=preview_dir,
        preview_count=preview_count,
        browser_name=target.browser_name,
        executable_path=target.executable_path,
        profile_path=target.profile_path,
    )


def collect_image_preview_items(bundle: ExportBundle, options: ViewOptions) -> list[ImagePreviewItem]:
    results: list[ImagePreviewItem] = []
    for message in bundle.messages:
        if options.author and message.author.lower() != options.author.lower():
            continue
        if options.query and options.query.lower() not in message.text.lower():
            continue
        for attachment in message.attachments:
            if not _attachment_is_image(attachment):
                continue
            if not attachment.href:
                continue
            results.append(
                ImagePreviewItem(
                    message_id=message.identifier,
                    author=message.author or "[unknown]",
                    timestamp=message.timestamp or "",
                    label=attachment.label or attachment.href or "image",
                    href=normalize_attachment_url(attachment.href),
                    attachment_type=attachment.type,
                )
            )
            if len(results) >= options.limit:
                return results
    return results


def render_preview_listing(
    items: list[ImagePreviewItem],
    *,
    preview_dir: Path,
    protocol: str | None,
    previews: list[Path | None] | None = None,
    failures: list[str | None] | None = None,
) -> str:
    previews = previews or [None] * len(items)
    failures = failures or [None] * len(items)
    lines = [_banner("Image Preview")]
    if protocol is None:
        lines.append(f"{DIM}inline terminal image protocol not detected; showing saved preview paths instead{RESET}")
    else:
        lines.append(f"{DIM}inline preview protocol: {protocol}{RESET}")
    for index, item in enumerate(items, start=1):
        preview_path = previews[index - 1] if index - 1 < len(previews) else None
        failure = failures[index - 1] if index - 1 < len(failures) else None
        lines.append("")
        lines.append(f"{NEON_MAGENTA}[{index:03d}]{RESET} {item.author}")
        lines.append(f"  {DIM}{item.timestamp or 'unknown time'}{RESET}")
        lines.append(f"  {NEON_CYAN}label{RESET}: {item.label}")
        lines.append(f"  {NEON_CYAN}url{RESET}: {item.href}")
        if preview_path and preview_path.exists():
            lines.append(f"  {NEON_CYAN}preview{RESET}: {preview_path}")
            if protocol is not None:
                rendered = _render_inline_image(preview_path, protocol)
                if rendered:
                    lines.append(rendered)
        else:
            reason = failure or "preview unavailable"
            lines.append(f"  {NEON_YELLOW}preview{RESET}: {reason}")
    if preview_dir.exists():
        lines.append("")
        lines.append(f"{NEON_CYAN}preview-dir{RESET}: {preview_dir}")
    return "\n".join(lines)


def _capture_attachment_preview(context: Any, url: str, output_path: Path, timeout_ms: int) -> None:
    page = context.new_page()
    try:
        page.goto(url, wait_until="load", timeout=timeout_ms)
        page.wait_for_timeout(1_200)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        image_locator = page.locator("img").first
        if image_locator.count():
            image_locator.screenshot(path=str(output_path))
        else:
            page.screenshot(path=str(output_path), full_page=True)
    finally:
        page.close()


def _render_inline_image(path: Path, protocol: str) -> str | None:
    try:
        raw = path.read_bytes()
    except OSError:
        return None
    payload = base64.b64encode(raw).decode("ascii")
    if protocol == "iterm2":
        name = base64.b64encode(path.name.encode("utf-8")).decode("ascii")
        return f"\033]1337;File=name={name};inline=1;width=40:{payload}\a"
    if protocol == "kitty":
        return f"\033_Gf=100,a=T,t=d;{payload}\033\\"
    return None


def _detect_terminal_protocol(mode: str) -> str | None:
    if mode == "files":
        return None
    term_program = os.environ.get("TERM_PROGRAM", "")
    if term_program == "iTerm.app":
        return "iterm2"
    term = os.environ.get("TERM", "")
    if "kitty" in term or os.environ.get("KITTY_WINDOW_ID"):
        return "kitty"
    return None


def _attachment_is_image(attachment: Attachment) -> bool:
    kind = (attachment.type or "").strip().lower().lstrip(".")
    if kind in IMAGE_TYPES:
        return True
    for candidate in [attachment.label, attachment.href]:
        suffix = _suffix(candidate)
        if suffix in IMAGE_TYPES:
            return True
    return False


def _suffix(value: str | None) -> str:
    if not value:
        return ""
    parsed = urlparse(value)
    suffix = Path(parsed.path or value).suffix.lower().lstrip(".")
    return suffix


def _preview_filename(item: ImagePreviewItem) -> str:
    stem = re.sub(r"[^A-Za-z0-9._-]+", "-", item.label).strip("-") or "image"
    # Non-security digest used only to disambiguate generated preview filenames.
    digest = hashlib.sha1(item.href.encode("utf-8"), usedforsecurity=False).hexdigest()[:12]
    return f"{digest}-{stem[:80]}.png"


def _banner(title: str) -> str:
    bar = "=" * len(title)
    return f"{NEON_MAGENTA}{bar}\n{title}\n{bar}{RESET}"
