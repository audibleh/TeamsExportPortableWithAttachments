from __future__ import annotations

import argparse
from contextlib import contextmanager
from pathlib import Path
import signal
import sys

from msteams_export.browser.session import DEFAULT_TEAMS_URL, check_session, open_session
from msteams_export.exporters.attachment_mirror import (
    AttachmentMirrorProgress,
    MirrorAttachmentsRequest,
    MirrorStopController,
    mirror_bundle_attachments,
)
from msteams_export.exporters.json_export import ExportRequest, plan_export
from msteams_export.exporters.teams_chat import ChatExportRequest, export_chat_to_json
from msteams_export.exporters.teams_conversations import (
    ConversationListRequest,
    ExportAllRequest,
    ExportProgress,
    ExportStopController,
    export_all_conversations,
    list_conversations,
)
from msteams_export.models import ExportBundle
from msteams_export.viewer.image_preview import PreviewImagesRequest, preview_images_from_export
from msteams_export.viewer.render import ViewOptions, render_messages, render_summary
from msteams_export.webapp.server import ViewerServeRequest, resolve_export_root, serve_viewer


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="teams-export",
        description="Python-first CLI for Microsoft Teams export and JSON inspection.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    session_parser = subparsers.add_parser("session-check", help="Check browser automation readiness.")
    session_parser.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome", "chromium"],
        default="auto",
        help="Browser executable to use for the probe.",
    )
    session_parser.add_argument("--profile", type=Path, help="Path to a persistent browser profile.")
    session_parser.add_argument("--url", type=str, default=DEFAULT_TEAMS_URL, help="Teams URL to open.")
    session_parser.add_argument(
        "--timeout-ms",
        type=int,
        default=15_000,
        help="Navigation timeout for the session probe.",
    )
    session_parser.add_argument(
        "--headless",
        action="store_true",
        help="Run the probe in headless mode.",
    )
    session_parser.set_defaults(func=_cmd_session_check)

    session_open_parser = subparsers.add_parser(
        "session-open",
        help="Open a persistent browser session for manual Teams sign-in.",
    )
    session_open_parser.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome", "chromium"],
        default="auto",
        help="Browser executable to use for the interactive session.",
    )
    session_open_parser.add_argument("--profile", type=Path, help="Path to a persistent browser profile.")
    session_open_parser.add_argument("--url", type=str, default=DEFAULT_TEAMS_URL, help="Teams URL to open.")
    session_open_parser.add_argument(
        "--also-url",
        action="append",
        default=[],
        metavar="URL",
        help="Open an extra tab (repeatable), e.g. your SharePoint/OneDrive root, "
        "so the same profile also captures those sign-in cookies for attachment downloads.",
    )
    session_open_parser.set_defaults(func=_cmd_session_open)

    export_parser = subparsers.add_parser("export", help="Planned export commands.")
    export_subparsers = export_parser.add_subparsers(dest="target", required=True)

    export_chat = export_subparsers.add_parser("chat", help="Export the active Teams chat to JSON.")
    export_chat.add_argument("--output", type=Path, required=True, help="Target JSON file.")
    export_chat.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome", "chromium"],
        default="auto",
        help="Browser executable to use for export.",
    )
    export_chat.add_argument("--profile", type=Path, help="Path to a persistent browser profile.")
    export_chat.add_argument("--url", type=str, default=DEFAULT_TEAMS_URL, help="Teams URL to open.")
    export_chat.add_argument("--conversation-id", type=str, help="Export a specific conversation ID directly.")
    export_chat.add_argument("--title", type=str, help="Optional title override for direct conversation export.")
    export_chat.add_argument(
        "--timeout-ms",
        type=int,
        default=30_000,
        help="Navigation timeout for the export probe.",
    )
    export_chat.add_argument(
        "--headed",
        action="store_true",
        help="Run export with a visible browser window instead of headless mode.",
    )
    export_chat.set_defaults(func=_cmd_export)

    export_team = export_subparsers.add_parser("team", help="Plan a team channel export.")
    export_team.add_argument("--output", type=Path, required=True, help="Target JSON file.")
    export_team.set_defaults(func=_cmd_export)

    export_all = export_subparsers.add_parser("all", help="Export all discovered Teams chats to an output directory.")
    export_all.add_argument("--outdir", type=Path, required=True, help="Target directory for index.json and chats/.")
    export_all.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome", "chromium"],
        default="auto",
        help="Browser executable to use for export.",
    )
    export_all.add_argument("--profile", type=Path, help="Path to a persistent browser profile.")
    export_all.add_argument("--url", type=str, default=DEFAULT_TEAMS_URL, help="Teams URL to open.")
    export_all.add_argument(
        "--timeout-ms",
        type=int,
        default=30_000,
        help="Navigation timeout for the export probe.",
    )
    export_all.add_argument(
        "--headed",
        action="store_true",
        help="Run export with a visible browser window instead of headless mode.",
    )
    export_all.add_argument("--max-chats", type=int, help="Optional limit for smoke testing export-all.")
    export_all.add_argument(
        "--skip-existing",
        action="store_true",
        help="Reuse already exported chat files when present.",
    )
    export_all.set_defaults(func=_cmd_export)

    chats_parser = subparsers.add_parser("chats", help="Discover Teams conversations.")
    chats_subparsers = chats_parser.add_subparsers(dest="chats_target", required=True)

    chats_list = chats_subparsers.add_parser("list", help="List all discoverable Teams conversations.")
    chats_list.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome", "chromium"],
        default="auto",
        help="Browser executable to use for discovery.",
    )
    chats_list.add_argument("--profile", type=Path, help="Path to a persistent browser profile.")
    chats_list.add_argument("--url", type=str, default=DEFAULT_TEAMS_URL, help="Teams URL to open.")
    chats_list.add_argument(
        "--timeout-ms",
        type=int,
        default=30_000,
        help="Navigation timeout for the discovery probe.",
    )
    chats_list.add_argument(
        "--headed",
        action="store_true",
        help="Run discovery with a visible browser window instead of headless mode.",
    )
    chats_list.add_argument("--output", type=Path, help="Optional path for the conversations index JSON.")
    chats_list.set_defaults(func=_cmd_chats)

    inspect_parser = subparsers.add_parser("inspect", help="Summarize an export JSON file.")
    inspect_parser.add_argument("input", type=Path, help="Path to export JSON.")
    inspect_parser.set_defaults(func=_cmd_inspect)

    view_parser = subparsers.add_parser("view", help="Render a terminal-friendly chat view.")
    view_parser.add_argument("input", type=Path, help="Path to export JSON.")
    view_parser.add_argument("--limit", type=int, default=10, help="How many messages to render.")
    view_parser.add_argument("--author", type=str, help="Filter by exact author name.")
    view_parser.add_argument("--query", type=str, help="Filter by text substring.")
    view_parser.add_argument(
        "--hide-system",
        action="store_true",
        help="Hide system messages for a cleaner conversational view.",
    )
    view_parser.set_defaults(func=_cmd_view)

    preview_parser = subparsers.add_parser(
        "preview-images",
        help="Best-effort preview for image attachments using the existing Teams browser session.",
    )
    preview_parser.add_argument("input", type=Path, help="Path to export JSON.")
    preview_parser.add_argument("--limit", type=int, default=3, help="How many image attachments to preview.")
    preview_parser.add_argument("--author", type=str, help="Filter by exact author name.")
    preview_parser.add_argument("--query", type=str, help="Filter by text substring.")
    preview_parser.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome", "chromium"],
        default="auto",
        help="Browser executable to use for authenticated preview capture.",
    )
    preview_parser.add_argument("--profile", type=Path, help="Path to a persistent browser profile.")
    preview_parser.add_argument("--url", type=str, default=DEFAULT_TEAMS_URL, help="Teams URL to open.")
    preview_parser.add_argument(
        "--timeout-ms",
        type=int,
        default=30_000,
        help="Navigation timeout for preview capture.",
    )
    preview_parser.add_argument(
        "--mode",
        choices=["auto", "files"],
        default="auto",
        help="Use inline terminal images when supported, otherwise save preview files and print their paths.",
    )
    preview_parser.set_defaults(func=_cmd_preview_images)

    attachments_parser = subparsers.add_parser("attachments", help="Manage mirrored attachment assets.")
    attachments_subparsers = attachments_parser.add_subparsers(dest="attachments_target", required=True)

    attachments_mirror = attachments_subparsers.add_parser(
        "mirror",
        help="Mirror kept attachments into a local assets/ archive for offline viewing.",
    )
    attachments_mirror.add_argument(
        "target",
        type=Path,
        help="Path to export index.json or the export bundle directory containing index.json and chats/.",
    )
    attachments_mirror.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome", "chromium"],
        default="auto",
        help="Browser executable to use for authenticated attachment downloads.",
    )
    attachments_mirror.add_argument("--profile", type=Path, help="Path to a persistent browser profile.")
    attachments_mirror.add_argument("--url", type=str, default=DEFAULT_TEAMS_URL, help="Teams URL to open.")
    attachments_mirror.add_argument(
        "--timeout-ms",
        type=int,
        default=30_000,
        help="Navigation timeout for attachment downloads.",
    )
    attachments_mirror.add_argument(
        "--max-assets",
        type=int,
        help="Optional limit for smoke testing the mirror pass.",
    )
    attachments_mirror.add_argument(
        "--min-free-gb",
        type=float,
        default=30.0,
        help="Stop mirroring before the disk gets too full. Default keeps at least 30 GB free.",
    )
    attachments_mirror.add_argument(
        "--retry-failed",
        action="store_true",
        help="Also re-attempt attachments previously recorded as failed or too-large. "
        "By default a resume skips those to avoid re-walking permanent failures.",
    )
    attachments_mirror.add_argument(
        "--spacing-ms",
        type=int,
        default=400,
        help="Pause between attachment downloads, in milliseconds (default 400). "
        "Lower is faster; raise it if Teams/SharePoint starts throttling (HTTP 429).",
    )
    attachments_mirror.add_argument(
        "--retry-limit",
        type=int,
        help="Maximum retries for throttled (429/503) downloads before giving up. "
        "Lower it to spend less time on stuck items (default uses the polite-mode value).",
    )
    attachments_mirror.set_defaults(func=_cmd_attachments)

    serve_parser = subparsers.add_parser(
        "serve",
        help="Run a local web viewer for an export bundle directory.",
    )
    serve_parser.add_argument(
        "target",
        type=Path,
        help="Path to export index.json or the export bundle directory containing index.json and chats/.",
    )
    serve_parser.add_argument(
        "--host",
        type=str,
        default="127.0.0.1",
        help="Host interface to bind the local viewer to.",
    )
    serve_parser.add_argument(
        "--port",
        type=int,
        default=8765,
        help="TCP port for the local viewer.",
    )
    serve_parser.add_argument(
        "--open-browser",
        action="store_true",
        help="Open the local viewer URL in the default browser.",
    )
    serve_parser.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome", "chromium"],
        default="auto",
        help="Browser executable to use for authenticated attachment fetches.",
    )
    serve_parser.add_argument("--profile", type=Path, help="Path to a persistent browser profile.")
    serve_parser.add_argument("--url", type=str, default=DEFAULT_TEAMS_URL, help="Teams URL to open for attachment fetches.")
    serve_parser.add_argument(
        "--timeout-ms",
        type=int,
        default=30_000,
        help="Navigation timeout for authenticated attachment fetches.",
    )
    serve_parser.set_defaults(func=_cmd_serve)

    archive_parser = subparsers.add_parser(
        "generate-html-archive",
        help="Generate a standalone HTML archive of all exported chats.",
    )
    archive_parser.add_argument(
        "exports_dir",
        type=Path,
        help="Path to the exports directory (containing index.json).",
    )
    archive_parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output HTML file path. Defaults to teams-archive.html next to exports dir.",
    )
    archive_parser.add_argument(
        "--with-images",
        action="store_true",
        help="Write a folder (index.html + images/) with mirrored images shown inline.",
    )
    archive_parser.set_defaults(func=_cmd_generate_html_archive)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return int(args.func(args))


def _cmd_session_check(args: argparse.Namespace) -> int:
    result = check_session(
        browser_name=args.browser,
        profile_path=args.profile,
        teams_url=args.url,
        headless=args.headless,
        timeout_ms=max(1_000, args.timeout_ms),
    )
    print(result.message)
    if result.browser_name:
        print(f"browser: {result.browser_name}")
    if result.executable_path:
        print(f"executable: {result.executable_path}")
    if result.profile_path:
        print(f"profile: {result.profile_path}")
    if result.snapshot:
        print(f"url: {result.snapshot.page_url}")
        print(f"title: {result.snapshot.title}")
        print(f"token-count: {result.snapshot.token_count}")
    for note in result.notes or []:
        print(f"- {note}")
    return 0 if result.ok else 1


def _cmd_session_open(args: argparse.Namespace) -> int:
    result = open_session(
        browser_name=args.browser,
        profile_path=args.profile,
        teams_url=args.url,
        extra_urls=args.also_url,
    )
    print(result.message)
    if result.browser_name:
        print(f"browser: {result.browser_name}")
    if result.executable_path:
        print(f"executable: {result.executable_path}")
    if result.profile_path:
        print(f"profile: {result.profile_path}")
    return 0 if result.ok else 1


def _cmd_export(args: argparse.Namespace) -> int:
    if args.target == "chat":
        result = export_chat_to_json(
            ChatExportRequest(
                output=args.output,
                browser_name=args.browser,
                profile_path=args.profile,
                teams_url=args.url,
                headless=not args.headed,
                timeout_ms=max(1_000, args.timeout_ms),
                conversation_id=args.conversation_id,
                conversation_title=args.title,
            )
        )
        print(result.message)
        if result.browser_name:
            print(f"browser: {result.browser_name}")
        if result.executable_path:
            print(f"executable: {result.executable_path}")
        if result.profile_path:
            print(f"profile: {result.profile_path}")
        if result.output_path:
            print(f"output: {result.output_path}")
        if result.title:
            print(f"title: {result.title}")
        if result.conversation_id:
            print(f"conversation-id: {result.conversation_id}")
        if result.message_count:
            print(f"messages: {result.message_count}")
        return 0 if result.ok else 1

    if args.target == "all":
        progress_renderer = _CliExportProgressRenderer()
        with _export_interrupt_context() as stop_controller:
            result = export_all_conversations(
                ExportAllRequest(
                    outdir=args.outdir,
                    browser_name=args.browser,
                    profile_path=args.profile,
                    teams_url=args.url,
                    headless=not args.headed,
                    timeout_ms=max(1_000, args.timeout_ms),
                    max_chats=args.max_chats,
                    skip_existing=args.skip_existing,
                    progress=progress_renderer,
                    stop_controller=stop_controller,
                )
            )
            progress_renderer.finish()
        print(result.message)
        if result.browser_name:
            print(f"browser: {result.browser_name}")
        if result.executable_path:
            print(f"executable: {result.executable_path}")
        if result.profile_path:
            print(f"profile: {result.profile_path}")
        if result.outdir:
            print(f"outdir: {result.outdir}")
        if result.index_path:
            print(f"index: {result.index_path}")
        print(f"conversations: {result.conversation_count}")
        print(f"exported: {result.exported_count}")
        print(f"failed: {result.failed_count}")
        print(f"hidden: {result.hidden_count}")
        print(f"meeting: {result.meeting_count}")
        print(f"messages: {result.message_count}")
        if result.interrupted:
            return 130
        return 0 if result.ok else 1

    request = ExportRequest(target=args.target, output=args.output)
    print(plan_export(request))
    return 1


def _cmd_chats(args: argparse.Namespace) -> int:
    if args.chats_target != "list":
        raise ValueError(f"Unsupported chats target: {args.chats_target}")

    result = list_conversations(
        ConversationListRequest(
            browser_name=args.browser,
            profile_path=args.profile,
            teams_url=args.url,
            headless=not args.headed,
            timeout_ms=max(1_000, args.timeout_ms),
            output=args.output,
        )
    )
    print(result.message)
    if result.browser_name:
        print(f"browser: {result.browser_name}")
    if result.executable_path:
        print(f"executable: {result.executable_path}")
    if result.profile_path:
        print(f"profile: {result.profile_path}")
    if result.output_path:
        print(f"output: {result.output_path}")
    print(f"conversations: {result.conversation_count}")
    print(f"hidden: {result.hidden_count}")
    print(f"meeting: {result.meeting_count}")
    return 0 if result.ok else 1


def _cmd_inspect(args: argparse.Namespace) -> int:
    bundle = ExportBundle.load(args.input)
    print(render_summary(bundle))
    return 0


def _cmd_view(args: argparse.Namespace) -> int:
    bundle = ExportBundle.load(args.input)
    options = ViewOptions(
        limit=max(1, args.limit),
        author=args.author,
        query=args.query,
        hide_system=args.hide_system,
    )
    print(render_messages(bundle, options))
    return 0


def _cmd_preview_images(args: argparse.Namespace) -> int:
    result = preview_images_from_export(
        PreviewImagesRequest(
            input_path=args.input,
            limit=max(1, args.limit),
            author=args.author,
            query=args.query,
            browser_name=args.browser,
            profile_path=args.profile,
            teams_url=args.url,
            timeout_ms=max(1_000, args.timeout_ms),
            mode=args.mode,
        )
    )
    print(result.rendered_output)
    if result.browser_name:
        print(f"browser: {result.browser_name}")
    if result.executable_path:
        print(f"executable: {result.executable_path}")
    if result.profile_path:
        print(f"profile: {result.profile_path}")
    if result.preview_dir:
        print(f"preview-dir: {result.preview_dir}")
    print(f"preview-count: {result.preview_count}")
    return 0 if result.ok else 1


def _cmd_attachments(args: argparse.Namespace) -> int:
    if args.attachments_target != "mirror":
        raise ValueError(f"Unsupported attachments target: {args.attachments_target}")

    progress_renderer = _CliAttachmentMirrorRenderer()
    with _attachment_interrupt_context() as stop_controller:
        result = mirror_bundle_attachments(
            MirrorAttachmentsRequest(
                target=args.target,
                browser_name=args.browser,
                profile_path=args.profile,
                teams_url=args.url,
                timeout_ms=max(1_000, args.timeout_ms),
                max_assets=args.max_assets,
                min_free_bytes=max(0, int(args.min_free_gb * 1024 * 1024 * 1024)),
                retry_failed=args.retry_failed,
                attachment_spacing_ms=max(0, args.spacing_ms),
                retry_limit=args.retry_limit,
                progress=progress_renderer,
                stop_controller=stop_controller,
            )
        )
    progress_renderer.finish()
    print(result.message)
    if result.bundle_root:
        print(f"bundle: {result.bundle_root}")
    if result.assets_dir:
        print(f"assets: {result.assets_dir}")
    print(f"chats: {result.chat_count}")
    print(f"attachments: {result.attachment_count}")
    print(f"mirrored: {result.mirrored_count}")
    print(f"reused: {result.reused_count}")
    print(f"failed: {result.failed_count}")
    if result.free_bytes is not None:
        print(f"free-space: {_format_bytes(result.free_bytes)}")
        print(f"min-free-space: {_format_bytes(result.min_free_bytes)}")
    if result.interrupted:
        return 130
    return 0 if result.ok else 1


def _cmd_serve(args: argparse.Namespace) -> int:
    try:
        index_path, chats_dir = resolve_export_root(args.target)
    except FileNotFoundError as exc:
        print(_format_serve_target_error(args.target, exc))
        return 1
    except ValueError as exc:
        print(f"Could not start viewer: {exc}")
        return 1

    url = f"http://{args.host}:{args.port}/"
    print("Starting Teams export viewer.")
    print(f"url: {url}")
    print(f"index: {index_path}")
    print(f"chats: {chats_dir}")
    print("Press Ctrl+C to stop the viewer.")
    result = serve_viewer(
        ViewerServeRequest(
            target=args.target,
            host=args.host,
            port=max(1, args.port),
            open_browser=args.open_browser,
            browser_name=args.browser,
            profile_path=args.profile,
            teams_url=args.url,
            timeout_ms=max(1_000, args.timeout_ms),
        )
    )
    print(result.message)
    return 0 if result.ok else 1


def _cmd_generate_html_archive(args: argparse.Namespace) -> int:
    from msteams_export.archive.generate import generate_html_archive, generate_html_folder

    exports_dir = args.exports_dir.expanduser().resolve()
    if not exports_dir.is_dir():
        print(f"Error: {exports_dir} is not a directory.")
        return 1

    if args.with_images:
        output = args.output
        if output is None:
            output = exports_dir.parent / "teams-archive"
        output = output.expanduser().resolve()
        print(f"Reading exports from: {exports_dir}")
        try:
            stats = generate_html_folder(exports_dir, output)
        except FileNotFoundError as exc:
            print(f"Error: {exc}")
            return 1
        except Exception as exc:
            print(f"Error generating archive: {exc}")
            return 1
        print(f"Archive folder written to: {output}")
        print(f"Images copied: {stats['copied']} (missing: {stats['missing']})")
        print(f"Open {output / 'index.html'} in any browser to view your Teams chat history.")
        return 0

    output = args.output
    if output is None:
        output = exports_dir.parent / "teams-archive.html"
    output = output.expanduser().resolve()

    print(f"Reading exports from: {exports_dir}")
    try:
        generate_html_archive(exports_dir, output)
    except FileNotFoundError as exc:
        print(f"Error: {exc}")
        return 1
    except Exception as exc:
        print(f"Error generating archive: {exc}")
        return 1

    size_mb = output.stat().st_size / (1024 * 1024)
    print(f"Archive written to: {output} ({size_mb:.1f} MB)")
    print("Open this file in any browser to view your Teams chat history.")
    return 0


def _format_serve_target_error(target: Path, error: FileNotFoundError) -> str:
    resolved = target.expanduser().resolve()
    bundle_root = resolved if resolved.is_dir() else resolved.parent
    chats_dir = bundle_root / "chats"
    conversations_path = bundle_root / "conversations.json"
    hints: list[str] = [f"Could not start viewer: {error}"]

    if chats_dir.is_dir():
        chat_count = sum(1 for _ in chats_dir.glob("*.json"))
        hints.append(f"Detected chats directory with {chat_count} chat JSON file(s): {chats_dir}")
    if conversations_path.is_file():
        hints.append(f"Detected conversation listing file: {conversations_path}")
    if chats_dir.is_dir() or conversations_path.is_file():
        hints.append("Did a previous export run get interrupted before index.json was written?")
        hints.append(
            f"Try rebuilding the bundle index with: teams-export export all --outdir {bundle_root} --skip-existing"
        )
    else:
        hints.append(
            "Expected an export bundle directory containing index.json and chats/, or a direct path to index.json."
        )
    return "\n".join(hints)


class _CliExportProgressRenderer:
    def __init__(self) -> None:
        self._interactive = bool(getattr(sys.stdout, "isatty", lambda: False)())
        self._last_width = 0

    def __call__(self, snapshot: ExportProgress) -> None:
        line = _format_export_progress(snapshot)
        if self._interactive:
            width = max(self._last_width, len(line))
            sys.stdout.write("\r" + line.ljust(width))
            sys.stdout.flush()
            self._last_width = width
            return
        print(line)

    def finish(self) -> None:
        if self._interactive and self._last_width:
            sys.stdout.write("\n")
            sys.stdout.flush()


class _CliAttachmentMirrorRenderer:
    def __init__(self) -> None:
        self._interactive = bool(getattr(sys.stdout, "isatty", lambda: False)())
        self._last_width = 0

    def __call__(self, snapshot: AttachmentMirrorProgress) -> None:
        line = _format_attachment_progress(snapshot)
        if self._interactive:
            width = max(self._last_width, len(line))
            sys.stdout.write("\r" + line.ljust(width))
            sys.stdout.flush()
            self._last_width = width
            return
        print(line)

    def finish(self) -> None:
        if self._interactive and self._last_width:
            sys.stdout.write("\n")
            sys.stdout.flush()


def _format_export_progress(snapshot: ExportProgress) -> str:
    all_bar = _progress_bar(snapshot.processed_conversations, snapshot.total_conversations)
    hidden_bar = _progress_bar(snapshot.processed_hidden, snapshot.total_hidden)
    meeting_bar = _progress_bar(snapshot.processed_meeting, snapshot.total_meeting)
    status = snapshot.current_status or snapshot.phase
    label = snapshot.current_title or snapshot.current_conversation_id or snapshot.note or "starting"
    compact_label = " ".join(str(label).split())[:48]
    message_delta = f" ({snapshot.current_message_count} msg)" if snapshot.current_message_count else ""
    return (
        f"all {all_bar} {snapshot.processed_conversations}/{snapshot.total_conversations} | "
        f"hidden {hidden_bar} {snapshot.processed_hidden}/{snapshot.total_hidden} | "
        f"meeting {meeting_bar} {snapshot.processed_meeting}/{snapshot.total_meeting} | "
        f"ok:{snapshot.exported_conversations} skip:{snapshot.skipped_conversations} fail:{snapshot.failed_conversations} | "
        f"msgs:{snapshot.exported_messages} | {status}: {compact_label}{message_delta}"
    )


def _format_attachment_progress(snapshot: AttachmentMirrorProgress) -> str:
    asset_bar = _progress_bar(snapshot.processed_assets, snapshot.total_assets)
    chat_bar = _progress_bar(snapshot.processed_chats, snapshot.total_chats)
    status = snapshot.current_status or snapshot.phase
    label = snapshot.current_asset_label or snapshot.current_chat_title or snapshot.note or "starting"
    compact_label = " ".join(str(label).split())[:48]
    eta = _format_duration(snapshot.eta_seconds) if snapshot.eta_seconds is not None else "?"
    free_space = _format_bytes(snapshot.free_bytes) if snapshot.free_bytes is not None else "?"
    downloaded = _format_bytes(snapshot.bytes_downloaded)
    return (
        f"assets {asset_bar} {snapshot.processed_assets}/{snapshot.total_assets} | "
        f"chats {chat_bar} {snapshot.processed_chats}/{snapshot.total_chats} | "
        f"mirrored:{snapshot.mirrored_assets} reused:{snapshot.reused_assets} fail:{snapshot.failed_assets} | "
        f"data:{downloaded} | free:{free_space} | eta:{eta} | "
        f"{status}: {compact_label}"
    )


def _progress_bar(done: int, total: int, *, width: int = 10) -> str:
    if total <= 0:
        return "[" + ("-" * width) + "]"
    clamped = min(max(done, 0), total)
    filled = min(width, int((clamped / total) * width))
    if clamped == total:
        filled = width
    return "[" + ("#" * filled) + ("-" * (width - filled)) + "]"


def _format_bytes(value: int | None) -> str:
    if value is None:
        return "?"
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(max(0, int(value)))
    for unit in units:
        if size < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(size)}{unit}"
            return f"{size:.1f}{unit}"
        size /= 1024
    return f"{size:.1f}TB"


def _format_duration(seconds: float | None) -> str:
    if seconds is None:
        return "?"
    total = max(0, int(seconds))
    hours, remainder = divmod(total, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours}h{minutes:02d}m"
    if minutes:
        return f"{minutes}m{secs:02d}s"
    return f"{secs}s"


@contextmanager
def _export_interrupt_context() -> ExportStopController:
    controller = ExportStopController()
    previous_handler = signal.getsignal(signal.SIGINT)

    def handle_sigint(_signum: int, _frame: object | None) -> None:
        level = controller.request_interrupt()
        if level == 1:
            sys.stderr.write(
                "\nQuit requested. Finishing the current chat, then writing a partial index. "
                "Press Ctrl+C again to stop after the current Teams page.\n"
            )
        elif level == 2:
            sys.stderr.write(
                "\nForce-quit requested. Will stop after the current Teams page and still try to write a partial index.\n"
            )
        else:
            sys.stderr.write("\nInterrupt already pending. Waiting for a safe stop point.\n")
        sys.stderr.flush()

    signal.signal(signal.SIGINT, handle_sigint)
    try:
        yield controller
    finally:
        signal.signal(signal.SIGINT, previous_handler)


@contextmanager
def _attachment_interrupt_context() -> MirrorStopController:
    controller = MirrorStopController()
    previous_handler = signal.getsignal(signal.SIGINT)

    def handle_sigint(_signum: int, _frame: object | None) -> None:
        level = controller.request_interrupt()
        if level == 1:
            sys.stderr.write(
                "\nQuit requested. Finishing the current chat's attachments, then writing a resumable partial asset index. "
                "Press Ctrl+C again to stop after the current attachment.\n"
            )
        elif level == 2:
            sys.stderr.write(
                "\nForce-quit requested. Will stop after the current attachment and still write resumable partial mirror metadata.\n"
            )
        else:
            sys.stderr.write("\nInterrupt already pending. Waiting for a safe stop point.\n")
        sys.stderr.flush()

    signal.signal(signal.SIGINT, handle_sigint)
    try:
        yield controller
    finally:
        signal.signal(signal.SIGINT, previous_handler)
