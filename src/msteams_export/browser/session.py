from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from msteams_export.browser.detection import BrowserInstall, choose_browser, discover_browsers
from msteams_export.config import detect_project_paths


DEFAULT_TEAMS_URL = "https://teams.microsoft.com"


@dataclass(slots=True)
class SessionSnapshot:
    page_url: str
    title: str
    token_count: int
    has_ic3_token: bool
    has_skype_token: bool
    has_graph_token: bool
    has_teams_ui: bool
    storage_error: str | None = None


@dataclass(slots=True)
class SessionCheckResult:
    ok: bool
    message: str
    browser_name: str | None = None
    executable_path: Path | None = None
    profile_path: Path | None = None
    teams_url: str | None = None
    snapshot: SessionSnapshot | None = None
    notes: list[str] | None = None


@dataclass(slots=True)
class SessionOpenResult:
    ok: bool
    message: str
    browser_name: str | None = None
    executable_path: Path | None = None
    profile_path: Path | None = None
    teams_url: str | None = None


def check_session(
    browser_name: str = "auto",
    profile_path: Path | None = None,
    teams_url: str = DEFAULT_TEAMS_URL,
    *,
    headless: bool = False,
    timeout_ms: int = 15_000,
) -> SessionCheckResult:
    installs = discover_browsers()
    selected = choose_browser(browser_name, installs)
    if selected is None:
        names = ", ".join(item.name for item in installs) or "none"
        return SessionCheckResult(
            ok=False,
            message=f"No supported browser installation found for '{browser_name}'. Detected: {names}.",
            notes=[
                "Install Microsoft Edge or Google Chrome, or add more browser detection candidates.",
                "On first real run we recommend a dedicated persistent profile managed by this project.",
            ],
        )

    profile = _resolve_profile_path(profile_path, selected.name)
    profile.mkdir(parents=True, exist_ok=True)

    try:
        snapshot = _run_playwright_probe(
            browser=selected,
            profile_path=profile,
            teams_url=teams_url,
            headless=headless,
            timeout_ms=timeout_ms,
        )
    except Exception as exc:
        return SessionCheckResult(
            ok=False,
            message=f"Browser launch failed: {exc}",
            browser_name=selected.name,
            executable_path=selected.executable_path,
            profile_path=profile,
            teams_url=teams_url,
            notes=[
                "The browser executable was found, but Playwright could not complete the probe.",
                "If the profile is already in use by a running browser, try a dedicated project profile.",
            ],
        )

    authenticated, notes = assess_authentication(snapshot)
    if authenticated:
        message = "Teams session looks healthy and authenticated."
    else:
        message = "Browser works, but the Teams session does not look authenticated yet."

    return SessionCheckResult(
        ok=authenticated,
        message=message,
        browser_name=selected.name,
        executable_path=selected.executable_path,
        profile_path=profile,
        teams_url=teams_url,
        snapshot=snapshot,
        notes=notes,
    )


def open_session(
    browser_name: str = "auto",
    profile_path: Path | None = None,
    teams_url: str = DEFAULT_TEAMS_URL,
    extra_urls: list[str] | None = None,
) -> SessionOpenResult:
    installs = discover_browsers()
    selected = choose_browser(browser_name, installs)
    if selected is None:
        names = ", ".join(item.name for item in installs) or "none"
        return SessionOpenResult(
            ok=False,
            message=f"No supported browser installation found for '{browser_name}'. Detected: {names}.",
        )

    profile = _resolve_profile_path(profile_path, selected.name)
    profile.mkdir(parents=True, exist_ok=True)

    try:
        _run_interactive_session(
            browser=selected,
            profile_path=profile,
            teams_url=teams_url,
            extra_urls=extra_urls,
        )
    except Exception as exc:
        return SessionOpenResult(
            ok=False,
            message=f"Could not open interactive session: {exc}",
            browser_name=selected.name,
            executable_path=selected.executable_path,
            profile_path=profile,
            teams_url=teams_url,
        )

    return SessionOpenResult(
        ok=True,
        message="Interactive Teams browser session completed.",
        browser_name=selected.name,
        executable_path=selected.executable_path,
        profile_path=profile,
        teams_url=teams_url,
    )


def assess_authentication(snapshot: SessionSnapshot) -> tuple[bool, list[str]]:
    url_lower = snapshot.page_url.lower()
    title_lower = snapshot.title.lower()
    teams_hosts = (
        "teams.microsoft.com",
        "teams.microsoft.us",
        "cloud.microsoft",
        "teams.live.com",
    )
    sign_in_signals = (
        "signin",
        "sign in",
        "login.microsoftonline",
        "login.live.com",
    )
    on_teams_host = any(host in url_lower for host in teams_hosts)
    likely_sign_in = any(signal in url_lower or signal in title_lower for signal in sign_in_signals)

    notes: list[str] = []
    if snapshot.storage_error:
        notes.append(f"Local storage inspection was limited: {snapshot.storage_error}")
    if snapshot.has_teams_ui:
        notes.append("Detected Teams-like UI elements in the loaded page.")
    if snapshot.token_count:
        notes.append(f"Found {snapshot.token_count} localStorage access token entries on the active origin.")
    if snapshot.has_ic3_token:
        notes.append("IC3 token signal present.")
    if snapshot.has_skype_token:
        notes.append("Skype auth token signal present.")
    if snapshot.has_graph_token:
        notes.append("Graph token signal present.")
    if likely_sign_in:
        notes.append("Page still looks like a sign-in or redirect flow.")
    if not on_teams_host:
        notes.append("Current page is not on a Teams host yet.")

    authenticated = on_teams_host and not likely_sign_in and (
        snapshot.has_teams_ui or snapshot.has_ic3_token or snapshot.has_skype_token
    )
    if not authenticated and on_teams_host and snapshot.token_count == 0:
        notes.append("Try running a headed session and log in once with the same persistent profile.")
    return authenticated, notes


def _resolve_profile_path(profile_path: Path | None, browser_name: str) -> Path:
    if profile_path is not None:
        return profile_path.expanduser().resolve()
    project_paths = detect_project_paths()
    return (project_paths.root / ".state" / "profiles" / browser_name).resolve()


def _run_playwright_probe(
    *,
    browser: BrowserInstall,
    profile_path: Path,
    teams_url: str,
    headless: bool,
    timeout_ms: int,
) -> SessionSnapshot:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile_path),
            executable_path=str(browser.executable_path),
            headless=headless,
            args=[
                "--no-first-run",
                "--no-default-browser-check",
            ],
        )
        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(teams_url, wait_until="domcontentloaded", timeout=timeout_ms)
            page.wait_for_timeout(2_000)
            return _snapshot_page(page)
        finally:
            context.close()


def _run_interactive_session(
    *,
    browser: BrowserInstall,
    profile_path: Path,
    teams_url: str,
    extra_urls: list[str] | None = None,
) -> None:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile_path),
            executable_path=str(browser.executable_path),
            headless=False,
            args=[
                "--no-first-run",
                "--no-default-browser-check",
            ],
        )
        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(teams_url, wait_until="domcontentloaded", timeout=30_000)
            print("Teams browser session is open.")
            print("Use the browser window to sign in if needed.")
            for extra_url in extra_urls or []:
                if not extra_url:
                    continue
                try:
                    extra_page = context.new_page()
                    extra_page.goto(extra_url, wait_until="domcontentloaded", timeout=30_000)
                    print(f"Opened extra tab for sign-in: {extra_url}")
                except Exception as exc:
                    print(f"Could not open extra tab {extra_url}: {exc}")
            print("When you are done, return here and press Enter to close the session.")
            try:
                input()
            except EOFError:
                page.wait_for_timeout(5_000)
        finally:
            context.close()


def _snapshot_page(page: Any) -> SessionSnapshot:
    page_url = page.url or ""
    try:
        title = page.title()
    except Exception:
        title = ""
    raw = page.evaluate(
        """
        () => {
          const result = {
            tokenCount: 0,
            hasIc3Token: false,
            hasSkypeToken: false,
            hasGraphToken: false,
            hasTeamsUi: false,
            storageError: null,
          };
          try {
            const keys = [];
            for (let i = 0; i < localStorage.length; i += 1) {
              const key = localStorage.key(i) || "";
              if (key.includes("accesstoken")) {
                keys.push(key);
              }
            }
            result.tokenCount = keys.length;
            result.hasIc3Token = keys.some((key) => key.includes("ic3.teams.office.com") || key.includes("chatsvcagg"));
            result.hasSkypeToken = keys.some((key) => key.includes("api.spaces.skype"));
            result.hasGraphToken = keys.some((key) => key.includes("graph.microsoft"));
          } catch (error) {
            result.storageError = String(error);
          }
          try {
            result.hasTeamsUi = Boolean(
              document.querySelector('[data-tid="app-bar-wrapper"], [data-tid="message-pane"], [data-tid="message-pane-list-viewport"]')
            );
          } catch (error) {
            if (!result.storageError) {
              result.storageError = String(error);
            }
          }
          return result;
        }
        """
    )
    return SessionSnapshot(
        page_url=page_url,
        title=title,
        token_count=int(raw.get("tokenCount", 0) or 0),
        has_ic3_token=bool(raw.get("hasIc3Token", False)),
        has_skype_token=bool(raw.get("hasSkypeToken", False)),
        has_graph_token=bool(raw.get("hasGraphToken", False)),
        has_teams_ui=bool(raw.get("hasTeamsUi", False)),
        storage_error=str(raw["storageError"]) if raw.get("storageError") else None,
    )
