from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
import subprocess
from typing import Any, Iterator

from msteams_export.browser.detection import choose_browser, discover_browsers
from msteams_export.browser.session import DEFAULT_TEAMS_URL
from msteams_export.config import detect_project_paths
from msteams_export.polite_mode import DEFAULT_POLITE_MODE, build_browser_polite_mode_payload


@dataclass(slots=True)
class TeamsBrowserRequest:
    browser_name: str = "auto"
    profile_path: Path | None = None
    teams_url: str = DEFAULT_TEAMS_URL
    headless: bool = True
    timeout_ms: int = 30_000
    polite_mode: bool = True


@dataclass(slots=True)
class TeamsBrowserTarget:
    browser_name: str
    executable_path: Path
    profile_path: Path
    teams_url: str
    headless: bool
    timeout_ms: int
    polite_mode: bool


def resolve_browser_target(request: TeamsBrowserRequest) -> TeamsBrowserTarget:
    installs = discover_browsers()
    selected = choose_browser(request.browser_name, installs)
    if selected is None:
        names = ", ".join(item.name for item in installs) or "none"
        raise RuntimeError(
            f"No supported browser installation found for '{request.browser_name}'. Detected: {names}."
        )

    profile_path = _resolve_profile_path(request.profile_path, selected.name)
    profile_path.mkdir(parents=True, exist_ok=True)
    return TeamsBrowserTarget(
        browser_name=selected.name,
        executable_path=selected.executable_path,
        profile_path=profile_path,
        teams_url=request.teams_url,
        headless=request.headless,
        timeout_ms=request.timeout_ms,
        polite_mode=request.polite_mode,
    )


@contextmanager
def open_teams_page(target: TeamsBrowserTarget) -> Iterator[Any]:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        context = _launch_context_with_fallback(playwright, target)
        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(target.teams_url, wait_until="domcontentloaded", timeout=target.timeout_ms)
            page.wait_for_timeout(2_500)
            page.set_default_timeout(target.timeout_ms)
            yield page
        finally:
            context.close()


def run_api_action(page: Any, mode: str, **payload: Any) -> dict[str, Any]:
    return page.evaluate(
        _build_api_script(),
        {
            "mode": mode,
            "politeMode": build_browser_polite_mode_payload(DEFAULT_POLITE_MODE),
            **payload,
        },
    )


def _resolve_profile_path(profile_path: Path | None, browser_name: str) -> Path:
    if profile_path is not None:
        return profile_path.expanduser().resolve()
    project_paths = detect_project_paths()
    return (project_paths.root / ".state" / "profiles" / browser_name).resolve()


def _build_api_script() -> str:
    return r"""
    async (input) => {
      const settings = input || {};
      const politeMode = {
        enabled: true,
        requestSpacingMs: 650,
        pageSpacingMs: 450,
        conversationSpacingMs: 1200,
        retryLimit: 8,
        retryBaseMs: 2500,
        retryMaxMs: 60000,
        jitterMs: 350,
        retryStatuses: [429, 503],
        ...(settings.politeMode || {}),
      };
      let lastFetchAt = 0;

      const findValidToken = (scopePattern) => {
        const now = Math.floor(Date.now() / 1000);
        for (let i = 0; i < localStorage.length; i += 1) {
          const key = localStorage.key(i);
          if (!key || !key.includes("accesstoken") || !key.includes(scopePattern)) continue;
          try {
            const entry = JSON.parse(localStorage.getItem(key) || "");
            if (Number(entry.expiresOn) > now && entry.secret) {
              return entry.secret;
            }
          } catch (_) {
          }
        }
        return null;
      };

      const getIc3Token = () =>
        findValidToken("ic3.teams.office.com")
        || findValidToken("ic3.teams.office365.us")
        || findValidToken("chatsvcagg");

      const getSkypeToken = () => findValidToken("api.spaces.skype");
      const getGraphToken = () => findValidToken("graph.microsoft.com");

      const detectChatContext = () => {
        const navSelected = Boolean(
          document.querySelector('[data-tid="app-bar-wrapper"] button[aria-pressed="true"][aria-label^="Chat" i]')
        );
        const hasSurface = Boolean(
          document.querySelector('[data-tid="message-pane-list-viewport"], [data-tid="chat-message-list"], [data-tid="chat-pane"]')
        );
        if (navSelected && hasSurface) {
          return { ok: true, reason: null };
        }
        if (!navSelected) {
          return { ok: false, reason: "Switch to the Chat app in Teams before exporting." };
        }
        return { ok: false, reason: "Open a chat conversation before exporting." };
      };

      const stripTitleSuffix = (value) =>
        (value || "").replace(/\s*\|\s*Microsoft Teams\s*$/i, "").trim();

      const extractTitle = () => {
        const selectors = [
          '[data-tid="chat-header-title"]',
          '[data-tid="title"]',
          '[data-tid="app-layout-area--header"] h1',
          '[role="heading"]',
        ];
        for (const selector of selectors) {
          const element = document.querySelector(selector);
          const text = (element?.textContent || "").trim();
          if (text) {
            return text;
          }
        }
        const title = stripTitleSuffix(document.title);
        return title || "Teams Export";
      };

      const lookupConversationInIdb = async (mids) => {
        if (typeof indexedDB.databases !== "function") return null;
        const databases = await indexedDB.databases();
        const database = databases.find((item) => item.name && item.name.includes("replychain-manager:react-web-client"));
        if (!database?.name) return null;

        return await new Promise((resolve) => {
          const request = indexedDB.open(database.name);
          request.onerror = () => resolve(null);
          request.onsuccess = () => {
            const db = request.result;
            try {
              const tx = db.transaction("replychains", "readonly");
              const store = tx.objectStore("replychains");
              const getAll = store.getAll();
              getAll.onerror = () => {
                db.close();
                resolve(null);
              };
              getAll.onsuccess = () => {
                const midSet = new Set(mids);
                const records = getAll.result || [];
                for (const record of records) {
                  if (midSet.has(record?.replyChainId)) {
                    db.close();
                    resolve(record.conversationId || null);
                    return;
                  }
                }
                for (const record of records) {
                  const map = record?.messageMap || {};
                  for (const key of Object.keys(map)) {
                    for (const mid of mids) {
                      if (key.includes(mid)) {
                        db.close();
                        resolve(record.conversationId || null);
                        return;
                      }
                    }
                  }
                }
                db.close();
                resolve(null);
              };
            } catch (_) {
              db.close();
              resolve(null);
            }
          };
        });
      };

      const extractConversationId = async () => {
        try {
          const mids = Array.from(document.querySelectorAll("[data-mid]"))
            .map((element) => element.getAttribute("data-mid"))
            .filter(Boolean);
          if (mids.length) {
            const fromIdb = await lookupConversationInIdb(mids);
            if (fromIdb) return fromIdb;
          }
        } catch (_) {
        }

        const convRegex = /(?:19:[^|"}\s&?/]+@(?:unq\.gbl\.spaces|thread\.[a-z0-9]+)|48:notes)/i;
        try {
          const currentUrl = new URL(window.location.href);
          for (const [, value] of currentUrl.searchParams) {
            if (convRegex.test(value)) {
              const match = value.match(convRegex);
              if (match) return match[0];
            }
          }
          const hashMatch = window.location.hash.match(convRegex);
          if (hashMatch) return decodeURIComponent(hashMatch[0]);
        } catch (_) {
        }

        const chatPane = document.querySelector('[data-tid="message-pane"]');
        return (
          chatPane?.getAttribute("data-convid")
          || chatPane?.getAttribute("data-tid-convid")
          || chatPane?.getAttribute("data-conversation-id")
          || null
        );
      };

      const getAuthzUrl = () => {
        const host = location.hostname.toLowerCase();
        if (host.includes("teams.microsoft.us")) {
          return "https://authsvc.teams.microsoft.us/v1.0/authz";
        }
        return "https://authsvc.teams.microsoft.com/v1.0/authz";
      };

      const discover = async (skypeToken) => {
        const response = await fetch(getAuthzUrl(), {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${skypeToken}`,
            "Content-Type": "application/json",
          },
          body: "{}",
        });
        if (!response.ok) {
          throw new Error(`authz failed: ${response.status} ${response.statusText}`);
        }
        const data = await response.json();
        const chatServiceUrl = data?.regionGtms?.chatService;
        const userRegion = data?.userRegion || data?.region || "";
        if (!chatServiceUrl) {
          throw new Error("authz response missing chatService URL");
        }
        return { chatServiceUrl, userRegion };
      };

      const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

      const maybeSpaceRequests = async (spacingMs) => {
        if (!politeMode.enabled) return;
        const now = Date.now();
        const elapsedMs = now - lastFetchAt;
        const waitMs = Math.max(0, Number(spacingMs || 0) - elapsedMs);
        if (waitMs > 0) {
          await sleep(waitMs);
        }
      };

      const fetchPageWithRetry = async (url, token) => {
        const retryStatuses = new Set(Array.isArray(politeMode.retryStatuses) ? politeMode.retryStatuses : [429, 503]);
        const retryLimit = Math.max(0, Number(politeMode.retryLimit || 0));
        for (let attempt = 0; attempt <= retryLimit; attempt += 1) {
          await maybeSpaceRequests(politeMode.requestSpacingMs);
          lastFetchAt = Date.now();
          // Re-check token freshness before each request
          const freshToken = getIc3Token() || token;
          const response = await fetch(url, {
            headers: { "Authorization": `Bearer ${freshToken}` },
          });
          if (retryStatuses.has(response.status)) {
            if (attempt >= retryLimit) {
              throw new Error("Rate limited after max retries");
            }
            const retryAfter = response.headers.get("Retry-After");
            const waitMs = parseRetryAfter(retryAfter, attempt);
            await sleep(waitMs);
            continue;
          }
          if (response.status === 401 || response.status === 403) {
            // Token may have expired — try refreshing once
            const refreshedToken = getIc3Token();
            if (refreshedToken && refreshedToken !== freshToken && attempt === 0) {
              continue;  // Retry with refreshed token
            }
            throw new Error(`Teams API auth error: ${response.status} ${response.statusText}. Token may have expired — try re-opening session.`);
          }
          if (!response.ok) {
            throw new Error(`Teams API error: ${response.status} ${response.statusText}`);
          }
          return await response.json();
        }
        throw new Error("Unreachable");
      };

      const parseRetryAfter = (header, attempt) => {
        if (header) {
          const seconds = parseInt(header, 10);
          if (!isNaN(seconds)) {
            return Math.min(seconds * 1000, Number(politeMode.retryMaxMs || 60000));
          }
          const date = new Date(header);
          if (!isNaN(date.getTime())) {
            return Math.max(0, Math.min(date.getTime() - Date.now(), Number(politeMode.retryMaxMs || 60000)));
          }
        }
        const exponentialMs = Math.min(
          Number(politeMode.retryBaseMs || 2500) * (2 ** attempt),
          Number(politeMode.retryMaxMs || 60000),
        );
        const jitterMs = Math.floor(Math.random() * Math.max(0, Number(politeMode.jitterMs || 0)));
        return Math.min(exponentialMs + jitterMs, Number(politeMode.retryMaxMs || 60000));
      };

      const fetchAllMessages = async (chatServiceUrl, conversationId, ic3Token) => {
        const allMessages = [];
        const encodedConversationId = encodeURIComponent(conversationId);
        let nextUrl = `${chatServiceUrl}/v1/users/ME/conversations/${encodedConversationId}/messages?pageSize=200&startTime=1&view=msnp24Equivalent%7CsupportsMessageProperties`;
        let page = 0;
        let delayMs = 150;
        let warning = null;

        while (nextUrl && page < 500) {
          page += 1;
          const token = getIc3Token() || ic3Token;
          if (!token) {
            warning = `Token expired during export at page ${page}. Partial data returned.`;
            break;
          }
          let data;
          try {
            data = await fetchPageWithRetry(nextUrl, token);
          } catch (fetchError) {
            if (allMessages.length > 0) {
              // We already have partial data — return it instead of losing everything
              warning = `Stopped at page ${page}: ${fetchError.message}. Returning ${allMessages.length} messages collected so far.`;
              break;
            }
            throw fetchError;  // No data yet — propagate error
          }
          if (data?.errorCode) {
            if (allMessages.length > 0) {
              warning = `API error at page ${page}: ${data.errorCode}. Returning ${allMessages.length} messages collected so far.`;
              break;
            }
            throw new Error(`Messages API error ${data.errorCode}: ${data.message}`);
          }
          const messages = Array.isArray(data?.messages) ? data.messages : [];
          if (!messages.length && !data?._metadata?.backwardLink) {
            break;
          }
          allMessages.push(...messages);
          nextUrl = data?._metadata?.backwardLink || null;
          if (nextUrl) {
            if (politeMode.enabled) {
              await sleep(Math.max(delayMs, Number(politeMode.pageSpacingMs || delayMs)));
            } else {
              await sleep(delayMs);
            }
            if (page > 20) {
              delayMs = Math.min(delayMs + 75, 900);
            }
          }
        }

        return {
          rawMessages: allMessages,
          rawCount: allMessages.length,
          pageCount: page,
          warning: warning,
        };
      };

      const fetchAllConversations = async (chatServiceUrl, ic3Token) => {
        const allConversations = [];
        let nextUrl = `${chatServiceUrl}/v1/users/ME/conversations`;
        let page = 0;
        let delayMs = 120;

        while (nextUrl && page < 200) {
          page += 1;
          const token = getIc3Token() || ic3Token;
          if (!token) {
            if (allConversations.length > 0) break;
            throw new Error("Token expired before fetching conversations.");
          }
          const data = await fetchPageWithRetry(nextUrl, token);
          if (data?.errorCode) {
            if (allConversations.length > 0) break;
            throw new Error(`Conversations API error ${data.errorCode}: ${data.message}`);
          }
          const conversations = Array.isArray(data?.conversations) ? data.conversations : [];
          if (!conversations.length && !data?._metadata?.backwardLink) {
            break;
          }
          allConversations.push(...conversations);
          nextUrl = data?._metadata?.backwardLink || null;
          if (nextUrl) {
            if (politeMode.enabled) {
              await sleep(Math.max(delayMs, Number(politeMode.pageSpacingMs || delayMs)));
            } else {
              await sleep(delayMs);
            }
            if (page > 20) {
              delayMs = Math.min(delayMs + 50, 800);
            }
          }
        }

        return {
          rawConversations: allConversations,
          rawCount: allConversations.length,
          pageCount: page,
        };
      };

      const readConversationCache = async () => {
        if (typeof indexedDB.databases !== "function") {
          return {
            rawCachedConversations: [],
            cacheCount: 0,
            cacheHiddenCount: 0,
          };
        }
        const databases = await indexedDB.databases();
        const database = databases.find((item) => item.name && item.name.includes("Teams:conversation-manager:react-web-client"));
        if (!database?.name) {
          return {
            rawCachedConversations: [],
            cacheCount: 0,
            cacheHiddenCount: 0,
          };
        }

        const rawCachedConversations = await new Promise((resolve) => {
          const request = indexedDB.open(database.name);
          request.onerror = () => resolve([]);
          request.onsuccess = () => {
            const db = request.result;
            try {
              const tx = db.transaction("conversations", "readonly");
              const store = tx.objectStore("conversations");
              const getAll = store.getAll();
              getAll.onerror = () => {
                db.close();
                resolve([]);
              };
              getAll.onsuccess = () => {
                const records = Array.isArray(getAll.result) ? getAll.result : [];
                db.close();
                resolve(records);
              };
            } catch (_) {
              db.close();
              resolve([]);
            }
          };
        });
        const cacheHiddenCount = rawCachedConversations.filter(
          (item) => String(item?.threadProperties?.hidden || "").toLowerCase() === "true" || item?.threadProperties?.hidden === true
        ).length;
        return {
          rawCachedConversations,
          cacheCount: rawCachedConversations.length,
          cacheHiddenCount,
        };
      };

      const readReplychainCache = async (conversationId) => {
        if (typeof indexedDB.databases !== "function") {
          return { rawCachedMessages: [], cacheCount: 0, replychainCount: 0 };
        }
        const databases = await indexedDB.databases();
        const database = databases.find((item) => item.name && item.name.includes("replychain-manager:react-web-client"));
        if (!database?.name) {
          return { rawCachedMessages: [], cacheCount: 0, replychainCount: 0 };
        }

        const records = await new Promise((resolve) => {
          const request = indexedDB.open(database.name);
          request.onerror = () => resolve([]);
          request.onsuccess = () => {
            const db = request.result;
            try {
              const tx = db.transaction("replychains", "readonly");
              const store = tx.objectStore("replychains");
              const getAll = store.getAll();
              getAll.onerror = () => {
                db.close();
                resolve([]);
              };
              getAll.onsuccess = () => {
                const result = Array.isArray(getAll.result) ? getAll.result : [];
                db.close();
                resolve(result);
              };
            } catch (_) {
              db.close();
              resolve([]);
            }
          };
        });

        const matches = records.filter((item) => item?.conversationId === conversationId);
        const flattened = [];
        for (const record of matches) {
          const map = record?.messageMap || {};
          for (const value of Object.values(map)) {
            if (value && typeof value === "object") {
              flattened.push(value);
            }
          }
        }
        return {
          rawCachedMessages: flattened,
          cacheCount: flattened.length,
          replychainCount: matches.length,
        };
      };

      const skypeToken = getSkypeToken();
      if (!skypeToken) {
        return { ok: false, error: "No valid Skype token found in Teams session." };
      }
      const ic3Token = getIc3Token();
      if (!ic3Token) {
        return { ok: false, error: "No valid IC3 token found in Teams session." };
      }
      const { chatServiceUrl, userRegion } = await discover(skypeToken);

      if (settings.mode === "conversation-list") {
        const [result, cacheResult] = await Promise.all([
          fetchAllConversations(chatServiceUrl, ic3Token),
          readConversationCache(),
        ]);
        return {
          ok: true,
          userRegion,
          ...result,
          ...cacheResult,
        };
      }

      if (settings.mode === "cached-conversation-messages") {
        const conversationId = settings.conversationId || null;
        if (!conversationId) {
          return { ok: false, error: "conversationId is required for cached-conversation-messages mode." };
        }
        const cacheResult = await readReplychainCache(conversationId);
        return {
          ok: true,
          conversationId,
          userRegion,
          ...cacheResult,
        };
      }

      if (settings.mode === "conversation-messages") {
        const conversationId = settings.conversationId || null;
        if (!conversationId) {
          return { ok: false, error: "conversationId is required for conversation-messages mode." };
        }
        const result = await fetchAllMessages(chatServiceUrl, conversationId, ic3Token);
        return {
          ok: true,
          userRegion,
          conversationId,
          title: settings.title || conversationId,
          ...result,
        };
      }

      if (settings.mode === "conversation-messages-page") {
        const conversationId = settings.conversationId || null;
        if (!conversationId) {
          return { ok: false, error: "conversationId is required for conversation-messages-page mode." };
        }
        const encodedConversationId = encodeURIComponent(conversationId);
        const nextUrl = settings.nextUrl || `${chatServiceUrl}/v1/users/ME/conversations/${encodedConversationId}/messages?pageSize=200&startTime=1&view=msnp24Equivalent%7CsupportsMessageProperties`;
        const token = getIc3Token() || ic3Token;
        const data = await fetchPageWithRetry(nextUrl, token);
        if (data?.errorCode) {
          throw new Error(`Messages API error ${data.errorCode}: ${data.message}`);
        }
        const messages = Array.isArray(data?.messages) ? data.messages : [];
        return {
          ok: true,
          userRegion,
          conversationId,
          title: settings.title || conversationId,
          rawMessages: messages,
          rawCount: messages.length,
          nextUrl: data?._metadata?.backwardLink || null,
          pageCount: 1,
        };
      }

      if (settings.mode === "thread-messages-page") {
        const threadUrl = settings.threadUrl || null;
        if (!threadUrl) {
          return { ok: false, error: "threadUrl is required for thread-messages-page mode." };
        }
        const nextUrl = settings.nextUrl || `${threadUrl}/messages?pageSize=200&startTime=1&view=msnp24Equivalent%7CsupportsMessageProperties`;
        const token = getIc3Token() || ic3Token;
        const data = await fetchPageWithRetry(nextUrl, token);
        if (data?.errorCode) {
          throw new Error(`Messages API error ${data.errorCode}: ${data.message}`);
        }
        const messages = Array.isArray(data?.messages) ? data.messages : [];
        return {
          ok: true,
          userRegion,
          threadUrl,
          rawMessages: messages,
          rawCount: messages.length,
          nextUrl: data?._metadata?.backwardLink || null,
          pageCount: 1,
        };
      }

      if (settings.mode === "thread-resource") {
        const threadUrl = settings.threadUrl || null;
        if (!threadUrl) {
          return { ok: false, error: "threadUrl is required for thread-resource mode." };
        }
        const token = getIc3Token() || ic3Token;
        const data = await fetchPageWithRetry(threadUrl, token);
        return {
          ok: true,
          userRegion,
          threadUrl,
          resource: data,
        };
      }

      if (settings.mode === "graph-team-channels-page") {
        const teamId = settings.teamId || null;
        if (!teamId) {
          return { ok: false, error: "teamId is required for graph-team-channels-page mode." };
        }
        const graphToken = getGraphToken();
        if (!graphToken) {
          return { ok: false, error: "No valid Graph token found in Teams session." };
        }
        const nextUrl = settings.nextUrl || `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels?$top=200`;
        const data = await fetchPageWithRetry(nextUrl, graphToken);
        const channels = Array.isArray(data?.value) ? data.value : [];
        return {
          ok: true,
          teamId,
          userRegion,
          rawChannels: channels,
          rawCount: channels.length,
          nextUrl: data?.["@odata.nextLink"] || null,
          pageCount: 1,
        };
      }

      if (settings.mode === "graph-channel-messages-page") {
        const teamId = settings.teamId || null;
        const channelId = settings.channelId || null;
        if (!teamId || !channelId) {
          return { ok: false, error: "teamId and channelId are required for graph-channel-messages-page mode." };
        }
        const graphToken = getGraphToken();
        if (!graphToken) {
          return { ok: false, error: "No valid Graph token found in Teams session." };
        }
        const nextUrl = settings.nextUrl || `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages?$top=50`;
        const data = await fetchPageWithRetry(nextUrl, graphToken);
        const messages = Array.isArray(data?.value) ? data.value : [];
        return {
          ok: true,
          teamId,
          channelId,
          userRegion,
          rawMessages: messages,
          rawCount: messages.length,
          nextUrl: data?.["@odata.nextLink"] || null,
          pageCount: 1,
        };
      }

      if (settings.mode === "graph-channel-message-replies-page") {
        const teamId = settings.teamId || null;
        const channelId = settings.channelId || null;
        const messageId = settings.messageId || null;
        if (!teamId || !channelId || !messageId) {
          return { ok: false, error: "teamId, channelId and messageId are required for graph-channel-message-replies-page mode." };
        }
        const graphToken = getGraphToken();
        if (!graphToken) {
          return { ok: false, error: "No valid Graph token found in Teams session." };
        }
        const nextUrl = settings.nextUrl || `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(messageId)}/replies?$top=50`;
        const data = await fetchPageWithRetry(nextUrl, graphToken);
        const replies = Array.isArray(data?.value) ? data.value : [];
        return {
          ok: true,
          teamId,
          channelId,
          messageId,
          userRegion,
          rawReplies: replies,
          rawCount: replies.length,
          nextUrl: data?.["@odata.nextLink"] || null,
          pageCount: 1,
        };
      }

      if (settings.mode === "active-chat") {
        const context = detectChatContext();
        if (!context.ok) {
          return { ok: false, error: context.reason };
        }
        const conversationId = await extractConversationId();
        if (!conversationId) {
          return { ok: false, error: "Could not detect the active chat conversation. Open a chat first." };
        }
        const result = await fetchAllMessages(chatServiceUrl, conversationId, ic3Token);
        return {
          ok: true,
          title: extractTitle(),
          conversationId,
          userRegion,
          ...result,
        };
      }

      return { ok: false, error: `Unsupported Teams API action: ${settings.mode || "unknown"}` };
    }
    """


def _launch_context_with_fallback(playwright: Any, target: TeamsBrowserTarget) -> Any:
    try:
        return _launch_persistent_context(playwright, target=target, headless=target.headless)
    except Exception as exc:
        launch_note = _maybe_recover_profile_lock(target, exc)
        if launch_note is not None:
            try:
                return _launch_persistent_context(playwright, target=target, headless=target.headless)
            except Exception as retry_exc:
                exc = retry_exc
        if not target.headless:
            raise RuntimeError(_format_launch_failure(target, exc, launch_note=launch_note)) from exc
        try:
            return _launch_persistent_context(playwright, target=target, headless=False)
        except Exception as headed_exc:
            raise RuntimeError(
                _format_launch_failure(target, exc, headed_exc=headed_exc, launch_note=launch_note)
            ) from headed_exc


def _launch_persistent_context(playwright: Any, *, target: TeamsBrowserTarget, headless: bool) -> Any:
    return playwright.chromium.launch_persistent_context(
        user_data_dir=str(target.profile_path),
        executable_path=str(target.executable_path),
        headless=headless,
        args=[
            "--no-first-run",
            "--no-default-browser-check",
        ],
    )


def _format_launch_failure(
    target: TeamsBrowserTarget,
    exc: Exception,
    *,
    headed_exc: Exception | None = None,
    launch_note: str | None = None,
) -> str:
    details = [
        f"Could not launch {target.browser_name} with persistent profile {target.profile_path}.",
    ]
    if launch_note:
        details.append(launch_note)
    if target.headless:
        details.append(f"Headless launch failed: {exc}")
        if headed_exc is not None:
            details.append(f"Visible-window retry also failed: {headed_exc}")
        else:
            details.append("A visible browser retry may work better in managed Windows Edge environments.")
    else:
        details.append(f"Launch failed: {exc}")
    details.append(
        "Close any running browser that uses the same profile, or try session-open first, "
        "then rerun export with the same profile. If Edge still exits immediately, try browser=chrome."
    )
    return " ".join(str(part) for part in details if part)


def _maybe_recover_profile_lock(target: TeamsBrowserTarget, exc: Exception) -> str | None:
    if not _looks_like_profile_lock_error(exc):
        return None

    lock_paths = [path for path in _profile_lock_paths(target.profile_path) if path.exists() or path.is_symlink()]
    if not lock_paths:
        return (
            "Chromium reported the profile as already locked, but no Singleton lock files were found in the profile "
            "directory."
        )

    if _profile_path_appears_in_use(target.profile_path):
        return (
            "The persistent browser profile appears to be in use already. Close the other TeamsExport/Chromium window "
            "that uses this same profile before retrying."
        )

    removed_any = False
    for path in lock_paths:
        try:
            if path.is_dir() and not path.is_symlink():
                continue
            path.unlink()
            removed_any = True
        except OSError:
            continue

    if removed_any:
        return "A stale Chromium Singleton lock was found in the profile and was removed automatically before retrying."
    return (
        "A stale Chromium Singleton lock may be present in the profile directory, but it could not be removed "
        "automatically."
    )


def _looks_like_profile_lock_error(exc: Exception) -> bool:
    text = str(exc)
    return "ProcessSingleton" in text or "SingletonLock" in text or "profile directory is already in use" in text


def _profile_lock_paths(profile_path: Path) -> tuple[Path, ...]:
    return (
        profile_path / "SingletonLock",
        profile_path / "SingletonCookie",
        profile_path / "SingletonSocket",
    )


def _profile_path_appears_in_use(profile_path: Path) -> bool:
    try:
        result = subprocess.run(
            ["ps", "-axo", "pid=,command="],
            check=True,
            capture_output=True,
            text=True,
        )
    except Exception:
        return True

    profile_text = str(profile_path)
    for line in result.stdout.splitlines():
        command = line.strip()
        if not command or profile_text not in command:
            continue
        if "ps -axo" in command:
            continue
        return True
    return False
