# Outlook Classic MCP

A local Windows tool that indexes **Classic Outlook** mail into a SQLite FTS5
database and exposes it to Goose Desktop through a read-only MCP server.

- One **index worker** reads Outlook via COM and keeps the index up to date.
- One **MCP server** serves searches (and, on request, opens the original
  message in Outlook).
- No Graph credentials, no network listeners, no send/delete tools.

> **Status:** SQLite/FTS logic is unit-tested. The real Outlook COM + Goose
> integration has not been verified end-to-end; treat this as a functional
> implementation needing a Windows smoke test.

## Requirements

- Windows 10/11 with **Classic Outlook** (not New Outlook) and a configured profile.
- Python 3.11+.
- Outlook must have the mail you want indexed (Cached Exchange Mode retention,
  PST permissions, etc. limit coverage).
- Run the worker and Goose under the same signed-in account; keep the database
  on a local NTFS drive.

## Install

```powershell
cd C:\path\to\outlook-classic-mcp
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[tray]"
```

Or install directly from git into a venv:

```powershell
.\.venv\Scripts\python.exe -m pip install "outlook-classic-mcp[tray] @ git+https://github.com/koenvaneijk/outlook-classic-mcp.git"
```

This provides `outlook-index`, `outlook-search-tray`, and `outlook-open-mail`
commands.

## Quick start

1. **Smoke test** (with Outlook open) before indexing a large mailbox:

   ```powershell
   .\.venv\Scripts\python.exe -m outlook_classic_mcp.worker doctor
   .\.venv\Scripts\python.exe -m outlook_classic_mcp.worker bootstrap --folder-contains Inbox --stores default
   .\.venv\Scripts\python.exe -m outlook_classic_mcp.worker status
   ```

2. **Full mailbox bootstrap** (resumable; Ctrl+C and re-run to continue):

   ```powershell
   .\.venv\Scripts\python.exe -m outlook_classic_mcp.worker bootstrap --stores default
   ```

3. **Keep it updated** — either a manual pass, a poll loop, or the tray:

   ```powershell
   .\.venv\Scripts\python.exe -m outlook_classic_mcp.worker sync --stores default      # one pass
   .\.venv\Scripts\python.exe -m outlook_classic_mcp.worker watch --stores default    # poll every 15 min
   .\.venv\Scripts\pythonw.exe  -m outlook_classic_mcp.tray                           # tray + auto worker
   ```

   The tray auto-starts a single background worker (syncs on Outlook open, then
   every 15 min; weekly full reconciliation). Menu: **Start with Windows**,
   **Sync now**, **Quit**. Logs go to `%LOCALAPPDATA%\OutlookSQLiteMCP\sync.log`.

4. Run a periodic **full reconciliation** to catch moves/deletions that
   incremental syncs miss:

   ```powershell
   .\.venv\Scripts\python.exe -m outlook_classic_mcp.worker reconcile --stores default
   ```

Only one worker may run at a time. Keep `sync`/`watch`/`reconcile` and the MCP
server on the same database (set `OUTLOOK_INDEX_DB` consistently).

## Connect to Goose Desktop

**Extensions → Add custom extension**, type **Standard IO**:

```
Command: C:\path\to\outlook-classic-mcp\.venv\Scripts\python.exe
Args:    -m outlook_classic_mcp.mcp_server
```

(or edit `%USERPROFILE%\.config\goose\config.yaml` with `type: stdio`.) Use
`outlook_classic_mcp.mcp_server`, **not** the worker. Then ask: *"Use
search_mail to find invoices from ACME in 2022, then open the best match in
Outlook."*

## MCP tools

- `search_mail(query, mode, sender, after, before, folder, ...)` — BM25-ranked
  search. `mode` = `all` (both words), `phrase`, or `fts` (raw FTS5 grammar:
  `OR`, quoted phrases, prefixes).
- `get_mail(message_id, max_body_chars)` — full body excerpt for a result ID.
- `open_mail_in_outlook(message_id)` — open the original item in Outlook.
- `index_status()` — counts, incomplete folders, errors.
- `list_mail_folders()` — folder listing.

Dates use `received_at` → `sent_at` → `modified_at` (UTC). `after` is
inclusive, `before` exclusive. Search is keyword-based — no semantic/vector
search or attachment-body extraction.

## Notes and limits

- **Read-only mail**: the agent cannot send/delete/modify mail.
  `open_mail_in_outlook` brings a local Outlook window to the foreground.
- **Bodies at rest are plaintext** in SQLite — keep the DB in a user-private
  local directory and do not sync/share it.
- No Outlook hooks: the indexer polls, so there is lag; stale entries are
  removed only during full scans/reconciliation.
- Plain-text `MailItem.Body` only; large bodies are capped (default 65,536
  chars). Attachment filenames/contents are not indexed.

## Tests

```powershell
.\.venv\Scripts\python.exe -m pip install pytest
.\.venv\Scripts\python.exe -m pytest -q
```

Tests cover FTS, triggers, pagination, and schema with mocked metadata — no
Outlook required. Real Windows smoke testing is still needed.
