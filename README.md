# Outlook Classic MCP + Windows Tray Search (v0.3: stable message handles)

## Upgrade an existing 1M+ message index (no rebuild)

**This version reuses your existing SQLite index and upgrades its schema in-place
from v1 to v2.** It does **not** rebuild the full-text index or re-download
messages. Back up the database first while the tray and worker are stopped:

```powershell
$source = "$env:LOCALAPPDATA\OutlookSQLiteMCP\mail.sqlite3"
Copy-Item $source "$source.backup"   # only when worker AND Goose MCP are closed
```

Unzip this
release **over the same `outlook-classic-mcp` project directory** that Goose
already uses (overwrite `.py`/`README.md`/`pyproject.toml` files; keep `.venv`).
Your actual database normally lives at
`%LOCALAPPDATA%\OutlookSQLiteMCP\mail.sqlite3` and is never deleted.

**1. Stop your old `worker watch` console first** (`Ctrl+C`). Do not run two
indexing workers at once.

**2. In PowerShell in the updated project directory:**

```powershell
.\.venv\Scripts\python.exe -m pip install -e ".[tray]"
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker status
.\.venv\Scripts\pythonw.exe -m outlook_classic_mcp.tray
```

If you installed the update into a *new* directory, create a new `.venv` there
and update Goose's extension command path accordingly; the default SQLite DB
path is unchanged. To test Tk support, run `python -m tkinter`. The tray UI needs
Python's **Tk/Tcl** component, included with the usual Windows Python installer
but not necessarily with Python's embedded/standalone distributions.

**3. Use it:** Find the envelope icon in the Windows notification area (possibly
behind the `^` overflow button), click **Search emails**. Search the FTS5 index
and double-click a result to show the *original* Outlook Classic email. Search
results include subject/sender/folder/date and a safe plain-text preview.

For a first launch that shows errors in PowerShell, use
`.\.venv\Scripts\python.exe -m outlook_classic_mcp.tray --search`, then choose
**Quit** from the tray icon and relaunch with `pythonw.exe` for a
console-free session.

The tray automatically runs a hidden, single-instance index worker:

* Checks for Outlook Classic every **5 seconds** without launching it.
* Syncs immediately when Outlook opens, then every **15 minutes** while open.
* Pauses when Outlook closes; the tray/worker remain running until you quit.
* Checks for **weekly full reconciliation** to catch moves/deletions. Only folders
  due for reconciliation are scanned; interrupted folders resume from their
  existing checkpoints. A million-message reconciliation can be resource-heavy.
* **Sync now** requests the next pass (if Outlook is closed, the request waits).
* Writes diagnostic logs to `%LOCALAPPDATA%\OutlookSQLiteMCP\sync.log`.
* Tray menu **Start with Windows** creates/removes a per-user Startup shortcut;
  no admin account or background Windows service is required. It runs only
  while you are signed in, **not** when Windows is shut down or logged out.
* Tray menu **Quit** stops the worker. Closing the search window only hides it.

If the tray says the worker could not start, check `sync.log`. If the log reports
"Another index worker is already running", stop your original manual `watch`
process; then restart the tray app. The tray deliberately doesn't take over
another process's lock. It is still possible to search while the worker is stopped.

### Open messages from Goose Desktop

The same **Standard IO extension command** still works:

```text
C:\path\to\outlook-classic-mcp\.venv\Scripts\python.exe -m outlook_classic_mcp.mcp_server
```

Restart/re-enable the Goose extension after updating. Now you can ask:

> Find the email from Alice discussing the contract and **open the top result in Outlook Classic**.

Goose calls `search_mail(...)` followed by the new
`open_mail_in_outlook(message_id=123)` tool, using the real result ID (not a
made-up number). This opens a visible Outlook message window. It can even launch
Outlook when closed **because this is an explicit user-requested opening**, unlike
background synchronization. You can also call the helper manually:

```powershell
.\.venv\Scripts\python.exe -m outlook_classic_mcp.open_message 123
```

`123` is the SQLite `id` returned by `search_mail`, not an Outlook EntryID.
The opener first uses **EntryID + StoreID**. If the item moved, it tries:

1. Already indexed candidates with matching **PR_SEARCH_KEY** or Internet **Message-ID**.
2. Restricted per-folder MAPI table lookups (Message-ID, then exact subject).
3. For pre-upgrade rows without keys: an exact **subject + sender + timestamp** match.

Every candidate is verified before opening. **Copies can share identifiers**;
ambiguous matches fail closed rather than opening the wrong mail. Successfully
located items get their current EntryID cached in the existing SQLite row when
possible. Old numeric IDs survive subsequent reconciliation through lightweight
`message_handles` tombstones. Deleting a message permanently means it cannot
be opened, even though its historical handle remains.

This upgrade captures the two strong identifiers on newly indexed/changed mail.
To populate identifiers for *existing* mail without re-reading any message
bodies, run this **optional, resumable** process with Outlook open, **after
quitting the tray** (the command uses the exclusive worker lock):

```powershell
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker backfill-ids --backfill-limit 10000
```

Repeat as desired; the cursor persists between runs. For 1M+ messages this may
take a long time, but it only fetches MAPI identifiers rather than text or
attachments. The existing index and search work even without this backfill,
with a more conservative metadata-based fallback for old entries.

This is not a globally immutable Outlook identifier. Moves across stores,
provider limitations, duplicate/copy messages, deleted mail, and unsupported
Outlook MAPI filters can still prevent a match. No mail is sent, deleted or
modified by the opener.

If you set a custom `OUTLOOK_INDEX_DB`, pass `--db "C:\...\mail.sqlite3"` to
`outlook_classic_mcp.tray` as well, to ensure all three components use one index.

---

## Original setup and detailed engineering documentation

# Outlook Classic → SQLite FTS5 → MCP for Goose

A **local Windows project**: one foreground/background-ish **index worker** reads Outlook Classic via COM; one MCP server serves searches from SQLite and, on explicit request,
opens the original message in Outlook Classic. No Microsoft Graph credentials, remote APIs, network listeners, send/delete tools, or automatic dependency downloads after installation. The MCP process doesn't import `win32com`; the opener uses a separate subprocess.

> **Status:** Tested against SQLite with fake mailbox data on Linux. The real Outlook COM/Windows and Goose Desktop integration **has not been executed here**; treat this as a functional implementation requiring a Windows smoke test, not a verified production release.

## Requirements

* Windows 10/11, **Classic Outlook** with a configured profile (not New Outlook).
* Python **3.11+** and free disk space. **1M+ emails can require many gigabytes/tens of GB**, depending on body sizes, and initial export may take hours or longer.
* Outlook must have downloaded/access to the mail you wish to index. Cached Exchange Mode retention, disconnected PSTs, permissions, and Exchange online-only items can limit coverage. This index cannot magically index mail Outlook's COM API cannot access.
* Run worker and Goose under the **same signed-in Windows account**. Keep the database on a **local NTFS drive**, not a synced/network folder.

## 1. Install (PowerShell)

```powershell
cd C:\path\to\outlook-classic-mcp
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -e .
```

Alternatively, install straight from this GitHub repository without cloning:

```powershell
.\.venv\Scripts\python.exe -m pip install "outlook-classic-mcp[tray] @ git+https://github.com/koenvaneijk/outlook-classic-mcp.git"
```

That installs the package into the venv and provides the `outlook-index`,
`outlook-search-tray`, and `outlook-open-mail` commands. For development or
frequent updates, use the clone + `pip install -e .` above instead. To use
the tray UI, make sure the venv (or its Scripts directory) is on your PATH,
or launch it as `python -m outlook_classic_mcp.tray`.

Use a reviewed, trusted venv. Optional: pin transitive dependencies with `pip freeze > installed-versions.txt`.

## 2. Small smoke test **before** indexing a million emails

With Outlook Classic running, **first verify COM enumeration**, then index a small folder:

```powershell
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker doctor
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker bootstrap --folder-contains Inbox --stores default --batch-size 100
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker status
```

**Check the output and whether searches work.** `--folder-contains Inbox` is for the smoke test only, not a complete mailbox scan. Results depend on how your Outlook folders are named (e.g. localized folder names). You can use `--folder-contains` with a nested folder path to reduce test size further.

## 3. One-time full mailbox bootstrap (resumable)

```powershell
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker bootstrap --stores default
```

* `--stores default` = only the default Outlook store. Use `--stores all` to include other accessible stores/PSTs/shared mailboxes; verify which stores you're allowed to index.
* Each folder is committed in batches (250 metadata rows by default), and a **last-modification timestamp checkpoint** is saved. If interrupted, repeat the same `bootstrap` command. Completed folders are skipped. The newest-first sort makes resumes efficient in terms of *full-body fetches*; revisiting earlier metadata has a cost.
* Large bodies are capped at **65,536 characters** to constrain disk usage and COM-to-Python transfer. Change with `--max-body-chars N` (up to 1M). No attachment *contents* are indexed.
* After a successful full scan, stale messages no longer in a folder are deleted from SQLite. To prevent accidental mass deletion when Outlook returns unexpectedly few rows, pruning >25% of a folder with 500+ indexed emails requires `--allow-large-prune` after you've verified it is intentional. Moves can create new EntryIDs, so index record IDs are not permanent identifiers.
* Start with the full set of stores/folders you intend to keep using. A full scan doesn't discover messages in inaccessible stores.
* `--scan-all-metadata` disables Outlook date filtering. For the **full** scan, this is already the default internally for safe resumes (see Engineering notes below). Use that flag for region/locale-sensitive **incremental** runs too if needed.

You can **Ctrl+C** and resume. One failed folder is logged and the next folder continues; inspect status/errors afterwards.

## 4. Incremental updates while Goose runs

```powershell
# One sync pass, just recently modified items (default 30-day window + 2-day overlap)
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker sync --stores default

# Or continuously poll Outlook every 15 minutes (separate terminal/process)
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker watch --stores default --every-minutes 15
```

The worker and MCP can run concurrently: SQLite uses WAL readers/writer mode. Keep
**only one worker** running at a time; v0.2 enforces a cross-process writer lock.

`sync` uses Outlook DASL date filtering (UTC) on LastModificationTime. **Outlook date-filter interpretation can depend on OS regional settings.** If you observe implausibly few results, run `sync --scan-all-metadata`, which scans metadata in every folder and only fetches bodies for changed items. This is slower, but avoids reliance on a locale-specific filter. If the filter raises an error, the worker also falls back to the unfiltered path automatically.

**Deletions/moves/old imports are not reliably discoverable via incremental dates alone.** Run a periodic full reconciliation (weekly/monthly, depending on how often you rearrange old mail):

```powershell
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker reconcile --stores default
```

`reconcile` scans every mailbox folder, updates moved/changed messages, and removes no-longer-present entries **only after each folder's scan finishes**. It can be interrupted and resumed. Run a `sync` pass afterwards to catch changes while the full scan was in progress.

## 5. Connect to Goose Desktop

1. In Goose Desktop: **Extensions → Add custom extension**.
2. **Type:** `Standard IO`; **ID:** `outlook-classic-search`; **Name:** `Outlook Classic Search`.
3. **Command:** use your Python interpreter + module, e.g. `C:\path\to\outlook-classic-mcp\.venv\Scripts\python.exe -m outlook_classic_mcp.mcp_server` (the GUI may separate the executable and arguments; if so, command = absolute `python.exe`, args = `-m` and `outlook_classic_mcp.mcp_server`).
4. Optionally set an environment variable: `OUTLOOK_INDEX_DB=C:\Users\you\AppData\Local\OutlookSQLiteMCP\mail.sqlite3` if using a custom location. The default uses `%LOCALAPPDATA%\OutlookSQLiteMCP\mail.sqlite3`.
5. Restart/re-enable the extension and ask: *"Use search_mail to find invoices from ACME from 2022, then get_mail for the best match."*

Alternatively edit `%USERPROFILE%\.config\goose\config.yaml` (Goose's supported format):

```yaml
extensions:
  outlook_sqlite_search:
    name: Outlook Classic Search
    cmd: 'C:\path\to\outlook-classic-mcp\.venv\Scripts\python.exe'
    args: ['-m', 'outlook_classic_mcp.mcp_server']
    enabled: true
    type: stdio
    timeout: 60
```

The `pip install -e .` step makes the `outlook_classic_mcp` module importable even if Goose starts in a different working directory. Do **not** run the worker as the Goose extension; use `outlook_classic_mcp.mcp_server`.

## MCP tools and search examples

Five tools: `search_mail`, `get_mail`, `index_status`, `list_mail_folders`
(read-only SQLite), and `open_mail_in_outlook` (explicit **local GUI action**).

* `search_mail(query="enterprise agreement", mode="all")` – both words, ranked by BM25 with subject weighting.
* `search_mail(query="security audit", mode="phrase", sender="alice", after="2021-01-01", before="2025-01-01")` – literal phrase, sender and dates.
* `search_mail(query='subject:budget OR "project phoenix"', mode="fts", folder="Sent Items")` – **raw FTS5** grammar, including quoted phrases, OR/NOT, prefixes (`renew*`), and column selectors; trusted users only.
* `get_mail(message_id=1234, max_body_chars=20000)` – full body excerpt for a search result ID.
* `open_mail_in_outlook(message_id=1234)` – display the original item in Outlook Classic.
* `index_status()` – counts, incomplete folders, and item-level errors.

Date filtering uses `received_at`, then `sent_at`, then `modified_at`, in UTC. `after` is inclusive and `before` exclusive. `offset` (up to 10k) and `limit` (up to 100) support pagination. Search does **not** provide semantic/vector search, fuzzy matching, or attachment-body extraction.

## Operational limits and security

1. **Read-only mail content**: the agent cannot send/delete/modify mail from this extension.
   `open_mail_in_outlook` brings a local Outlook window to the foreground (and can
   launch Outlook) when requested. The indexer reads Outlook. The MCP can expose
   locally indexed email contents to the Goose model: inspect your provider and
   privacy settings. Untrusted emails can contain prompt-injection instructions.
2. **Sensitive data at rest**: SQLite holds plaintext email bodies. Keep it in a user-private local directory; use BitLocker/full-disk encryption and appropriate NTFS ACLs. Do not share or sync the database file; backups also contain mail.
3. **No Outlook hooks**: the indexer polls; there is a lag. Periodic full reconciliation is required to purge stale search results. EntryIDs can change on moves; the opener resolves these when identity information is sufficient. No guarantee of a perfect snapshot while Outlook changes.
4. **Attachments**: attachment presence is indexed; file names and contents are not. Encrypted/SMIME bodies may not be available. Outlook errors are recorded under `failures` and shown in `index_status`.
5. **No concurrency between writers**: only one worker process at a time. Multiple Goose MCP readers are fine.
6. **Date filters**: the most dependable option across locales is `sync --scan-all-metadata`; it still avoids reading unchanged message bodies, but walks metadata for everything.
7. **Body truncation and mail content**: Only plain-text `MailItem.Body` is copied; inline/HTML-only rendering and huge threads may be truncated. Changing the body cap does **not** retroactively re-index unchanged messages. Run a complete rebuild into a fresh database to change historical caps consistently.
8. **Outlook cache**: ensure Outlook's Cached Exchange Mode downloads the period you need. Test against an old message known to exist, not only the latest email.

### Database maintenance

```powershell
.\.venv\Scripts\python.exe -m outlook_classic_mcp.worker status
```

Back up the database with SQLite's online backup API, not by copying the main `.sqlite3` while the writer is running without the `-wal`. For best maintenance, stop the worker and Goose before compacting (`VACUUM`), especially after deleting large numbers of mails.

### Engineering notes

* `Folder.GetTable()` enumerates lightweight metadata (EntryID / LastModificationTime / MessageClass) in batches. Only new/changed items trigger `NameSpace.GetItemFromID` and `.Body` reads.
* `bootstrap`/`reconcile` scan the complete **metadata** stream, newest first, and use a persisted cursor to skip already completed batches on restart. A 48-hour overlap protects against time conversion/DST slop and timestamp ties. The cursor is **not** a transactional Outlook snapshot, so a `sync` pass after full scans remains important.
* WAL allows read-only MCP queries during worker writes. FTS5 uses external content with INSERT/DELETE/UPDATE triggers; unchanged messages are *not* re-tokenized.
* SQLite schema version is now 2. Migration is in-place and retains the existing
  FTS5 index. New tiny tombstone records preserve numeric IDs of subsequently
  removed rows; historical rows pruned **before** this upgrade cannot be restored.
  Back up the database before installation.
* Non-mail default folders are skipped; mail-like folders under them are discovered. Virtual Search Folders and duplicate copies in PSTs may still produce duplicates. Outlook folder/store enumeration depends on the user's profile and permissions.
* The worker only handles one COM STA thread, initialized and uninitialized in its process; no COM object crosses processes.

## Tests

```powershell
.\.venv\Scripts\python.exe -m pip install pytest
.\.venv\Scripts\python.exe -m pytest -q
```

Tests exercise SQLite FTS, triggers, pagination, basic schema, and mocked metadata batches without needing Outlook. **Real Windows smoke testing is still needed.**
