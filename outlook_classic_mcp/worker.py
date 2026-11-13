"""Outlook -> SQLite backfill, incremental sync, reconciliation, error reporting.

Run in a normal interactive Windows user session with Classic Outlook configured.
The MCP server never imports this module.
"""
from __future__ import annotations

import argparse
import json
import os
from contextlib import contextmanager
import sqlite3
import sys
import time
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from .integration import consume_sync_request
from .db import UPSERT, default_db_path, existing_batch, now_utc, open_db, status, upsert_folder, upsert_store
from .outlook import MAIL_CLASS_PREFIX, OutlookSession, chunks_from_table, normalize_time, open_table, read_mail, mail_identity



class OutlookClosedError(RuntimeError):
    """The tray watcher pauses when Outlook is closed instead of starting it."""


def is_outlook_running() -> bool:
    """Non-launching check. If the system lacks psutil, consult the COM ROT."""
    try:
        import psutil
    except ImportError:
        # GetActiveObject never launches Outlook (unlike Dispatch).
        try:
            import pythoncom
            from win32com.client import GetActiveObject
            pythoncom.CoInitialize()
            try:
                obj = GetActiveObject('Outlook.Application')
                del obj
                return True
            finally:
                pythoncom.CoUninitialize()
        except Exception:
            return False
    else:
        try:
            return any(p.info['name'] and p.info['name'].lower() == 'outlook.exe'
                       for p in psutil.process_iter(['name']))
        except (psutil.Error, OSError):
            return False


@contextmanager
def worker_lock(db_path: Path):
    """Prevent overlapping tray/manual writers, including across processes."""
    lock_path = db_path.with_name(db_path.name + '.writer.lock')
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    with open(lock_path, 'a+b') as handle:
        try:
            if os.name == 'nt':
                import msvcrt
                if handle.tell() == 0 and lock_path.stat().st_size == 0:
                    handle.write(b'\0')
                    handle.flush()
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            raise RuntimeError('Another index worker is already running. '
                               'Exit the tray app or stop the other watch process.') from exc
        try:
            yield
        finally:
            if os.name == 'nt':
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle, fcntl.LOCK_UN)


def utc_dt(iso: str) -> datetime:
    return datetime.strptime(iso, '%Y-%m-%dT%H:%M:%SZ').replace(tzinfo=timezone.utc)


def start_full_scan(db: sqlite3.Connection, folder_id: int, state: str, refresh: bool) -> tuple[str, str | None] | None:
    row = db.execute('SELECT scan_state,scan_generation,scan_cursor FROM folders WHERE id=?', (folder_id,)).fetchone()
    if row['scan_state'] == 'running' and row['scan_generation']:
        return row['scan_generation'], row['scan_cursor']
    if not refresh and row['scan_state'] == 'complete':
        return None
    generation = str(uuid.uuid4())
    db.execute('UPDATE folders SET scan_state="running",scan_generation=?,scan_cursor=NULL,last_error=NULL WHERE id=?',
               (generation, folder_id))
    db.commit()
    return generation, None


def process_block(db: sqlite3.Connection, session: OutlookSession, store_id: int, store_outlook_id: str,
                  folder_id: int, block: list[tuple[Any, Any, Any]], generation: str | None,
                  max_body_chars: int) -> tuple[int, int, int]:
    """Fetch bodies only if missing/changed. Batch writes including 'seen' stamps."""
    rows = []
    for entry, modified, klass in block:
        if entry and str(klass).startswith(MAIL_CLASS_PREFIX):
            rows.append((str(entry), normalize_time(modified)))
    existing = existing_batch(db, store_id, [entry for entry, _ in rows])
    seen_ids: list[int] = []
    updates: list[dict[str, Any]] = []
    failures: list[tuple[int, int, str, str, str]] = []
    skips = 0
    for entry, modified in rows:
        old = existing.get(entry)
        if old and old['modified_at'] == modified and old['folder_id'] == folder_id:
            skips += 1
            if generation:
                seen_ids.append(old['id'])
            continue
        try:
            item = session.ns.GetItemFromID(entry, store_outlook_id)
            # Some IPM.Note subclasses may not be MailItems. Read only if the
            # required properties are actually exposed.
            data = read_mail(item, store_id, folder_id, entry, modified, generation, max_body_chars)
            updates.append(data)
        except Exception as exc:
            if old and generation:
                seen_ids.append(old['id'])  # existing content must not be pruned
            failures.append((store_id, folder_id, entry, repr(exc)[:1500], now_utc()))
    with db:
        if seen_ids:
            db.executemany('UPDATE messages SET seen_generation=? WHERE id=?',
                           [(generation, msg_id) for msg_id in seen_ids])
        if updates:
            db.executemany(UPSERT, updates)
            db.executemany('DELETE FROM failures WHERE store_id=? AND entry_id=?',
                           [(store_id, update['entry_id']) for update in updates])
        if failures:
            db.executemany('''INSERT INTO failures(store_id,folder_id,entry_id,details,last_attempt_at)
               VALUES (?,?,?,?,?) ON CONFLICT(store_id,entry_id) DO UPDATE SET
               folder_id=excluded.folder_id, details=excluded.details,
               last_attempt_at=excluded.last_attempt_at''', failures)
    return len(updates), skips, len(failures)


def index_folder(db: sqlite3.Connection, session: OutlookSession, folder: Any,
                 store_id: int, store_outlook_id: str, mode: str, refresh: bool,
                 batch_size: int, max_body_chars: int, lookback_days: int,
                 scan_all_metadata: bool, allow_large_prune: bool = False) -> tuple[int, int, int]:
    folder_row = upsert_folder(db, store_id, str(folder.EntryID),
                               str(folder.Name), str(folder.FolderPath))
    folder_id = folder_row['id']
    db.commit()
    # All subfolders are traversed, but folders for calendar/contacts/tasks are
    # not indexed; Outlook OlItemType.olMailItem == 0.
    if int(folder.DefaultItemType) != 0:
        return 0, 0, 0
    generation: str | None = None
    cutoff: datetime | None = None
    if mode == 'full':
        full = start_full_scan(db, folder_id, folder_row['scan_state'], refresh)
        if full is None:
            return 0, 0, 0
        generation, cursor = full
        if cursor:
            # Cover same-minute timestamps, local timezone/DST conversions, and
            # changing rows. Duplicates are skipped using EntryID + modified_at.
            cutoff = utc_dt(cursor) + timedelta(days=2)
    else:
        # A lookback window intentionally captures edits to old emails made recently.
        # A full reconciliation is still necessary to discover deletions and moves.
        cutoff = datetime.now(timezone.utc) - timedelta(days=lookback_days + 2)
    path = str(folder.FolderPath)
    print(f'[{mode}] {path}', flush=True)
    t0 = time.monotonic()
    added = unchanged = errors = scanned = 0
    last_cursor: str | None = None
    try:
        table = open_table(folder, cutoff, mode, no_filter=(scan_all_metadata or mode == 'full'))
        for block in chunks_from_table(table, batch_size):
            # Some regional Outlook installations do not honor the filter as
            # expected. Never advance a full-scan cursor across skipped rows.
            # With a failed filter, rows are still sorted for resume.
            if mode == 'full' and cutoff is not None:
                block = [r for r in block if (normalize_time(r[1]) or '0000') <=
                         cutoff.strftime('%Y-%m-%dT%H:%M:%SZ')]
            if not block:
                continue
            n, s, e = process_block(db, session, store_id, store_outlook_id,
                                    folder_id, block, generation, max_body_chars)
            added += n
            unchanged += s
            errors += e
            scanned += len(block)
            if mode == 'full':
                last_cursor = normalize_time(block[-1][1])
                if last_cursor:
                    # The scan cursor and completed batch are persisted independently.
                    # Replaying a batch after a crash is idempotent.
                    db.execute('UPDATE folders SET scan_cursor=? WHERE id=?', (last_cursor, folder_id))
                    db.commit()
            elif scanned % (batch_size * 10) == 0:
                print(f'  scanned={scanned:,} new/changed={added:,} unchanged={unchanged:,}', flush=True)
            if mode == 'full' and scanned % (batch_size * 10) == 0:
                print(f'  scanned={scanned:,} new/changed={added:,} unchanged={unchanged:,}', flush=True)
        with db:
            if mode == 'full':
                previous_count = db.execute('SELECT count(*) FROM messages WHERE folder_id=?',
                                            (folder_id,)).fetchone()[0]
                stale_count = db.execute('''SELECT count(*) FROM messages WHERE folder_id=?
                    AND (seen_generation IS NULL OR seen_generation != ?)''',
                    (folder_id, generation)).fetchone()[0]
                if (not allow_large_prune and previous_count >= 500 and
                        stale_count > max(500, int(previous_count * 0.25))):
                    raise RuntimeError(
                        f'Prune safety stop: {stale_count:,}/{previous_count:,} '
                        'indexed messages disappeared from folder. Check Outlook cache/'
                        'permissions; rerun with --allow-large-prune if intentional.')
                # Remove deleted/moved items only after the FULL scan of THIS folder
                # completed successfully; a failed/incomplete scan cannot prune.
                db.execute('DELETE FROM messages WHERE folder_id=? '
                           'AND (seen_generation IS NULL OR seen_generation != ?)',
                           (folder_id, generation))
                db.execute('UPDATE folders SET scan_state="complete",scan_generation=NULL,'
                           'scan_cursor=NULL,last_full_at=?,last_error=NULL WHERE id=?', (now_utc(), folder_id))
            else:
                db.execute('UPDATE folders SET last_incremental_at=?,last_error=NULL WHERE id=?',
                           (now_utc(), folder_id))
    except Exception as exc:
        with db:
            db.execute('UPDATE folders SET last_error=? WHERE id=?', (repr(exc)[:1000], folder_id))
        raise
    print(f'  done rows={scanned:,} updated={added:,} unchanged={unchanged:,} '
          f'errors={errors:,} seconds={time.monotonic()-t0:.1f}', flush=True)
    return added, unchanged, errors


def run_index(db: sqlite3.Connection, mode: str, stores: str, *, refresh: bool = False,
              batch_size: int = 250, max_body_chars: int = 65536,
              lookback_days: int = 30, scan_all_metadata: bool = False,
              folder_contains: str = '', allow_large_prune: bool = False,
              require_outlook_running: bool = False,
              reconcile_due_days: int = 0) -> int:
    total_updated = total_failures = total_unchanged = folders_ok = folders_failed = 0
    if require_outlook_running and not is_outlook_running():
        raise OutlookClosedError('Outlook Classic is closed; waiting to reopen.')
    with OutlookSession() as session:
        for store in session.stores(stores):
            outlook_id = str(store.StoreID)
            store_name = str(store.DisplayName)
            store_id = upsert_store(db, outlook_id, store_name)
            db.commit()
            print(f'\nStore: {store_name}', flush=True)
            for folder in session.folders(store):
                if require_outlook_running and not is_outlook_running():
                    raise OutlookClosedError('Outlook Classic closed during sync; resuming on reopen.')
                path = str(folder.FolderPath)
                if folder_contains and folder_contains.casefold() not in path.casefold():
                    continue
                if mode == 'full' and refresh and reconcile_due_days:
                    previous = db.execute('''SELECT scan_state,last_full_at FROM folders
                        WHERE store_id=? AND outlook_id=?''',
                        (store_id, str(folder.EntryID))).fetchone()
                    if (previous is not None and previous['scan_state'] == 'complete'
                            and previous['last_full_at'] is not None and
                            previous['last_full_at'] >=
                            (datetime.now(timezone.utc) - timedelta(days=reconcile_due_days)).strftime('%Y-%m-%dT%H:%M:%SZ')):
                        continue
                try:
                    n, s, e = index_folder(db, session, folder, store_id, outlook_id,
                                             mode, refresh, batch_size, max_body_chars,
                                             lookback_days, scan_all_metadata, allow_large_prune)
                    total_updated += n
                    total_unchanged += s
                    total_failures += e
                    folders_ok += 1
                except (KeyboardInterrupt, OutlookClosedError):
                    raise
                except Exception as exc:
                    if require_outlook_running and not is_outlook_running():
                        raise OutlookClosedError('Outlook closed during sync.') from exc
                    folders_failed += 1
                    print(f'ERROR in folder {path}: {exc!r}', file=sys.stderr, flush=True)
    print(f'\nDone. updated={total_updated:,} unchanged={total_unchanged:,} '
          f'failed_items={total_failures:,} failed_folders={folders_failed:,}', flush=True)
    return 1 if folders_failed or total_failures else 0


def backfill_identities(db: sqlite3.Connection, limit: int = 10000) -> int:
    """Enrich pre-upgrade rows using ONLY MAPI identity properties (no bodies).

    One bounded batch per invocation; persistent high-water cursor makes this
    resumable across crashes and avoids endlessly retrying missing message IDs.
    """
    db.execute('INSERT OR IGNORE INTO identity_backfill_state(id,cursor_id) VALUES (1,0)')
    db.commit()
    cursor = int(db.execute('SELECT cursor_id FROM identity_backfill_state WHERE id=1').fetchone()[0])
    rows = db.execute('''SELECT m.id,m.entry_id,s.outlook_id AS outlook_store_id
        FROM messages m JOIN stores s ON s.id=m.store_id
        WHERE m.id>? AND (m.search_key IS NULL OR m.internet_message_id IS NULL)
        ORDER BY m.id LIMIT ?''', (cursor, limit)).fetchall()
    if not rows:
        print('Identity backfill complete for currently indexed messages.', flush=True)
        return 0
    updated = errors = 0
    with OutlookSession() as session:
        for number, row in enumerate(rows, 1):
            try:
                item = session.ns.GetItemFromID(str(row['entry_id']), str(row['outlook_store_id']))
                search_key, internet_id = mail_identity(item)
                if search_key or internet_id:
                    db.execute('''UPDATE messages SET search_key=coalesce(?,search_key),
                        internet_message_id=coalesce(?,internet_message_id)
                        WHERE id=? AND entry_id=?''',
                        (search_key, internet_id, row['id'], row['entry_id']))
                    updated += 1
            except Exception:
                errors += 1  # stale/removed item; opening fallback remains available
            if number % 200 == 0 or number == len(rows):
                db.execute('UPDATE identity_backfill_state SET cursor_id=? WHERE id=1',
                           (row['id'],))
                db.commit()
    print(f'Identity backfill checked={len(rows):,} enriched={updated:,} '
          f'unavailable={errors:,} cursor={rows[-1]["id"]:,}. '
          'Run again to continue.', flush=True)
    return 0


def doctor() -> int:
    """Non-mutating COM smoke test before indexing a huge mailbox."""
    with OutlookSession() as session:
        store = session.ns.DefaultStore
        print('Outlook default store:', str(store.DisplayName), flush=True)
        for folder in session.folders(store):
            try:
                if int(folder.DefaultItemType) != 0:
                    continue
                table = open_table(folder, None, 'full')
                sample = next(chunks_from_table(table, 2), [])
                if not sample:
                    continue
                print('Sample folder:', str(folder.FolderPath), flush=True)
                print('Sample metadata rows:', len(sample), flush=True)
                print('Sample EntryID present:', bool(sample[0][0]), flush=True)
                print('Sample modified UTC:', normalize_time(sample[0][1]), flush=True)
                print('Sample message class:', str(sample[0][2]), flush=True)
                print('COM metadata enumeration works. No messages were indexed.', flush=True)
                return 0
            except Exception as exc:
                print(f'WARNING: {folder.Name}: {exc}', file=sys.stderr, flush=True)
        print('No readable mail metadata found. Check Outlook cache/profile.', file=sys.stderr)
        return 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description='Local Outlook Classic -> SQLite FTS5 worker')
    parser.add_argument('action', choices=['doctor', 'bootstrap', 'sync', 'reconcile', 'watch', 'status', 'backfill-ids'])
    parser.add_argument('--db', type=Path, default=default_db_path(), help='SQLite path (also OUTLOOK_INDEX_DB)')
    parser.add_argument('--stores', choices=['default', 'all'], default='default')
    parser.add_argument('--batch-size', type=int, default=250)
    parser.add_argument('--backfill-limit', type=int, default=10000,
                        help='backfill-ids: max existing mail IDs checked per run (default 10000)')
    parser.add_argument('--max-body-chars', type=int, default=65536)
    parser.add_argument('--lookback-days', type=int, default=30)
    parser.add_argument('--every-minutes', type=int, default=15, help='for watch')
    parser.add_argument('--reconcile-every-days', type=int, default=0,
                        help='watch: reconcile folders that were last fully scanned N+ days ago (0 disables)')
    parser.add_argument('--only-if-outlook-running', action='store_true',
                        help='watch: pause without launching Outlook and resume when reopened')
    parser.add_argument('--scan-all-metadata', action='store_true',
                        help='disable locale-sensitive Outlook date filtering (slower but safer)')
    parser.add_argument('--folder-contains', default='', help='index only matching folder paths (testing)')
    parser.add_argument('--allow-large-prune', action='store_true',
                        help='allow full reconcile to remove >25%% of a folder with 500+ emails')
    args = parser.parse_args(argv)
    if not 1 <= args.backfill_limit <= 1000000:
        parser.error('--backfill-limit must be between 1 and 1000000')
    if not 10 <= args.batch_size <= 1000:
        parser.error('--batch-size must be between 10 and 1000')
    if not 1000 <= args.max_body_chars <= 1000000:
        parser.error('--max-body-chars must be between 1000 and 1000000')
    if not 1 <= args.lookback_days <= 3650:
        parser.error('--lookback-days must be 1..3650')
    if not 1 <= args.every_minutes <= 1440:
        parser.error('--every-minutes must be 1..1440')
    if not 0 <= args.reconcile_every_days <= 365:
        parser.error('--reconcile-every-days must be 0..365')
    with open_db(args.db) as db:
        if args.action == 'status':
            print(json.dumps(status(db, args.db), indent=2))
            return 0
        if sys.platform != 'win32':
            parser.error('The index worker requires Windows and Outlook Classic; the MCP can run anywhere.')
        opts = dict(stores=args.stores, batch_size=args.batch_size,
                    max_body_chars=args.max_body_chars, lookback_days=args.lookback_days,
                    scan_all_metadata=args.scan_all_metadata, folder_contains=args.folder_contains,
                    allow_large_prune=args.allow_large_prune)
        if args.action == 'doctor':
            return doctor()
        try:
            with worker_lock(args.db):
                if args.action == 'backfill-ids':
                    return backfill_identities(db, args.backfill_limit)
                if args.action == 'watch':
                    # Poll while Outlook is closed. Don't launch Outlook just to sync.
                    interval = args.every_minutes * 60
                    next_run = 0.0
                    was_running = False
                    next_reconcile_check = 0.0
                    print('Watch started. Press Ctrl+C to stop.', flush=True)
                    while True:
                        try:
                            running = (not args.only_if_outlook_running or is_outlook_running())
                            if running and not was_running:
                                print('Outlook available; scheduling sync.', flush=True)
                                next_run = 0.0
                            if not running and was_running:
                                print('Outlook closed; sync paused.', flush=True)
                            was_running = running
                            request = args.db.with_name(args.db.name + '.sync-request')
                            if running and (time.monotonic() >= next_run or request.exists()):
                                consume_sync_request(args.db)
                                try:
                                    run_index(db, 'incremental', require_outlook_running=
                                              args.only_if_outlook_running, **opts)
                                except OutlookClosedError:
                                    print('Outlook closed during sync; waiting to reopen.', flush=True)
                                    was_running = False
                                except Exception as exc:
                                    print(f'Sync failed (retrying next interval): {exc!r}',
                                          file=sys.stderr, flush=True)
                                next_run = time.monotonic() + interval
                            # Weekly (tray default) reconciliation catches deletes/moves.
                            # Each completed folder is skipped until its last_full_at is due,
                            # so closing/reopening Outlook does not restart completed work.
                            if (running and args.reconcile_every_days and
                                    time.monotonic() >= next_reconcile_check):
                                try:
                                    run_index(db, 'full', refresh=True, reconcile_due_days=
                                              args.reconcile_every_days,
                                              require_outlook_running=args.only_if_outlook_running,
                                              **opts)
                                except OutlookClosedError:
                                    print('Reconciliation paused (Outlook closed).', flush=True)
                                    was_running = False
                                except Exception as exc:
                                    print(f'Reconciliation problem: {exc!r}', file=sys.stderr,
                                          flush=True)
                                next_reconcile_check = time.monotonic() + 6 * 3600
                            time.sleep(5 if args.only_if_outlook_running else min(5, interval))
                        except KeyboardInterrupt:
                            print('Watch stopped.', flush=True)
                            return 0
                return run_index(db, 'full' if args.action != 'sync' else 'incremental',
                                 refresh=args.action == 'reconcile', **opts)
        except RuntimeError as exc:
            print(f'ERROR: {exc}', file=sys.stderr, flush=True)
            return 1



if __name__ == '__main__':
    sys.exit(main())
