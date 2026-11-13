"""Safely display original messages after Outlook moves/reassigns EntryIDs.

Read-only Outlook COM operations. An indexed integer ID is kept usable after
reconciliation removes its original row by retaining a local handle tombstone.
No entire-mailbox body scan: fallback uses per-folder restricted MAPI tables.
"""
from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .db import default_db_path, open_db, upsert_folder
from .outlook import (INTERNET_ID_DASL, SUBJECT_DASL, OutlookSession,
                      mail_identity, normalize_time, read_string)


@dataclass(frozen=True)
class MessageRef:
    message_id: int
    store: str
    entry_id: str | None
    search_key: str | None
    internet_id: str | None
    subject: str
    sender_email: str
    received_at: str | None
    sent_at: str | None


def lookup_reference(db: sqlite3.Connection, message_id: int) -> MessageRef:
    if not isinstance(message_id, int) or isinstance(message_id, bool) or message_id < 1:
        raise ValueError('message_id must be a positive integer from search_mail')
    row = db.execute('''SELECT m.entry_id,s.outlook_id AS outlook_store_id,m.search_key,
               m.internet_message_id,m.subject,m.sender_email,m.received_at,m.sent_at
               FROM messages m JOIN stores s ON s.id=m.store_id WHERE m.id=?''',
               (message_id,)).fetchone()
    if row is None:
        row = db.execute('''SELECT NULL AS entry_id,outlook_store_id,search_key,
                 internet_message_id,subject,sender_email,received_at,sent_at
                 FROM message_handles WHERE message_id=?''', (message_id,)).fetchone()
    if row is None:
        raise LookupError(f'Message {message_id} is no longer in the index.')
    return MessageRef(message_id, str(row['outlook_store_id']), row['entry_id'],
                      row['search_key'], row['internet_message_id'], row['subject'],
                      row['sender_email'], row['received_at'], row['sent_at'])


def lookup_location(db: sqlite3.Connection, message_id: int) -> tuple[str, str]:
    """Compatibility with v2: return current entry + store identifiers."""
    ref = lookup_reference(db, message_id)
    if not ref.entry_id:
        raise LookupError('Old handle requires live Outlook lookup.')
    return ref.entry_id, ref.store


def mail_matches(item: Any, ref: MessageRef) -> bool:
    """Reject mismatches; PR_SEARCH_KEY / Message-ID can be shared by copies."""
    search_key, internet_id = mail_identity(item)
    if ref.search_key and search_key:
        if ref.search_key != search_key:
            return False
        if ref.internet_id and internet_id and ref.internet_id != internet_id:
            return False
        return True
    if ref.internet_id and internet_id:
        if ref.internet_id != internet_id:
            return False
        # Message-ID alone is not necessarily unique; require metadata corroboration.
        if ref.sender_email and ref.sender_email.casefold() != read_string(item, 'SenderEmailAddress').casefold():
            return False
        if ref.received_at and ref.received_at != normalize_time(getattr(item, 'ReceivedTime', None)):
            return False
        if ref.subject and ref.subject != read_string(item, 'Subject'):
            return False
        return bool(ref.sender_email or ref.received_at or ref.subject)
    # Legacy v2 row, before identity extraction was added. Require an exact
    # composite match; cannot prove identity with only a subject or sender.
    if not ref.subject or not (ref.received_at or ref.sent_at):
        return False
    if ref.subject != read_string(item, 'Subject'):
        return False
    if ref.sender_email and ref.sender_email.casefold() != read_string(item, 'SenderEmailAddress').casefold():
        return False
    if ref.received_at and ref.received_at != normalize_time(getattr(item, 'ReceivedTime', None)):
        return False
    if ref.sent_at and ref.sent_at != normalize_time(getattr(item, 'SentOn', None)):
        return False
    return True


def _try_item(session: OutlookSession, entry: str, store: str, ref: MessageRef):
    try:
        item = session.ns.GetItemFromID(entry, store)
        if item is not None and mail_matches(item, ref):
            return item
    except Exception:
        pass
    return None


def _indexed_candidates(db: sqlite3.Connection, ref: MessageRef) -> list[tuple[str, str]]:
    if ref.search_key:
        rows = db.execute('''SELECT m.entry_id,s.outlook_id FROM messages m
             JOIN stores s ON s.id=m.store_id
             WHERE m.id != ? AND m.search_key=? LIMIT 200''',
             (ref.message_id, ref.search_key)).fetchall()
    elif ref.internet_id:
        rows = db.execute('''SELECT m.entry_id,s.outlook_id FROM messages m
             JOIN stores s ON s.id=m.store_id
             WHERE m.id != ? AND m.internet_message_id=? LIMIT 200''',
             (ref.message_id, ref.internet_id)).fetchall()
    elif ref.subject and ref.received_at:
        rows = db.execute('''SELECT m.entry_id,s.outlook_id FROM messages m
             JOIN stores s ON s.id=m.store_id
             WHERE m.id != ? AND m.subject=? AND m.received_at=? LIMIT 200''',
             (ref.message_id, ref.subject, ref.received_at)).fetchall()
    else:
        rows = []
    return [(str(r[0]), str(r[1])) for r in rows]


def _escape_dasl(text: str) -> str:
    return text.replace("'", "''")


def _find_in_outlook(session: OutlookSession, ref: MessageRef,
                     excluded: set[tuple[str, str]]) -> list[tuple[str, str, Any]]:
    """Restricted MAPI lookup. Fall back to subject if Internet-ID is unsupported."""
    filters = []
    if ref.internet_id:
        filters.append(f'@SQL="{INTERNET_ID_DASL}" = \'{_escape_dasl(ref.internet_id)}\'')
    if ref.subject and (ref.received_at or ref.sent_at):
        filters.append(f'@SQL="{SUBJECT_DASL}" = \'{_escape_dasl(ref.subject)}\'')
    if not filters:
        return []

    for filter_sql in filters:
        found: list[tuple[str, str, Any]] = []
        checked = 0
        for store in session.stores('all'):
            store_id = str(store.StoreID)
            for folder in session.folders(store):
                try:
                    if int(folder.DefaultItemType) != 0:
                        continue
                    table = folder.GetTable(filter_sql)
                    table.Columns.RemoveAll()
                    table.Columns.Add('EntryID')
                    count = 0
                    while not table.EndOfTable:
                        count += 1
                        checked += 1
                        if count > 150 or checked > 1500:
                            raise LookupError('Too many matching Outlook rows; cannot safely select one.')
                        row = table.GetNextRow()
                        entry = str(row['EntryID'])
                        if (entry, store_id) in excluded:
                            continue
                        excluded.add((entry, store_id))
                        item = _try_item(session, entry, store_id, ref)
                        if item is not None:
                            found.append((entry, store_id, item))
                            if len(found) > 1:
                                return found
                except LookupError:
                    raise
                except Exception:
                    # Some folders do not support the filter. Never scan unfiltered.
                    continue
        if found:
            return found
    return []


def _remember_location(db: sqlite3.Connection, ref: MessageRef,
                       item: Any, entry: str, store: str) -> None:
    """Best-effort update of an existing indexed row to the moved location.

    No copy is created when the destination was indexed already; the old row
    remains valid through the tombstone after the next reconciliation.
    """
    current = db.execute('SELECT id FROM messages WHERE id=?', (ref.message_id,)).fetchone()
    if current is None:
        return
    if db.execute('''SELECT m.id FROM messages m JOIN stores s ON s.id=m.store_id
                     WHERE m.entry_id=? AND s.outlook_id=? AND m.id<>?''',
                  (entry, store, ref.message_id)).fetchone():
        return
    try:
        folder = item.Parent
        store_row = db.execute('SELECT id FROM stores WHERE outlook_id=?', (store,)).fetchone()
        if store_row:
            store_id = int(store_row['id'])
        else:
            # This can happen when a message moved into an unindexed data file.
            return
        parent_entry = str(folder.EntryID)
        folder_row = upsert_folder(db, store_id, parent_entry,
                                   str(folder.Name), str(folder.FolderPath))
        search_key, internet_id = mail_identity(item)
        db.execute('''UPDATE messages SET store_id=?,folder_id=?,entry_id=?,
                   search_key=coalesce(?,search_key),
                   internet_message_id=coalesce(?,internet_message_id)
                   WHERE id=?''', (store_id, int(folder_row['id']), entry,
                                   search_key, internet_id, ref.message_id))
        db.commit()
    except (sqlite3.Error, AttributeError, TypeError):
        db.rollback()
        # Opening a message is more important than repairing the SQLite cache.


def open_indexed_message(message_id: int, db_path: Path | None = None) -> str:
    """Open a referenced email, after moved EntryIDs, without modifying Outlook."""
    # Writable only for schema upgrades and caching a resolved location.
    with closing(open_db(db_path)) as db:
        ref = lookup_reference(db, message_id)
        with OutlookSession() as session:
            if ref.entry_id:
                # A current EntryID + StoreID is already the strongest COM
                # locator. Legacy rows lack the newer correlation metadata.
                try:
                    item = session.ns.GetItemFromID(ref.entry_id, ref.store)
                    if item is not None and (not (ref.search_key or ref.internet_id)
                                             or mail_matches(item, ref)):
                        item.Display(False)
                        return f'Opened message {message_id} in Outlook Classic.'
                except Exception:
                    pass

            candidates: list[tuple[str, str, Any]] = []
            checked = set()
            for entry, store in _indexed_candidates(db, ref):
                if (entry, store) in checked:
                    continue
                checked.add((entry, store))
                item = _try_item(session, entry, store, ref)
                if item is not None:
                    candidates.append((entry, store, item))
                    if len(candidates) > 1:
                        break
            if len(candidates) > 1:
                raise LookupError('Multiple matching copies of this email exist; refusing to open the wrong one.')
            if not candidates:
                candidates = _find_in_outlook(session, ref, checked)
            if len(candidates) > 1:
                raise LookupError('Multiple matching copies of this email exist; refusing to open the wrong one.')
            if not candidates:
                raise LookupError('Email not found in Outlook. It may have been deleted, moved to a '
                                  'store that is not synchronized, or Outlook may not support the '
                                  'filtered lookup. Try syncing and reconciling.')
            entry, store, item = candidates[0]
            item.Display(False)
            _remember_location(db, ref, item, entry, store)
            return f'Opened message {message_id} in Outlook Classic (relocated).'


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description='Display indexed email in Outlook Classic')
    parser.add_argument('message_id', type=int)
    parser.add_argument('--db', type=Path, default=default_db_path())
    args = parser.parse_args(argv)
    if sys.platform != 'win32':
        print(json.dumps({'ok': False, 'error': 'Requires Windows and Outlook Classic'}))
        return 1
    try:
        msg = open_indexed_message(args.message_id, args.db)
    except Exception as exc:
        print(json.dumps({'ok': False, 'error': str(exc)}))
        return 1
    print(json.dumps({'ok': True, 'message': msg}))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
