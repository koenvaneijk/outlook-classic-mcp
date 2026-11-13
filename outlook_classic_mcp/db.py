"""SQLite storage, FTS5 search and schema. No Outlook/Windows dependencies."""
from __future__ import annotations

import os
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def default_db_path() -> Path:
    base = Path(os.environ.get('LOCALAPPDATA') or (Path.home() / '.local' / 'share'))
    return Path(os.environ.get('OUTLOOK_INDEX_DB', str(base / 'OutlookSQLiteMCP' / 'mail.sqlite3'))).expanduser().resolve()


def now_utc() -> str:
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')


def open_db(path: Path | None = None, *, readonly: bool = False) -> sqlite3.Connection:
    path = (path or default_db_path()).expanduser().resolve()
    if readonly:
        if not path.is_file():
            raise FileNotFoundError(f'Index not found at {path}. Run the bootstrap command first.')
        conn = sqlite3.connect(path.as_uri() + '?mode=ro', uri=True, timeout=30)
        conn.execute('PRAGMA query_only=ON')
    else:
        path.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(path, timeout=30)
        conn.execute('PRAGMA journal_mode=WAL')
        conn.execute('PRAGMA synchronous=NORMAL')
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA busy_timeout=30000')
    if not readonly:
        init_schema(conn)
    return conn


SCHEMA = '''
CREATE TABLE IF NOT EXISTS stores (
  id INTEGER PRIMARY KEY,
  outlook_id TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS folders (
  id INTEGER PRIMARY KEY,
  store_id INTEGER NOT NULL REFERENCES stores(id),
  outlook_id TEXT NOT NULL,
  name TEXT NOT NULL,
  path TEXT NOT NULL,
  scan_state TEXT NOT NULL DEFAULT 'never',
  scan_generation TEXT,
  scan_cursor TEXT,
  last_full_at TEXT,
  last_incremental_at TEXT,
  last_error TEXT,
  UNIQUE(store_id, outlook_id)
);
CREATE INDEX IF NOT EXISTS folders_path_idx ON folders(path);
CREATE TABLE IF NOT EXISTS messages (
  id INTEGER PRIMARY KEY,
  store_id INTEGER NOT NULL REFERENCES stores(id),
  folder_id INTEGER NOT NULL REFERENCES folders(id),
  entry_id TEXT NOT NULL,
  search_key TEXT,
  internet_message_id TEXT,
  modified_at TEXT,
  received_at TEXT,
  sent_at TEXT,
  subject TEXT NOT NULL DEFAULT '',
  sender_name TEXT NOT NULL DEFAULT '',
  sender_email TEXT NOT NULL DEFAULT '',
  recipients_to TEXT NOT NULL DEFAULT '',
  recipients_cc TEXT NOT NULL DEFAULT '',
  conversation_id TEXT NOT NULL DEFAULT '',
  body TEXT NOT NULL DEFAULT '',
  body_truncated INTEGER NOT NULL DEFAULT 0,
  is_unread INTEGER NOT NULL DEFAULT 0,
  has_attachments INTEGER NOT NULL DEFAULT 0,
  seen_generation TEXT,
  UNIQUE(store_id, entry_id)
);
CREATE INDEX IF NOT EXISTS messages_folder_idx ON messages(folder_id);
CREATE INDEX IF NOT EXISTS messages_received_idx ON messages(received_at DESC);
CREATE INDEX IF NOT EXISTS messages_sender_idx ON messages(sender_email);
CREATE INDEX IF NOT EXISTS messages_seen_idx ON messages(folder_id, seen_generation);
CREATE INDEX IF NOT EXISTS messages_search_key_idx ON messages(search_key) WHERE search_key IS NOT NULL;
CREATE INDEX IF NOT EXISTS messages_internet_id_idx ON messages(internet_message_id) WHERE internet_message_id IS NOT NULL;
CREATE INDEX IF NOT EXISTS messages_subject_received_idx ON messages(subject, received_at);
-- Preserve stale numeric Goose handles for messages removed by reconciliation.
CREATE TABLE IF NOT EXISTS identity_backfill_state (
  id INTEGER PRIMARY KEY CHECK(id=1),
  cursor_id INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS message_handles (
  message_id INTEGER PRIMARY KEY,
  outlook_store_id TEXT NOT NULL,
  search_key TEXT,
  internet_message_id TEXT,
  subject TEXT NOT NULL DEFAULT '',
  sender_email TEXT NOT NULL DEFAULT '',
  received_at TEXT,
  sent_at TEXT
);
CREATE TRIGGER IF NOT EXISTS messages_keep_handle AFTER DELETE ON messages BEGIN
  INSERT INTO message_handles(message_id,outlook_store_id,search_key,internet_message_id,
      subject,sender_email,received_at,sent_at)
  SELECT old.id, s.outlook_id, old.search_key, old.internet_message_id,
      old.subject, old.sender_email, old.received_at, old.sent_at
  FROM stores s WHERE s.id=old.store_id
  ON CONFLICT(message_id) DO UPDATE SET
    outlook_store_id=excluded.outlook_store_id,search_key=excluded.search_key,
    internet_message_id=excluded.internet_message_id,subject=excluded.subject,
    sender_email=excluded.sender_email,received_at=excluded.received_at,sent_at=excluded.sent_at;
END;
CREATE TABLE IF NOT EXISTS failures (
  id INTEGER PRIMARY KEY,
  store_id INTEGER NOT NULL,
  folder_id INTEGER NOT NULL,
  entry_id TEXT NOT NULL,
  details TEXT NOT NULL,
  last_attempt_at TEXT NOT NULL,
  UNIQUE(store_id, entry_id)
);
CREATE VIRTUAL TABLE IF NOT EXISTS mail_fts USING fts5(
  subject, sender_name, sender_email, recipients_to, recipients_cc, body,
  content='messages', content_rowid='id', tokenize='unicode61 remove_diacritics 2'
);
CREATE TRIGGER IF NOT EXISTS messages_ai AFTER INSERT ON messages BEGIN
  INSERT INTO mail_fts(rowid,subject,sender_name,sender_email,recipients_to,recipients_cc,body)
  VALUES(new.id,new.subject,new.sender_name,new.sender_email,new.recipients_to,new.recipients_cc,new.body);
END;
CREATE TRIGGER IF NOT EXISTS messages_ad AFTER DELETE ON messages BEGIN
  INSERT INTO mail_fts(mail_fts,rowid,subject,sender_name,sender_email,recipients_to,recipients_cc,body)
  VALUES('delete',old.id,old.subject,old.sender_name,old.sender_email,old.recipients_to,old.recipients_cc,old.body);
END;
CREATE TRIGGER IF NOT EXISTS messages_au AFTER UPDATE OF
 subject,sender_name,sender_email,recipients_to,recipients_cc,body ON messages BEGIN
  INSERT INTO mail_fts(mail_fts,rowid,subject,sender_name,sender_email,recipients_to,recipients_cc,body)
  VALUES('delete',old.id,old.subject,old.sender_name,old.sender_email,old.recipients_to,old.recipients_cc,old.body);
  INSERT INTO mail_fts(rowid,subject,sender_name,sender_email,recipients_to,recipients_cc,body)
  VALUES(new.id,new.subject,new.sender_name,new.sender_email,new.recipients_to,new.recipients_cc,new.body);
END;
'''


def init_schema(conn: sqlite3.Connection) -> None:
    version = conn.execute('PRAGMA user_version').fetchone()[0]
    if version not in (0, 1, 2):
        raise RuntimeError(f'Unsupported database schema version {version}')
    # In-place migration; preserves the original million-message FTS index.
    if version == 1:
        columns = {r[1] for r in conn.execute('PRAGMA table_info(messages)')}
        if 'search_key' not in columns:
            conn.execute('ALTER TABLE messages ADD COLUMN search_key TEXT')
        if 'internet_message_id' not in columns:
            conn.execute('ALTER TABLE messages ADD COLUMN internet_message_id TEXT')
        conn.commit()
    conn.executescript(SCHEMA)
    conn.execute('PRAGMA user_version=2')
    conn.commit()


def upsert_store(conn: sqlite3.Connection, outlook_id: str, name: str) -> int:
    conn.execute('INSERT INTO stores(outlook_id,name) VALUES (?,?) '
                 'ON CONFLICT(outlook_id) DO UPDATE SET name=excluded.name', (outlook_id, name))
    return int(conn.execute('SELECT id FROM stores WHERE outlook_id=?', (outlook_id,)).fetchone()[0])


def upsert_folder(conn: sqlite3.Connection, store_id: int, outlook_id: str, name: str, path: str) -> sqlite3.Row:
    conn.execute('INSERT INTO folders(store_id,outlook_id,name,path) VALUES (?,?,?,?) '
                 'ON CONFLICT(store_id,outlook_id) DO UPDATE SET name=excluded.name,path=excluded.path',
                 (store_id, outlook_id, name, path))
    return conn.execute('SELECT * FROM folders WHERE store_id=? AND outlook_id=?',
                        (store_id, outlook_id)).fetchone()


def existing_batch(conn: sqlite3.Connection, store_id: int, entry_ids: list[str]) -> dict[str, sqlite3.Row]:
    if not entry_ids:
        return {}
    placeholders = ','.join('?' for _ in entry_ids)
    rows = conn.execute(f'SELECT id,entry_id,modified_at,folder_id FROM messages '
                        f'WHERE store_id=? AND entry_id IN ({placeholders})', [store_id, *entry_ids])
    return {r['entry_id']: r for r in rows}


UPSERT = '''
INSERT INTO messages(id,store_id,folder_id,entry_id,search_key,internet_message_id,
 modified_at,received_at,sent_at,subject,sender_name,sender_email,recipients_to,
 recipients_cc,conversation_id,body,body_truncated,is_unread,has_attachments,seen_generation)
VALUES((SELECT MAX(COALESCE((SELECT MAX(id) FROM messages),0),
                   COALESCE((SELECT MAX(message_id) FROM message_handles),0)) + 1),
 :store_id,:folder_id,:entry_id,:search_key,:internet_message_id,
 :modified_at,:received_at,:sent_at,:subject,:sender_name,:sender_email,
 :recipients_to,:recipients_cc,:conversation_id,:body,:body_truncated,
 :is_unread,:has_attachments,:seen_generation)
ON CONFLICT(store_id,entry_id) DO UPDATE SET
 folder_id=excluded.folder_id,search_key=coalesce(excluded.search_key,messages.search_key),
 internet_message_id=coalesce(excluded.internet_message_id,messages.internet_message_id),
 modified_at=excluded.modified_at,received_at=excluded.received_at,
 sent_at=excluded.sent_at,subject=excluded.subject,sender_name=excluded.sender_name,
 sender_email=excluded.sender_email,recipients_to=excluded.recipients_to,
 recipients_cc=excluded.recipients_cc,conversation_id=excluded.conversation_id,
 body=excluded.body,body_truncated=excluded.body_truncated,
 is_unread=excluded.is_unread,has_attachments=excluded.has_attachments,
 seen_generation=excluded.seen_generation
'''

def fts_query(query: str, mode: str = 'all') -> str:
    """Safe user-entered search; 'fts' is opt-in advanced FTS5 syntax."""
    if len(query) > 2000:
        raise ValueError('query too long (max 2000 characters)')
    if mode == 'fts':
        if not query.strip():
            raise ValueError('query cannot be empty')
        return query
    if mode == 'phrase':
        if not query.strip():
            raise ValueError('query cannot be empty')
        return '"' + query.strip().replace('"', '""') + '"'
    tokens = re.findall(r'\w+', query, re.UNICODE)
    if not tokens:
        raise ValueError('query needs at least one word')
    if mode not in ('all', 'any'):
        raise ValueError('mode must be all, any, phrase, or fts')
    return (' AND ' if mode == 'all' else ' OR ').join('"' + x.replace('"', '""') + '"' for x in tokens)


def _check_date(date: str | None) -> str | None:
    if date is None or date == '':
        return None
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', date):
        raise ValueError('Dates must use YYYY-MM-DD')
    try:
        datetime.strptime(date, '%Y-%m-%d')
    except ValueError as exc:
        raise ValueError('Dates must use YYYY-MM-DD') from exc
    return date


def search(conn: sqlite3.Connection, query: str, mode: str = 'all', sender: str = '',
           folder: str = '', after: str | None = None, before: str | None = None,
           limit: int = 20, offset: int = 0) -> list[dict[str, Any]]:
    if not 1 <= limit <= 100:
        raise ValueError('limit must be 1..100')
    if not 0 <= offset <= 10000:
        raise ValueError('offset must be 0..10000')
    after, before = _check_date(after), _check_date(before)
    clauses = ['mail_fts MATCH ?']
    params: list[Any] = [fts_query(query, mode)]
    if sender:
        clauses.append('instr(lower(m.sender_email || " " || m.sender_name),lower(?)) > 0')
        params.append(sender)
    if folder:
        clauses.append('instr(lower(f.path),lower(?)) > 0')
        params.append(folder)
    if after:
        clauses.append('coalesce(m.received_at,m.sent_at,m.modified_at) >= ?')
        params.append(after + 'T00:00:00Z')
    if before:
        clauses.append('coalesce(m.received_at,m.sent_at,m.modified_at) < ?')
        params.append(before + 'T00:00:00Z')
    sql = f'''SELECT m.id, m.subject, m.sender_name, m.sender_email, m.received_at, m.sent_at,
      m.is_unread, m.has_attachments, f.path AS folder, s.name AS store,
      snippet(mail_fts,5,'[',']','…',20) AS snippet,
      bm25(mail_fts, 8.0, 3.0, 5.0, 1.0, 1.0, 1.0) AS rank
      FROM mail_fts JOIN messages m ON m.id=mail_fts.rowid
      JOIN folders f ON f.id=m.folder_id JOIN stores s ON s.id=m.store_id
      WHERE {' AND '.join(clauses)} ORDER BY rank LIMIT ? OFFSET ?'''
    try:
        return [dict(r) for r in conn.execute(sql, [*params, limit, offset])]
    except sqlite3.OperationalError as exc:
        if mode == 'fts':
            raise ValueError(f'Invalid FTS5 expression: {exc}') from exc
        raise


def get_message(conn: sqlite3.Connection, message_id: int, max_body_chars: int = 20000) -> dict[str, Any] | None:
    if not 0 <= max_body_chars <= 200000:
        raise ValueError('max_body_chars must be 0..200000')
    r = conn.execute('''SELECT m.*, f.path AS folder, s.name AS store FROM messages m
             JOIN folders f ON f.id=m.folder_id JOIN stores s ON s.id=m.store_id
             WHERE m.id=?''', (message_id,)).fetchone()
    original_id = message_id
    if not r:
        old = conn.execute('SELECT * FROM message_handles WHERE message_id=?', (message_id,)).fetchone()
        if old:
            if old['search_key']:
                hits = conn.execute('SELECT id FROM messages WHERE search_key=? LIMIT 3',
                                    (old['search_key'],)).fetchall()
            elif old['internet_message_id']:
                hits = conn.execute('''SELECT id FROM messages
                    WHERE internet_message_id=? AND subject=? AND sender_email=? LIMIT 3''',
                    (old['internet_message_id'], old['subject'], old['sender_email'])).fetchall()
            elif old['received_at'] and old['subject']:
                hits = conn.execute('''SELECT id FROM messages
                    WHERE subject=? AND received_at=? AND sender_email=? LIMIT 3''',
                    (old['subject'], old['received_at'], old['sender_email'])).fetchall()
            else:
                hits = []
            if len(hits) == 1:
                message_id = hits[0]['id']
                r = conn.execute('''SELECT m.*,f.path AS folder,s.name AS store FROM messages m
                    JOIN folders f ON f.id=m.folder_id JOIN stores s ON s.id=m.store_id
                    WHERE m.id=?''', (message_id,)).fetchone()
    if not r:
        return None
    msg = dict(r)
    if message_id != original_id:
        msg['requested_message_id'] = original_id
        msg['relocated_to_message_id'] = message_id
    msg['body'] = msg['body'][:max_body_chars]
    msg['body_response_truncated'] = bool(msg['body_truncated'] or len(r['body']) > max_body_chars)
    msg.pop('seen_generation', None)
    # Don't expose raw COM IDs by default; they are never needed for read-only searching.
    msg.pop('entry_id', None)
    return msg


def status(conn: sqlite3.Connection, path: Path) -> dict[str, Any]:
    count = conn.execute('SELECT count(*) FROM messages').fetchone()[0]
    folders = conn.execute('SELECT scan_state,count(*) AS count FROM folders GROUP BY scan_state').fetchall()
    errors = conn.execute('SELECT count(*) FROM failures').fetchone()[0]
    recent_failures = [dict(r) for r in conn.execute('''SELECT f.path, e.details, e.last_attempt_at
        FROM failures e JOIN folders f ON f.id=e.folder_id
        ORDER BY e.last_attempt_at DESC LIMIT 5''')]
    return {'database': str(path), 'messages': count,
            'historical_handles': conn.execute('SELECT count(*) FROM message_handles').fetchone()[0],
            'stores': conn.execute('SELECT count(*) FROM stores').fetchone()[0],
            'folders_by_state': {r['scan_state']: r['count'] for r in folders},
            'failed_items': errors, 'recent_failures': recent_failures,
            'size_mib': round(path.stat().st_size / 1048576, 1) if path.exists() else 0}
