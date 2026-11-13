"""Move-resistant handles: fake COM only; a real Windows smoke test is still required."""
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace

import pytest

from outlook_classic_mcp.db import UPSERT, open_db
from outlook_classic_mcp.open_message import open_indexed_message, lookup_reference, mail_matches
from outlook_classic_mcp.outlook import INTERNET_ID_DASL, SEARCH_KEY_DASL, mail_identity
from test_index import message, newdb


WHEN = datetime(2025, 12, 30, tzinfo=timezone.utc)
SENT = datetime(2025, 12, 29, tzinfo=timezone.utc)


class Mail:
    def __init__(self, parent, *, key=b'abc', message_id='<msg-123@example.net>',
                 subject='Phoenix renewal', sender='alice@example.test'):
        self.Parent = parent
        self.ReceivedTime = WHEN
        self.SentOn = SENT
        self.Subject = subject
        self.SenderEmailAddress = sender
        self.calls = []
        self.PropertyAccessor = SimpleNamespace(GetProperty=self.property)
        self.key = key
        self.message_id = message_id

    def property(self, tag):
        if tag == SEARCH_KEY_DASL:
            return self.key
        if tag == INTERNET_ID_DASL:
            return self.message_id
        raise KeyError(tag)

    def Display(self, modal):
        self.calls.append(('Display', modal))


class Folder:
    DefaultItemType = 0
    EntryID = 'f'
    FolderPath = '\\Inbox'
    Name = 'Inbox'

    def __init__(self, rows):
        self.rows = rows
        self.filters = []

    def GetTable(self, filter_text):
        self.filters.append(filter_text)
        folder = self
        class Table:
            def __init__(self):
                self.ids = list(folder.rows)
                self.position = 0
                self.Columns = SimpleNamespace(RemoveAll=lambda: None, Add=lambda name: None)
            @property
            def EndOfTable(self):
                return self.position >= len(self.ids)
            def GetNextRow(self):
                entry = self.ids[self.position]
                self.position += 1
                return {'EntryID': entry}
        return Table()


class Session:
    def __init__(self, ids, folders):
        self.ids = ids
        self._folders = folders
        self.ns = SimpleNamespace(GetItemFromID=self.get_item)
    def get_item(self, entry, store):
        return self.ids[(entry, store)]
    def stores(self, which):
        return [SimpleNamespace(StoreID='s')]
    def folders(self, store):
        return self._folders
    def __enter__(self):
        return self
    def __exit__(self, *args):
        return None


def patch_session(monkeypatch, session):
    monkeypatch.setattr('outlook_classic_mcp.open_message.OutlookSession', lambda: session)


def setup_old(db, *, keyed=True):
    old = message(entry_id='old', sent_at=None)
    if keyed:
        old.update(search_key='616263', internet_message_id='<msg-123@example.net>')
    with db:
        db.execute(UPSERT, old)


def test_pr_search_key_binary_and_string():
    a = Mail(Folder({}))
    assert mail_identity(a) == ('616263', '<msg-123@example.net>')
    a.key = '0x616263'
    assert mail_identity(a) == ('616263', '<msg-123@example.net>')


def test_move_before_sync_uses_filtered_table_and_keeps_numeric_id(tmp_path, monkeypatch):
    db, path = newdb(tmp_path)
    setup_old(db)
    db.close()
    folder = Folder({'new': 1})
    new = Mail(folder)
    patch_session(monkeypatch, Session({('new', 's'): new}, [folder]))
    assert 'relocated' in open_indexed_message(1, path)
    assert new.calls == [('Display', False)]
    assert any(INTERNET_ID_DASL in filt for filt in folder.filters)
    with open_db(path) as conn:
        assert conn.execute('SELECT entry_id FROM messages WHERE id=1').fetchone()[0] == 'new'
    assert 'Opened message' in open_indexed_message(1, path)
    assert new.calls == [('Display', False), ('Display', False)]


def test_move_after_reconcile_old_handle_redirects(tmp_path, monkeypatch):
    db, path = newdb(tmp_path)
    setup_old(db)
    with db:
        db.execute('DELETE FROM messages WHERE id=1')
        db.execute(UPSERT, message(entry_id='new', search_key='616263',
                                   internet_message_id='<msg-123@example.net>'))
    assert db.execute('SELECT id FROM messages').fetchone()[0] == 2
    assert lookup_reference(db, 1).entry_id is None
    db.close()
    folder = Folder({})
    new = Mail(folder)
    patch_session(monkeypatch, Session({('new', 's'): new}, [folder]))
    assert 'relocated' in open_indexed_message(1, path)
    assert new.calls == [('Display', False)]


def test_duplicate_copy_must_not_open_arbitrary_item(tmp_path, monkeypatch):
    db, path = newdb(tmp_path)
    setup_old(db)
    with db:
        db.execute(UPSERT, message(entry_id='new1', search_key='616263',
                                   internet_message_id='<msg-123@example.net>'))
        db.execute(UPSERT, message(entry_id='new2', search_key='616263',
                                   internet_message_id='<msg-123@example.net>'))
    db.close()
    folder = Folder({})
    a, b = Mail(folder), Mail(folder)
    patch_session(monkeypatch, Session({('new1','s'): a, ('new2','s'): b}, [folder]))
    with pytest.raises(LookupError, match='Multiple matching copies'):
        open_indexed_message(1, path)
    assert not a.calls and not b.calls


def test_legacy_v2_subject_and_date_fallback(tmp_path, monkeypatch):
    db, path = newdb(tmp_path)
    setup_old(db, keyed=False)
    db.close()
    folder = Folder({'new': 1})
    new = Mail(folder)
    patch_session(monkeypatch, Session({('new', 's'): new}, [folder]))
    assert 'relocated' in open_indexed_message(1, path)
    assert any('Phoenix renewal' in x for x in folder.filters)
    assert new.calls == [('Display', False)]


def test_different_message_does_not_open(tmp_path, monkeypatch):
    db, path = newdb(tmp_path)
    setup_old(db)
    db.close()
    folder = Folder({'wrong': 1})
    wrong = Mail(folder, key=b'other', message_id='<other@domain>')
    patch_session(monkeypatch, Session({('wrong', 's'): wrong}, [folder]))
    with pytest.raises(LookupError, match='Email not found'):
        open_indexed_message(1, path)
    assert not wrong.calls


def test_handles_survive_multiple_deletions_without_id_reuse(tmp_path):
    db, path = newdb(tmp_path)
    with db:
        db.execute(UPSERT, message(entry_id='first'))
        db.execute('DELETE FROM messages WHERE id=1')
        db.execute(UPSERT, message(entry_id='second'))
        db.execute('DELETE FROM messages WHERE id=2')
        db.execute(UPSERT, message(entry_id='third'))
    assert [r[0] for r in db.execute('SELECT message_id FROM message_handles ORDER BY message_id')] == [1,2]
    assert db.execute('SELECT id FROM messages').fetchone()[0] == 3
    db.close()


def test_get_mail_redirects_historical_id(tmp_path):
    from outlook_classic_mcp.db import get_message
    db, _ = newdb(tmp_path)
    setup_old(db)
    with db:
        db.execute('DELETE FROM messages WHERE id=1')
        db.execute(UPSERT, message(entry_id='new', search_key='616263',
                                   internet_message_id='<msg-123@example.net>'))
    result = get_message(db, 1)
    assert result['id'] == 2
    assert result['requested_message_id'] == 1
    db.close()


def test_resumable_backfill_reads_only_identity(tmp_path, monkeypatch):
    from outlook_classic_mcp.worker import backfill_identities
    import outlook_classic_mcp.worker as worker
    db, _ = newdb(tmp_path)
    with db:
        db.execute(UPSERT, message(1))
        db.execute(UPSERT, message(2))
    class FakeSession:
        def __enter__(self):
            self.ns = SimpleNamespace(GetItemFromID=lambda entry, store: Mail(Folder({})))
            return self
        def __exit__(self, *args):
            return None
    monkeypatch.setattr(worker, 'OutlookSession', FakeSession)
    assert backfill_identities(db, 1) == 0
    assert db.execute('SELECT cursor_id FROM identity_backfill_state').fetchone()[0] == 1
    assert backfill_identities(db, 1) == 0
    assert db.execute('SELECT cursor_id FROM identity_backfill_state').fetchone()[0] == 2
    assert all(r[0] == '616263' for r in db.execute('SELECT search_key FROM messages'))
    db.close()


def test_upgrade_v1_without_fts_rebuild(tmp_path):
    path = tmp_path / 'v1.sqlite3'
    c = sqlite3.connect(path)
    c.executescript('''
      CREATE TABLE stores (id INTEGER PRIMARY KEY, outlook_id TEXT UNIQUE, name TEXT);
      CREATE TABLE folders (id INTEGER PRIMARY KEY, store_id INTEGER, outlook_id TEXT,
        name TEXT, path TEXT, scan_state TEXT DEFAULT 'complete', scan_generation TEXT,
        scan_cursor TEXT, last_full_at TEXT, last_incremental_at TEXT, last_error TEXT,
        UNIQUE(store_id,outlook_id));
      CREATE TABLE messages (id INTEGER PRIMARY KEY,store_id INTEGER,folder_id INTEGER,
        entry_id TEXT,modified_at TEXT,received_at TEXT,sent_at TEXT,subject TEXT,
        sender_name TEXT,sender_email TEXT,recipients_to TEXT,recipients_cc TEXT,
        conversation_id TEXT,body TEXT,body_truncated INTEGER,is_unread INTEGER,
        has_attachments INTEGER,seen_generation TEXT, UNIQUE(store_id,entry_id));
      CREATE VIRTUAL TABLE mail_fts USING fts5(subject,sender_name,sender_email,
        recipients_to,recipients_cc,body,content='messages',content_rowid='id',
        tokenize='unicode61 remove_diacritics 2');
      INSERT INTO stores VALUES(1,'s','Old store');
      INSERT INTO folders(id,store_id,outlook_id,name,path) VALUES(1,1,'f','Inbox','\\Inbox');
      INSERT INTO messages(id,store_id,folder_id,entry_id,subject,sender_name,sender_email,
        recipients_to,recipients_cc,body,body_truncated,is_unread,has_attachments)
        VALUES(100,1,1,'oldentry','Old renewal','Alice','a@test','','','need renewal',0,0,0);
      INSERT INTO mail_fts(rowid,subject,sender_name,sender_email,recipients_to,
          recipients_cc,body) VALUES(100,'Old renewal','Alice','a@test','','','need renewal');
      PRAGMA user_version=1;
    ''')
    c.close()
    from outlook_classic_mcp.db import search
    with open_db(path) as db:
        assert db.execute('PRAGMA user_version').fetchone()[0] == 2
        assert db.execute('SELECT id FROM messages').fetchone()[0] == 100
        assert search(db, 'renewal')[0]['id'] == 100
        assert db.execute('SELECT count(*) FROM message_handles').fetchone()[0] == 0
