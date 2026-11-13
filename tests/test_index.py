from __future__ import annotations

from datetime import datetime, timedelta, timezone
from types import SimpleNamespace

import pytest

from outlook_classic_mcp.db import UPSERT, fts_query, get_message, open_db, search, status
from outlook_classic_mcp.outlook import chunks_from_table
from outlook_classic_mcp.worker import index_folder


def message(i=1, text='Project Phoenix agreement signed', **changes):
    base = dict(store_id=1, folder_id=1, entry_id='e-' + str(i),
                search_key=None, internet_message_id=None,
                modified_at='2026-01-01T00:00:00Z', received_at='2025-12-30T00:00:00Z',
                sent_at=None, subject='Phoenix renewal', sender_name='Alice',
                sender_email='alice@example.test', recipients_to='bob@example.test',
                recipients_cc='', conversation_id='conversation', body=text,
                body_truncated=0, is_unread=0, has_attachments=0, seen_generation='gen')
    base.update(changes)
    return base


def newdb(tmp_path):
    path = tmp_path / 'test.sqlite3'
    conn = open_db(path)
    conn.execute('INSERT INTO stores(id,outlook_id,name) VALUES(1,\'s\',\'Test Store\')')
    conn.execute('INSERT INTO folders(id,store_id,outlook_id,name,path) VALUES(1,1,\'f\',\'Inbox\',\'\\\\Inbox\')')
    conn.commit()
    return conn, path


def test_fts_search_and_triggers(tmp_path):
    db, path = newdb(tmp_path)
    with db:
        db.execute(UPSERT, message())
        db.execute(UPSERT, message(2, text='Quarterly budget and board review', subject='Budget'))
    result = search(db, 'phoenix agreement')
    assert len(result) == 1 and result[0]['id'] == 1
    assert len(search(db, 'phoenix board', mode='any')) == 2
    assert len(search(db, 'phoenix agreement', mode='phrase')) == 1
    assert len(search(db, 'subject:budget', mode='fts')) == 1
    assert len(search(db, 'phoenix', sender='alice@example', after='2025-12-01')) == 1
    assert not search(db, 'phoenix', before='2025-01-01')
    assert len(search(db, 'phoenix', limit=1)) == 1
    assert not search(db, 'phoenix', offset=1)
    assert get_message(db, 1, 10)['body_response_truncated'] is True
    assert status(db, path)['messages'] == 2
    with db:
        db.execute(UPSERT, message(text='Entirely different contents', subject='Changed'))
    assert not search(db, 'phoenix')
    assert search(db, 'entirely')
    with db:
        db.execute('DELETE FROM messages WHERE id=1')
    assert not search(db, 'entirely')
    db.close()


def test_search_safe_and_date_validation(tmp_path):
    db, _ = newdb(tmp_path)
    with pytest.raises(ValueError):
        search(db, '   ')
    with pytest.raises(ValueError):
        search(db, 'test', mode='fail')
    with pytest.raises(ValueError):
        search(db, 'test', mode='fts', limit=101)
    with pytest.raises(ValueError):
        search(db, 'test', after='2026-2-01')
    with pytest.raises(ValueError):
        search(db, 'test', mode='fts', offset=10001)
    assert fts_query('"bad" syntax:stuff') == '"bad" AND "syntax" AND "stuff"'
    db.close()


class FakeTable:
    def __init__(self, rows):
        self.rows = rows[:]
        self.pos = 0
        self.Columns = SimpleNamespace(RemoveAll=lambda: None, Add=lambda name: None)

    @property
    def EndOfTable(self):
        return self.pos >= len(self.rows)

    def Sort(self, key, descending):
        self.rows.sort(key=lambda r: r[1], reverse=descending)

    def GetArray(self, count):
        rows = self.rows[self.pos:self.pos + count]
        self.pos += len(rows)
        return tuple(tuple(row[i] for row in rows) for i in range(3))


class FakeMail:
    def __init__(self, text):
        self.Body = text
        self.Subject = 'mail ' + text
        self.SenderName = 'Alice'
        self.SenderEmailAddress = 'a@test.test'
        self.To = 'Bob'
        self.CC = ''
        self.ConversationID = 'conversation'
        self.ReceivedTime = datetime(2026, 1, 1, tzinfo=timezone.utc)
        self.SentOn = datetime(2025, 12, 31, tzinfo=timezone.utc)
        self.LastModificationTime = datetime(2026, 1, 2, tzinfo=timezone.utc)
        self.UnRead = False
        self.Attachments = SimpleNamespace(Count=0)


class FakeFolder:
    EntryID = 'folder-1'
    Name = 'Inbox'
    FolderPath = '\\\\Inbox'
    DefaultItemType = 0

    def __init__(self, entries):
        self.entries = entries

    def GetTable(self, filter=''):
        # Fake Outlook doesn't parse DASL; all rows are returned so we can
        # test the full scans and metadata comparisons.
        return FakeTable([(entry, mod, 'IPM.Note') for entry, (mod, text) in self.entries.items()])


class FakeNS:
    def __init__(self, folder):
        self.folder = folder
        self.reads = 0

    def GetItemFromID(self, entry, storeid):
        self.reads += 1
        return FakeMail(self.folder.entries[entry][1])


def test_folder_indexing_idempotency_and_pruning(tmp_path):
    db, _ = newdb(tmp_path)
    db.execute('DELETE FROM folders')
    db.execute('DELETE FROM stores')
    db.execute('INSERT INTO stores(id,outlook_id,name) VALUES(1,\'store-1\',\'Store\')')
    db.commit()
    date = datetime(2026, 1, 2, tzinfo=timezone.utc)
    folder = FakeFolder({'a': (date, 'hello phoenix'), 'b': (date, 'hello budget')})
    ns = FakeNS(folder)
    session = SimpleNamespace(ns=ns)

    def full(refresh=False):
        return index_folder(db, session, folder, 1, 'store-1', 'full', refresh, 10, 65536, 30, True)

    assert full()[0] == 2
    assert ns.reads == 2
    assert full()[0] == 0  # complete folders skipped during bootstrap
    assert full(refresh=True)[0] == 0
    assert ns.reads == 2  # revisiting metadata doesn't fetch unchanged bodies
    assert len(search(db, 'phoenix')) == 1

    folder.entries['a'] = (date + timedelta(days=1), 'hello contract')
    assert full(refresh=True)[0] == 1
    assert not search(db, 'phoenix')
    assert len(search(db, 'contract')) == 1

    del folder.entries['b']
    full(refresh=True)
    assert db.execute('SELECT count(*) FROM messages').fetchone()[0] == 1
    assert not search(db, 'budget')
    db.close()


def test_table_batches():
    when = datetime(2026, 1, 1, tzinfo=timezone.utc)
    table = FakeTable([('id1', when, 'IPM.Note'), ('id2', when, 'IPM.Note')])
    blocks = list(chunks_from_table(table, 1))
    assert blocks == [[('id1', when, 'IPM.Note')], [('id2', when, 'IPM.Note')]]


class CrashTable(FakeTable):
    def GetArray(self, count):
        if self.pos > 0:
            raise RuntimeError('simulated Outlook disconnect')
        return super().GetArray(count)


class CrashOnceFolder(FakeFolder):
    def __init__(self, entries):
        super().__init__(entries)
        self.first = True

    def GetTable(self, filter=''):
        rows = [(entry, mod, 'IPM.Note') for entry, (mod, text) in self.entries.items()]
        if self.first:
            self.first = False
            return CrashTable(rows)
        return FakeTable(rows)


def test_resume_after_mid_folder_failure(tmp_path):
    db, _ = newdb(tmp_path)
    date = datetime(2026, 1, 10, tzinfo=timezone.utc)
    folder = CrashOnceFolder({'new': (date, 'new contract'),
                              'old': (date - timedelta(days=10), 'old contract')})
    ns = FakeNS(folder)
    session = SimpleNamespace(ns=ns)
    args = (db, session, folder, 1, 's', 'full', False, 1, 65536, 30, True)
    with pytest.raises(RuntimeError, match='disconnect'):
        index_folder(*args)
    assert db.execute("SELECT scan_state FROM folders WHERE outlook_id='folder-1'").fetchone()[0] == 'running'
    assert db.execute('SELECT count(*) FROM messages').fetchone()[0] == 1
    index_folder(*args)
    assert db.execute('SELECT count(*) FROM messages').fetchone()[0] == 2
    assert len(search(db, 'old')) == 1
    db.close()


def test_large_prune_safety(tmp_path):
    db, _ = newdb(tmp_path)
    with db:
        db.executemany(UPSERT, (message(i, text='old mail') for i in range(510)))
    folder = FakeFolder({})
    ns = FakeNS(folder)
    session = SimpleNamespace(ns=ns)
    args = (db, session, folder, 1, 's', 'full', True, 10, 65536, 30, True)
    # Different Folder.EntryID means new folder, so ensure existing location matches.
    db.execute("UPDATE folders SET outlook_id='folder-1' WHERE id=1")
    db.commit()
    with pytest.raises(RuntimeError, match='Prune safety stop'):
        index_folder(*args)
    assert db.execute('SELECT count(*) FROM messages').fetchone()[0] == 510
    index_folder(*args, allow_large_prune=True)
    assert db.execute('SELECT count(*) FROM messages').fetchone()[0] == 0
    db.close()


def test_row_major_getarray_compatibility():
    when = datetime(2026, 1, 1, tzinfo=timezone.utc)
    class RowMajorTable(FakeTable):
        def GetArray(self, count):
            rows = self.rows[self.pos:self.pos + count]
            self.pos += len(rows)
            return tuple(rows)
    table = RowMajorTable([('a', when, 'IPM.Note')] * 5)
    assert len(list(chunks_from_table(table, 5))[0]) == 5
