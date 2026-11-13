"""Cross-platform tests for new v2 features (real Windows COM requires smoke test)."""
import json
import subprocess
import sys
from types import SimpleNamespace

import pytest

from outlook_classic_mcp.db import UPSERT
from outlook_classic_mcp.integration import consume_sync_request, request_sync, start_open_message
from outlook_classic_mcp.open_message import lookup_location, open_indexed_message
from outlook_classic_mcp.worker import worker_lock

from test_index import message, newdb


def test_lookup_original_entry_and_store_and_missing(tmp_path):
    db, _ = newdb(tmp_path)
    with db:
        db.execute(UPSERT, message(42, entry_id='entry-42'))
    assert lookup_location(db, 1) == ('entry-42', 's')
    with pytest.raises(LookupError):
        lookup_location(db, 4444)
    with pytest.raises(ValueError):
        lookup_location(db, 0)
    db.close()


def test_com_opener_uses_store_and_displays_without_saving(tmp_path, monkeypatch):
    db, path = newdb(tmp_path)
    with db:
        db.execute(UPSERT, message(42, entry_id='outlook_entry_id'))
    db.close()
    actions = []
    item = SimpleNamespace(Display=lambda modal: actions.append(('display', modal)))

    class FakeNS:
        def GetItemFromID(self, entry, store):
            actions.append(('lookup', entry, store))
            return item

    class Session:
        def __enter__(self):
            return SimpleNamespace(ns=FakeNS())
        def __exit__(self, *_):
            return None
    monkeypatch.setattr('outlook_classic_mcp.open_message.OutlookSession', Session)
    assert 'Opened message' in open_indexed_message(1, path)
    assert actions == [('lookup', 'outlook_entry_id', 's'), ('display', False)]


def test_com_opener_stale_id_returns_helpful_error(tmp_path, monkeypatch):
    db, path = newdb(tmp_path)
    with db:
        db.execute(UPSERT, message())
    db.close()
    class Session:
        def __enter__(self):
            return SimpleNamespace(ns=SimpleNamespace(GetItemFromID=lambda *_: (_ for _ in ()).throw(RuntimeError('gone'))), stores=lambda _: [], folders=lambda _: [])
        def __exit__(self, *_):
            return None
    monkeypatch.setattr('outlook_classic_mcp.open_message.OutlookSession', Session)
    with pytest.raises(LookupError, match='Email not found'):
        open_indexed_message(1, path)


def test_mcp_helper_process_invocation_is_safe(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(sys, 'platform', 'win32')
    def run(cmd, **kwargs):
        calls.append((cmd, kwargs))
        return subprocess.CompletedProcess(cmd, 0, json.dumps({'ok': True, 'message': 'Opened'}), '')
    monkeypatch.setattr(subprocess, 'run', run)
    assert start_open_message(42, tmp_path / 'some data.sqlite3') == 'Opened'
    assert calls[0][0][1:4] == ['-m', 'outlook_classic_mcp.open_message', '42']
    assert calls[0][1]['timeout'] == 45
    assert calls[0][1]['capture_output'] is True
    with pytest.raises(ValueError):
        start_open_message(0)
    with pytest.raises(ValueError):
        start_open_message(True)


def test_mcp_helper_failure_and_timeout(monkeypatch):
    monkeypatch.setattr(sys, 'platform', 'win32')
    monkeypatch.setattr(subprocess, 'run', lambda cmd, **kwargs: subprocess.CompletedProcess(
        cmd, 1, json.dumps({'ok': False, 'error': 'stale item'}), ''))
    with pytest.raises(RuntimeError, match='stale item'):
        start_open_message(2)
    def timeout(cmd, **kwargs):
        raise subprocess.TimeoutExpired(cmd, 45)
    monkeypatch.setattr(subprocess, 'run', timeout)
    with pytest.raises(RuntimeError, match='45 seconds'):
        start_open_message(2)


def test_sync_request_and_single_worker(tmp_path):
    db = tmp_path / 'folder' / 'index.sqlite3'
    req = request_sync(db)
    assert req.exists()
    assert consume_sync_request(db)
    assert not consume_sync_request(db)
    with worker_lock(db):
        with pytest.raises(RuntimeError, match='Another index worker'):
            with worker_lock(db):
                pass
    with worker_lock(db):
        pass


def test_reconcile_due_does_not_restart_recent_completed_folders(tmp_path, monkeypatch):
    from outlook_classic_mcp import worker
    from outlook_classic_mcp.db import now_utc

    db, _ = newdb(tmp_path)
    db.execute('UPDATE folders SET scan_state="complete", last_full_at=? WHERE id=1', (now_utc(),))
    db.commit()
    folder = SimpleNamespace(EntryID='f', FolderPath='\\Inbox')
    store = SimpleNamespace(StoreID='s', DisplayName='Test Store')
    class Session:
        def __enter__(self):
            return self
        def __exit__(self, *_):
            return None
        def stores(self, which):
            return [store]
        def folders(self, unused):
            return [folder]
    calls = []
    monkeypatch.setattr(worker, 'OutlookSession', Session)
    monkeypatch.setattr(worker, 'index_folder', lambda *a, **kw: (calls.append(1) or (0, 0, 0)))
    options = dict(refresh=True, batch_size=10, max_body_chars=1000,
                   lookback_days=30, reconcile_due_days=7)
    assert worker.run_index(db, 'full', 'all', **options) == 0
    assert not calls
    db.execute("UPDATE folders SET last_full_at='2020-01-01T00:00:00Z'")
    db.commit()
    assert worker.run_index(db, 'full', 'all', **options) == 0
    assert calls == [1]
    db.close()


def test_pause_when_outlook_closed_never_enters_com(tmp_path, monkeypatch):
    from outlook_classic_mcp import worker
    db, _ = newdb(tmp_path)
    monkeypatch.setattr(worker, 'is_outlook_running', lambda: False)
    class NoSession:
        def __init__(self):
            raise AssertionError('COM must not be launched when Outlook is closed')
    monkeypatch.setattr(worker, 'OutlookSession', NoSession)
    with pytest.raises(worker.OutlookClosedError):
        worker.run_index(db, 'incremental', 'all', require_outlook_running=True)
    db.close()
