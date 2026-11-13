"""OS integration for local Outlook opening and tray background sync.

No network calls or message modification. stdio MCP uses this module to launch
an independent foreground Outlook COM helper.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

from .db import default_db_path


def start_open_message(message_id: int, db_path: Path | None = None, timeout: int = 45) -> str:
    """Launch a short-lived COM helper; only a numeric indexed message ID is accepted."""
    if not isinstance(message_id, int) or isinstance(message_id, bool) or message_id < 1:
        raise ValueError('message_id must be a positive integer')
    if sys.platform != 'win32':
        raise RuntimeError('Opening Outlook requires Windows and Outlook Classic.')
    python = Path(sys.executable)
    if python.name.lower() == 'pythonw.exe':
        python = python.with_name('python.exe')
    cmd = [str(python), '-m', 'outlook_classic_mcp.open_message', str(message_id),
           '--db', str(db_path or default_db_path())]
    flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8',
                                errors='replace', timeout=timeout, creationflags=flags,
                                check=False)
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError('Outlook did not respond within 45 seconds.') from exc
    try:
        payload = json.loads(result.stdout.strip())
    except (ValueError, TypeError) as exc:
        raise RuntimeError(f'Outlook helper failed (exit {result.returncode}). '
                           'Check that Outlook Classic is configured and running.') from exc
    if result.returncode or not payload.get('ok'):
        raise RuntimeError(payload.get('error', 'Outlook helper could not open the message.'))
    return payload.get('message', 'Opened in Outlook.')


def request_sync(db_path: Path | None = None) -> Path:
    """Ask an existing watch process to run its next pass now."""
    db = (db_path or default_db_path()).resolve()
    path = db.with_name(db.name + '.sync-request')
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text('1', encoding='ascii')
    return path


def consume_sync_request(db_path: Path) -> bool:
    path = db_path.with_name(db_path.name + '.sync-request')
    try:
        path.unlink()
        return True
    except FileNotFoundError:
        return False
