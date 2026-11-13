"""Local SQLite search MCP + explicit UI-only Outlook opener for Goose Desktop.

No network listeners, no mail mutation, no indexing triggered from an agent.
Opening Outlook runs in a short-lived COM subprocess, not in the MCP process.
"""
from __future__ import annotations

import sqlite3
from contextlib import closing
from typing import Any

from mcp.server import MCPServer

from . import __version__
from .db import default_db_path, get_message, open_db, search, status
from .integration import start_open_message

mcp = MCPServer('Outlook Classic Search', version=__version__,
                instructions='Offline search of Outlook Classic mail, plus optional local display. '
                'Use search_mail to find candidates, then get_mail for details. '
                'If the user asks to view an original email in Outlook, call open_mail_in_outlook. '
                'Search coverage depends on index worker progress and sync frequency.')


@mcp.tool()
def search_mail(query: str, mode: str = 'all', sender: str = '', folder: str = '',
                after: str = '', before: str = '', limit: int = 20,
                offset: int = 0) -> list[dict[str, Any]]:
    """Ranked full-text mailbox search; subject weighted above sender and body.

    query: words to find (e.g. 'invoice march 2024').
    mode: 'all' (every word), 'any', 'phrase' (contiguous words),
          or 'fts' (advanced SQLite FTS5 e.g. '"contract renewal" OR NDA',
          'subject:budget', 'NEAR(revenue forecast, 3)', 'alpha*').
    sender: case-insensitive partial sender name/address.
    folder: partial Outlook folder path, including store name if needed.
    after/before: YYYY-MM-DD dates. After inclusive, before exclusive.
    limit: maximum 100. offset: up to 10000.
    Returns index IDs for get_mail, not Outlook COM EntryIDs.
    """
    with closing(open_db(readonly=True)) as conn:
        return search(conn, query, mode, sender, folder,
                      after or None, before or None, limit, offset)


@mcp.tool()
def get_mail(message_id: int, max_body_chars: int = 20000) -> dict[str, Any]:
    """Retrieve a message by integer index ID returned by search_mail.

    Body is plain text and may be truncated at indexing time (default 64k chars).
    max_body_chars caps response size (maximum 200k). Attachments are NOT copied.
    This never touches or mutates Outlook.
    """
    with closing(open_db(readonly=True)) as conn:
        message = get_message(conn, message_id, max_body_chars)
    if message is None:
        return {'error': f'No indexed message with ID {message_id}; it may have been deleted or has multiple matches.'}
    return message


@mcp.tool()
def open_mail_in_outlook(message_id: int) -> dict[str, Any]:
    """Open a specific search result in Outlook Classic, visibly on the user's PC.

    message_id is the INTEGER `id` from search_mail/get_mail; never guess it.
    Launches Outlook Classic if closed. No message is edited, sent, or deleted.
    This is a local GUI action; only call when the user explicitly wants to open
    the original email in Outlook. Moved messages use a validated fallback lookup;
    genuinely deleted messages cannot be opened.
    """
    try:
        return {'ok': True, 'message': start_open_message(message_id)}
    except (ValueError, RuntimeError, FileNotFoundError) as exc:
        return {'ok': False, 'error': str(exc)}


@mcp.tool()
def index_status() -> dict[str, Any]:
    """Return number of indexed messages, incomplete folders, and indexing errors."""
    path = default_db_path()
    with closing(open_db(path, readonly=True)) as conn:
        return status(conn, path)


@mcp.tool()
def list_mail_folders(limit: int = 100, search: str = '') -> list[dict[str, Any]]:
    """List folder paths and indexed message counts. Filter paths by substring."""
    if not 1 <= limit <= 300:
        raise ValueError('limit must be 1..300')
    with closing(open_db(readonly=True)) as conn:
        return [dict(r) for r in conn.execute('''SELECT f.path, f.scan_state,
            f.last_full_at, f.last_incremental_at, count(m.id) AS messages
            FROM folders f LEFT JOIN messages m ON m.folder_id=f.id
            WHERE instr(lower(f.path),lower(?)) > 0
            GROUP BY f.id ORDER BY messages DESC LIMIT ?''', (search, limit))]


if __name__ == '__main__':
    mcp.run(transport='stdio')
