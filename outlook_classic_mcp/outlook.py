"""Classic Outlook adapter. Import win32com ONLY inside a Windows foreground worker.

No Outlook COM calls are exposed through the MCP process.
"""
from __future__ import annotations

import time
from datetime import datetime, timezone
from typing import Any, Iterator


MAIL_CLASS_PREFIX = 'IPM.Note'
# PR_LAST_MODIFICATION_TIME: documented DASL UTC date comparison property.
MODIFIED_DASL = 'http://schemas.microsoft.com/mapi/proptag/0x30080040'


# Outlook MAPI properties. Unlike EntryID, PR_SEARCH_KEY normally survives moves.
# It is NOT guaranteed unique: copied mail can share the same search key.
SEARCH_KEY_DASL = 'http://schemas.microsoft.com/mapi/proptag/0x300B0102'
INTERNET_ID_DASL = 'http://schemas.microsoft.com/mapi/proptag/0x1035001F'
SUBJECT_DASL = 'http://schemas.microsoft.com/mapi/proptag/0x0037001F'


def _mapi_property(item: Any, name: str) -> Any:
    try:
        return item.PropertyAccessor.GetProperty(name)
    except Exception:
        return None


def mail_identity(item: Any) -> tuple[str | None, str | None]:
    """Read-only move-resistant hints; no changes to Outlook items."""
    key = _mapi_property(item, SEARCH_KEY_DASL)
    search_key = None
    if key is not None:
        try:
            if isinstance(key, str):
                # Some COM versions return hexadecimal strings.
                key = key.removeprefix('0x').replace(' ', '')
                if key and all(c in '0123456789abcdefABCDEF' for c in key):
                    search_key = key.upper()
            elif isinstance(key, (bytes, bytearray, memoryview)):
                search_key = bytes(key).hex().upper()
            else:
                search_key = bytes(key).hex().upper()
        except (ValueError, TypeError, OverflowError):
            search_key = None
    raw_id = _mapi_property(item, INTERNET_ID_DASL)
    internet_id = str(raw_id).strip()[:1024] if raw_id else None
    return search_key or None, internet_id or None


def normalize_time(value: Any) -> str | None:
    """UTC ISO 8601, or None. Outlook OOM typically presents datetime in local time.

    Naive datetimes are treated as machine-local; historical DST is resolved by Python
    using the local Windows time-zone conversion. The filter overlaps by 48 hours.
    """
    if not isinstance(value, datetime) or not 1900 <= value.year <= 2100:
        return None
    try:
        if value.tzinfo is None:
            value = value.astimezone()
        return value.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    except (OverflowError, OSError, ValueError):
        return None


def filter_modified(*, older_than_utc: datetime | None = None,
                    newer_than_utc: datetime | None = None) -> str:
    """DASL uses UTC. Outlook date literals have no seconds (Microsoft documentation)."""
    if (older_than_utc is None) == (newer_than_utc is None):
        raise ValueError('Supply exactly one cutoff')
    cutoff = older_than_utc if older_than_utc is not None else newer_than_utc
    assert cutoff is not None
    cutoff = cutoff.astimezone(timezone.utc)
    # Use explicit 12h US-style Outlook string. A few regional builds need alternate
    # formatting: --scan-all-metadata bypasses this filter entirely.
    literal = cutoff.strftime('%m/%d/%Y %I:%M %p')
    operator = '<=' if older_than_utc is not None else '>='
    return f'@SQL="{MODIFIED_DASL}" {operator} \'{literal}\''


def chunks_from_table(table: Any, count: int) -> Iterator[list[tuple[Any, Any, Any]]]:
    """Stream GetArray batches. Expected SAFEARRAY shape is [column][row].

    Columns are explicitly EntryID, LastModificationTime, MessageClass.
    Keep each transfer bounded, including on million-item folders.
    """
    while not table.EndOfTable:
        block = table.GetArray(count)
        if not block:
            return
        if not all(isinstance(col, (tuple, list)) for col in block):
            raise RuntimeError('Unexpected Outlook GetArray data: expected a 2D array')
        if len(block) == 3:
            # Microsoft documents column-major SAFEARRAY: [column][row].
            if not (len(block[0]) == len(block[1]) == len(block[2])):
                raise RuntimeError('Outlook GetArray columns have mismatched sizes')
            yield list(zip(block[0], block[1], block[2]))
        elif len(block[0]) == 3:
            # Handle pywin32 variants materialized as [row][column] too.
            yield [tuple(row) for row in block]
        else:
            raise RuntimeError('Unexpected Outlook GetArray shape; expected 3 metadata columns')


def open_table(folder: Any, cutoff: datetime | None, mode: str, no_filter: bool = False) -> Any:
    """Filter and sort server-side, with a correctness-preserving unfiltered fallback."""
    flt = '' if no_filter or cutoff is None else filter_modified(
        older_than_utc=cutoff if mode == 'full' else None,
        newer_than_utc=cutoff if mode == 'incremental' else None)
    try:
        table = folder.GetTable(flt)
    except Exception as exc:
        if not flt:
            raise
        print(f'WARNING: Outlook rejected date filter ({exc}). Scanning all metadata instead.', flush=True)
        table = folder.GetTable()
    table.Columns.RemoveAll()
    for column in ('EntryID', 'LastModificationTime', 'MessageClass'):
        table.Columns.Add(column)
    if mode == 'full':
        table.Sort('LastModificationTime', True)
    return table


def read_string(item: Any, name: str) -> str:
    value = getattr(item, name, '')
    return str(value) if value is not None else ''


def read_mail(item: Any, store_id: int, folder_id: int, entry_id: str,
              modified_at: str | None, seen_generation: str | None,
              max_body_chars: int) -> dict[str, Any]:
    """Read only. No MailItem.Save/Move/Delete/Send calls."""
    body = read_string(item, 'Body')
    short_body = body[:max_body_chars]
    search_key, internet_id = mail_identity(item)
    return {
        'store_id': store_id, 'folder_id': folder_id, 'entry_id': entry_id,
        'search_key': search_key, 'internet_message_id': internet_id,
        'modified_at': modified_at or normalize_time(getattr(item, 'LastModificationTime', None)),
        'received_at': normalize_time(getattr(item, 'ReceivedTime', None)),
        'sent_at': normalize_time(getattr(item, 'SentOn', None)),
        'subject': read_string(item, 'Subject'),
        'sender_name': read_string(item, 'SenderName'),
        'sender_email': read_string(item, 'SenderEmailAddress'),
        'recipients_to': read_string(item, 'To'),
        'recipients_cc': read_string(item, 'CC'),
        'conversation_id': read_string(item, 'ConversationID'),
        'body': short_body, 'body_truncated': int(len(body) > max_body_chars),
        'is_unread': int(bool(getattr(item, 'UnRead', False))),
        'has_attachments': int(bool(getattr(item.Attachments, 'Count', 0))),
        'seen_generation': seen_generation,
    }


class OutlookSession:
    """COM must be initialized/uninitialized in the same OS thread."""

    def __enter__(self):
        import pythoncom
        from win32com.client import Dispatch

        self.pythoncom = pythoncom
        pythoncom.CoInitialize()
        try:
            self.application = Dispatch('Outlook.Application')
            self.ns = self.application.GetNamespace('MAPI')
        except BaseException:
            pythoncom.CoUninitialize()
            raise
        return self

    def __exit__(self, exc_type, exc, tb):
        self.application = None
        self.ns = None
        self.pythoncom.CoUninitialize()

    def stores(self, which: str = 'default') -> Iterator[Any]:
        if which == 'default':
            yield self.ns.DefaultStore
        elif which == 'all':
            # Copy only COM references; no full messages are loaded here.
            for index in range(1, self.ns.Stores.Count + 1):
                yield self.ns.Stores.Item(index)
        else:
            raise ValueError('stores must be default or all')

    @staticmethod
    def folders(store: Any) -> Iterator[Any]:
        stack = [store.GetRootFolder()]
        while stack:
            folder = stack.pop()
            try:
                children = folder.Folders
                for index in range(children.Count, 0, -1):
                    stack.append(children.Item(index))
            except Exception as exc:
                print(f'WARNING: Cannot enumerate subfolders of {folder.Name}: {exc}', flush=True)
            # Includes any folder that can contain messages, including a mail subfolder
            # under a nonmail parent. Virtual search folders should be excluded manually.
            yield folder
