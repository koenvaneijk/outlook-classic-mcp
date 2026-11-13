"""Windows tray search and resilient background Outlook synchronizer.

Run with pythonw.exe -m outlook_classic_mcp.tray for no console. The UI and MCP access the
existing SQLite DB. A separate worker subprocess owns all indexing COM calls.
"""
from __future__ import annotations

import argparse
import os
import queue
import subprocess
import sys
import threading
from contextlib import closing
from pathlib import Path
from typing import Any

from .db import default_db_path, get_message, open_db, search, status
from .integration import request_sync, start_open_message


def startup_path() -> Path:
    appdata = Path(os.environ.get('APPDATA', str(Path.home() / 'AppData' / 'Roaming')))
    return appdata / 'Microsoft' / 'Windows' / 'Start Menu' / 'Programs' / 'Startup' / 'Outlook Classic Search.lnk'


def startup_enabled() -> bool:
    return startup_path().is_file()


def set_startup(enabled: bool, db_path: Path) -> None:
    if not enabled:
        startup_path().unlink(missing_ok=True)
        return
    import win32com.client
    shortcut_file = startup_path()
    shortcut_file.parent.mkdir(parents=True, exist_ok=True)
    exe = Path(sys.executable)
    pythonw = exe.with_name('pythonw.exe')
    if not pythonw.exists():
        raise RuntimeError('Cannot find pythonw.exe in your virtual environment.')
    shell = win32com.client.Dispatch('WScript.Shell')
    shortcut = shell.CreateShortcut(str(shortcut_file))
    shortcut.TargetPath = str(pythonw)
    shortcut.Arguments = f'-m outlook_classic_mcp.tray --db "{db_path}"'
    shortcut.WorkingDirectory = str(exe.parent.parent)
    shortcut.Description = 'Search Outlook locally and keep SQLite index synchronized'
    shortcut.Save()


def create_icon() -> Any:
    from PIL import Image, ImageDraw
    img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((6, 11, 58, 51), radius=6, fill=(28, 98, 174, 255))
    draw.line([(10, 17), (32, 34), (54, 17)], fill='white', width=4, joint='curve')
    return img


class TrayApp:
    def __init__(self, db_path: Path, every_minutes: int = 15, show_at_start: bool = False):
        import tkinter as tk
        from tkinter import ttk
        import pystray

        self.tk, self.ttk, self.pystray = tk, ttk, pystray
        self.db_path = db_path.resolve()
        self.every_minutes = every_minutes
        self.jobs: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.proc: subprocess.Popen | None = None
        self.closed = False
        self.retry_seconds = 10
        self.next_worker_check = 0.0
        self.search_seq = 0

        self.root = tk.Tk()
        self.root.title('Outlook Classic Search')
        self.root.geometry('1150x700')
        self.root.minsize(760, 430)
        self.root.protocol('WM_DELETE_WINDOW', self.hide)

        top = ttk.Frame(self.root, padding=10)
        top.pack(fill='x')
        self.query = tk.StringVar()
        ttk.Label(top, text='Search:').pack(side='left')
        entry = ttk.Entry(top, textvariable=self.query)
        entry.pack(side='left', fill='x', expand=True, padx=(8, 8))
        entry.bind('<Return>', lambda event: self.do_search())
        self.entry = entry
        self.mode = tk.StringVar(value='all')
        ttk.Combobox(top, textvariable=self.mode, state='readonly', width=9,
                     values=['all', 'any', 'phrase', 'fts']).pack(side='left', padx=5)
        ttk.Button(top, text='Search', command=self.do_search).pack(side='left', padx=5)
        ttk.Button(top, text='Sync now', command=self.sync_now).pack(side='left', padx=5)

        info = ttk.Frame(self.root, padding=(10, 0, 10, 8))
        info.pack(fill='x')
        self.status_text = tk.StringVar(value='Starting background synchronization…')
        ttk.Label(info, textvariable=self.status_text).pack(side='left')
        ttk.Label(info, text='Double-click a result to open in Outlook').pack(side='right')

        area = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        area.pack(fill='both', expand=True)
        results_frame = ttk.Frame(area)
        results_frame.pack(fill='both', expand=True)
        cols = ('date', 'sender', 'subject', 'folder')
        tree = ttk.Treeview(results_frame, columns=cols, show='headings', selectmode='browse')
        for col, title, width in [('date', 'Date', 105), ('sender', 'Sender', 205),
                                  ('subject', 'Subject', 440), ('folder', 'Folder', 250)]:
            tree.heading(col, text=title)
            tree.column(col, width=width, minwidth=80,
                        stretch=col in ('subject', 'folder'))
        bar = ttk.Scrollbar(results_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=bar.set)
        tree.pack(side='left', fill='both', expand=True)
        bar.pack(side='right', fill='y')
        tree.bind('<<TreeviewSelect>>', lambda _: self.preview())
        tree.bind('<Double-1>', lambda _: self.open_selected())
        tree.bind('<Return>', lambda _: self.open_selected())
        self.tree = tree

        preview_frame = ttk.LabelFrame(area, text='Preview — plain text, not live HTML')
        preview_frame.pack(fill='both', expand=True, pady=(10, 0))
        preview = tk.Text(preview_frame, height=12, wrap='word', state='disabled')
        preview.pack(fill='both', expand=True, padx=5, pady=5)
        self.preview_box = preview
        if not show_at_start:
            self.root.withdraw()
        else:
            self.root.after(100, self.focus_search)

        Menu = pystray.Menu
        Item = pystray.MenuItem
        self.icon = pystray.Icon(
            'OutlookClassicSearch', create_icon(), 'Outlook Classic Search',
            menu=Menu(
                Item('Search emails', lambda icon, item: self.jobs.put(('show', None)), default=True),
                Item('Sync now', lambda icon, item: self.jobs.put(('sync', None))),
                Item('Start with Windows', lambda icon, item: self.jobs.put(('startup', None)),
                     checked=lambda _: startup_enabled()),
                Item('Open sync log', lambda icon, item: self.jobs.put(('log', None))),
                Item('Quit (stop background sync)', lambda icon, item: self.jobs.put(('quit', None))),
            ),
        )

    def show(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.focus_search()

    def focus_search(self) -> None:
        self.entry.focus_set()

    def hide(self) -> None:
        self.root.withdraw()

    def do_search(self) -> None:
        query = self.query.get().strip()
        if not query:
            self.status_text.set('Enter search words.')
            return
        self.search_seq += 1
        sequence = self.search_seq
        mode = self.mode.get()
        self.status_text.set('Searching SQLite…')
        def work() -> None:
            try:
                with closing(open_db(self.db_path, readonly=True)) as conn:
                    rows = search(conn, query, mode=mode, limit=100)
                self.jobs.put(('search_result', (sequence, rows)))
            except Exception as exc:
                self.jobs.put(('search_error', (sequence, str(exc))))
        threading.Thread(target=work, name='mail-search', daemon=True).start()

    def show_preview(self, message: str) -> None:
        self.preview_box.configure(state='normal')
        self.preview_box.delete('1.0', 'end')
        self.preview_box.insert('end', message)
        self.preview_box.configure(state='disabled')

    def preview(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        mid = int(selection[0])
        try:
            with closing(open_db(self.db_path, readonly=True)) as conn:
                m = get_message(conn, mid, max_body_chars=12000)
            if m is None:
                self.show_preview('Message no longer in index.')
            else:
                self.show_preview(f"From: {m['sender_name']} <{m['sender_email']}>\n"
                                  f"To: {m['recipients_to']}\nSubject: {m['subject']}\n\n"
                                  f"{m['body']}")
        except Exception as exc:
            self.show_preview(f'Preview failed: {exc}')

    def open_selected(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        mid = int(selection[0])
        self.status_text.set(f'Opening message {mid} in Outlook…')
        def work() -> None:
            try:
                answer = start_open_message(mid, self.db_path)
                self.jobs.put(('notice', answer))
            except Exception as exc:
                self.jobs.put(('error', str(exc)))
        threading.Thread(target=work, name='open-original', daemon=True).start()

    def sync_now(self) -> None:
        request_sync(self.db_path)
        self.status_text.set('Sync requested. Will run when Outlook Classic is available.')

    def _spawn_worker(self) -> None:
        if self.proc is not None and self.proc.poll() is None:
            return
        log = self.db_path.parent / 'sync.log'
        log.parent.mkdir(parents=True, exist_ok=True)
        # Limit local log growth. Logs contain folder paths, not email contents.
        if log.exists() and log.stat().st_size > 10_000_000:
            log.replace(log.with_suffix('.log.1'))
        exe = Path(sys.executable)
        if exe.name.lower() == 'pythonw.exe':
            exe = exe.with_name('python.exe')
        command = [str(exe), '-m', 'outlook_classic_mcp.worker', 'watch', '--stores', 'all',
                   '--db', str(self.db_path), '--every-minutes', str(self.every_minutes),
                   '--only-if-outlook-running', '--reconcile-every-days', '7']
        with open(log, 'ab', buffering=0) as stream:
            self.proc = subprocess.Popen(command, stdin=subprocess.DEVNULL,
                                         stdout=stream, stderr=subprocess.STDOUT,
                                         creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        self.status_text.set('Background sync started (pauses when Outlook is closed).')

    def check_worker(self) -> None:
        import time
        if self.closed:
            return
        if self.proc is None or self.proc.poll() is not None:
            if time.monotonic() >= self.next_worker_check:
                try:
                    self._spawn_worker()
                    self.retry_seconds = min(max(self.retry_seconds * 2, 20), 300)
                    self.next_worker_check = time.monotonic() + self.retry_seconds
                except Exception as exc:
                    self.status_text.set(f'Background worker failed: {exc}')
                    self.next_worker_check = time.monotonic() + self.retry_seconds
        request_show = self.db_path.with_name(self.db_path.name + '.show-request')
        if request_show.exists():
            request_show.unlink(missing_ok=True)
            self.show()
        self.root.after(2000, self.check_worker)

    def show_log(self) -> None:
        log = self.db_path.parent / 'sync.log'
        if not log.exists():
            self.status_text.set('No sync log yet.')
            return
        os.startfile(str(log))

    def poll_jobs(self) -> None:
        from tkinter import messagebox
        while True:
            try:
                action, payload = self.jobs.get_nowait()
            except queue.Empty:
                break
            if action == 'show':
                self.show()
            elif action == 'sync':
                self.sync_now()
            elif action == 'startup':
                try:
                    enabled = not startup_enabled()
                    set_startup(enabled, self.db_path)
                    self.status_text.set('Start with Windows: ' + ('enabled' if enabled else 'disabled'))
                except Exception as exc:
                    messagebox.showerror('Startup shortcut', str(exc))
            elif action == 'log':
                self.show_log()
            elif action == 'quit':
                self.quit()
                return
            elif action == 'search_result':
                sequence, rows = payload
                if sequence != self.search_seq:
                    continue
                children = self.tree.get_children()
                if children:
                    self.tree.delete(*children)
                for row in rows:
                    date = (row['received_at'] or row['sent_at'] or '')[:10]
                    self.tree.insert('', 'end', iid=str(row['id']), values=(
                        date, row['sender_name'] or row['sender_email'],
                        row['subject'], row['folder']))
                self.status_text.set(f'{len(rows)} results (first 100 by relevance).')
                self.show_preview('Select a result to preview the indexed plain-text body.')
            elif action == 'search_error':
                sequence, error = payload
                if sequence == self.search_seq:
                    self.status_text.set('Search failed: ' + error)
            elif action == 'notice':
                self.status_text.set(payload)
            elif action == 'error':
                self.status_text.set('Unable to open: ' + payload)
                messagebox.showerror('Open in Outlook', payload)
        if not self.closed:
            self.root.after(100, self.poll_jobs)

    def quit(self) -> None:
        self.closed = True
        if self.proc is not None and self.proc.poll() is None:
            # SQLite transactions are crash-safe; unfinished folder resumes later.
            self.proc.terminate()
        self.icon.stop()
        self.root.destroy()

    def run(self) -> None:
        # pystray on Windows supports its event loop in another thread.
        threading.Thread(target=self.icon.run, name='win-tray', daemon=True).start()
        self.root.after(100, self.poll_jobs)
        self.root.after(1000, self.check_worker)
        self.root.mainloop()


def single_instance() -> tuple[Any, bool]:
    """Only one icon per interactive Windows login (do not use global mutex)."""
    import ctypes
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.CreateMutexW.argtypes = (ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p)
    kernel.CreateMutexW.restype = ctypes.c_void_p
    handle = kernel.CreateMutexW(None, False, 'Local\\OutlookClassicMCPTray')
    if not handle:
        raise RuntimeError('Failed to create tray mutex.')
    return (handle, ctypes.get_last_error() != 183)  # ERROR_ALREADY_EXISTS


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description='Outlook Classic background tray search')
    parser.add_argument('--db', type=Path, default=default_db_path())
    parser.add_argument('--every-minutes', type=int, default=15)
    parser.add_argument('--search', action='store_true', help='show search UI immediately')
    args = parser.parse_args(argv)
    if sys.platform != 'win32':
        parser.error('Tray application is for Windows 10/11 only')
    if not 1 <= args.every_minutes <= 1440:
        parser.error('--every-minutes must be 1..1440')
    lock, created = single_instance()
    if not created:
        args.db.with_name(args.db.name + '.show-request').write_text('1', encoding='ascii')
        return 0
    try:
        TrayApp(args.db, args.every_minutes, args.search).run()
        return 0
    finally:
        import ctypes
        ctypes.windll.kernel32.CloseHandle(lock)


if __name__ == '__main__':
    raise SystemExit(main())
