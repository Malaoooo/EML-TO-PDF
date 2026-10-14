"""Email to PDF: turns Outlook emails (.msg / .eml) into PDFs that look like they do in Outlook.

Exact copy (.msg, when classic Outlook + Word are set up on this PC): Outlook saves the
email as MHT and Word, the engine classic Outlook renders and prints mail with, makes the PDF.

Close match (everything else): the email's own HTML plus an Outlook-style header block is
printed to PDF by Microsoft Edge in headless mode, the same Chromium engine that new
Outlook and Outlook on the web render mail with. No Outlook install needed.

Run it and pick emails, or drop .msg/.eml files or folders onto the app.
"""
import base64
import datetime
import email
import email.policy
import html
import mimetypes
import os
import re
import subprocess
import sys
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from urllib.parse import unquote
import winreg

import extract_msg
from bs4 import UnicodeDammit

EXTS = {'.msg', '.eml'}

# Print backgrounds/colors exactly, 0.5" margins like Outlook's printing. Scripts never run, as in Outlook.
PAGE_HEAD = ('<meta http-equiv="Content-Security-Policy" content="script-src \'none\'">'
             '<style>@page{margin:0.5in}'
             '*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}</style>')
TD = 'font:11pt Calibri,Arial,sans-serif;color:#000;background:#fff;border:0;padding:1pt 8pt 1pt 0;vertical-align:top;text-align:left'


def norm_cid(cid):
    return unquote(str(cid or '')).strip().strip('<>').lower()


def load_eml(path):
    with open(path, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=email.policy.default)
    body = msg.get_body(('html', 'plain'))
    text = body.get_content() if body else ''
    if body and body.get_content_type() == 'text/plain':
        text = plain_to_html(text)
    inline, files = {}, []
    for part in msg.walk():
        name, cid = part.get_filename(), norm_cid(part['Content-ID'])
        if part.is_multipart() or part is body or not (name or cid):
            continue
        if cid:
            inline[cid] = (part.get_content_type(), part.get_payload(decode=True) or b'')
        files.append((name or cid, cid))
    date = msg['Date']
    return dict(sender=msg['From'], sent=getattr(date, 'datetime', None) or date, to=msg['To'], cc=msg['Cc'],
                subject=msg['Subject'], html=text, inline=inline, files=files)


def load_msg(path):
    m = extract_msg.openMsg(str(path))
    try:
        # htmlBody falls back to HTML wrapped in the RTF body, then to the plain-text body
        text = UnicodeDammit(m.htmlBody or b'', is_html=True).unicode_markup or ''
        inline, files = {}, []
        for a in m.attachments:
            name = getattr(a, 'longFilename', None) or getattr(a, 'shortFilename', None) or getattr(a, 'displayName', None) or 'attachment'
            cid, data = norm_cid(getattr(a, 'cid', None)), getattr(a, 'data', None)
            if cid and isinstance(data, bytes):
                inline[cid] = (getattr(a, 'mimetype', None) or mimetypes.guess_type(name)[0] or 'image/png', data)
            files.append((name, cid))
        return dict(sender=m.sender, sent=m.date, to=m.to, cc=m.cc, subject=m.subject,
                    html=text, inline=inline, files=files)
    finally:
        m.close()


def plain_to_html(text):
    return f'<div style="font:11pt Calibri,Arial,sans-serif;white-space:pre-wrap">{html.escape(text)}</div>'


def fmt_date(d):
    if not isinstance(d, datetime.datetime):
        return str(d or '')
    if d.tzinfo:
        d = d.astimezone()
    return f'{d:%A, %B} {d.day}, {d.year} {d.hour % 12 or 12}:{d:%M %p}'   # Thursday, October 9, 2026 3:05 PM


def insert_after(pattern, page, add):
    m = re.search(pattern, page, re.I)
    return page[:m.end()] + add + page[m.end():] if m else None


def build_html(mail):
    used = set()

    def embed(m):
        cid = norm_cid(m.group(1))
        if cid not in mail['inline']:
            return m.group(0)
        used.add(cid)
        mime, data = mail['inline'][cid]
        return f'data:{mime};base64,{base64.b64encode(data).decode()}'

    page = re.sub(r'cid:([^"\'\s)>]+)', embed, mail['html'] or '', flags=re.I)
    rows = [('From', mail['sender']), ('Sent', fmt_date(mail['sent'])), ('To', mail['to']), ('Cc', mail['cc']),
            ('Subject', mail['subject']), ('Attachments', '; '.join(n for n, cid in mail['files'] if cid not in used))]
    header = ''.join(f'<tr><td style="{TD};font-weight:bold;white-space:nowrap">{k}:</td>'
                     f'<td style="{TD}">{html.escape(str(v))}</td></tr>' for k, v in rows if v)
    header = (f'<table style="border-collapse:collapse;margin:0 0 6pt;width:auto">{header}</table>'
              '<hr style="border:0;border-top:1.5pt solid #000;margin:0 0 12pt;height:0">')
    page = insert_after(r'<body\b[^>]*>', page, header) or header + page
    # head stuff goes after <head>/<html>/<!doctype> so a doctype stays first and the email keeps its layout mode
    for tag in (r'<head\b[^>]*>', r'<html\b[^>]*>', r'<!doctype[^>]*>'):
        if (done := insert_after(tag, page, PAGE_HEAD)) is not None:
            return done
    return PAGE_HEAD + page


def find_browser():
    for base in ('ProgramFiles(x86)', 'ProgramFiles', 'LOCALAPPDATA'):
        for rel in (r'Microsoft\Edge\Application\msedge.exe', r'Google\Chrome\Application\chrome.exe'):
            if os.environ.get(base) and (p := Path(os.environ[base], rel)).exists():
                return p
    raise RuntimeError('Microsoft Edge (or Google Chrome) is needed but was not found')


def to_pdf(page, pdf):
    # ponytail: one Edge launch per email (~1-2 s each); drive one Edge over DevTools if huge batches get slow
    pdf.unlink(missing_ok=True)
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        src = Path(tmp, 'email.html')
        src.write_text(page, encoding='utf-8-sig')   # BOM beats whatever charset the email's HTML declares
        subprocess.run([str(find_browser()), '--headless', '--disable-gpu', '--no-first-run', '--no-pdf-header-footer',
                        f'--user-data-dir={Path(tmp, "profile")}', f'--print-to-pdf={pdf}', src.as_uri()],
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)
    if not pdf.exists():
        raise RuntimeError('Edge did not produce a PDF')


def _reg_has(key):
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key))
        return True
    except OSError:
        return False


def setup_status():
    """Exact-copy checklist as (label, ok) rows."""
    try:   # Office 2016 / 2019 / 2021 / 365 all keep profiles under 16.0
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Office\16.0\Outlook\Profiles') as k:
            signed_in = winreg.QueryInfoKey(k)[0] > 0
    except OSError:
        signed_in = False
    return [('Microsoft Word is installed', _reg_has(r'Word.Application\CLSID')),
            ('Classic Outlook is installed', _reg_has(r'Outlook.Application\CLSID')),
            ('Classic Outlook is signed in to an email account', signed_in)]


class OutlookPrinter:
    """Exact copy: Outlook saves the email (its header block + body + images) as MHT, Word prints it."""

    def __init__(self):
        import win32com.client
        try:
            win32com.client.GetActiveObject('Outlook.Application')
            self.started = False
        except Exception:
            self.started = True
        self.ol = win32com.client.Dispatch('Outlook.Application')
        self.word = win32com.client.DispatchEx('Word.Application')   # own hidden Word, not the user's
        self.word.Visible = False
        self.word.DisplayAlerts = 0
        self.tmp = Path(tempfile.gettempdir(), f'email-to-pdf-{os.getpid()}.mht')

    def convert(self, path, pdf):
        item = self.ol.Session.OpenSharedItem(str(path))
        try:
            item.SaveAs(str(self.tmp), 10)                                   # olMHTML
            doc = self.word.Documents.Open(str(self.tmp), False, True, False)  # ConfirmConversions, ReadOnly, AddToRecent
            try:
                for m in ('TopMargin', 'BottomMargin', 'LeftMargin', 'RightMargin'):
                    setattr(doc.PageSetup, m, 36)                            # 0.5", like Outlook's printing
                doc.ExportAsFixedFormat(str(pdf), 17)                        # wdExportFormatPDF
            finally:
                doc.Close(0)
        finally:
            item.Close(1)                                                    # olDiscard, frees the .msg
            self.tmp.unlink(missing_ok=True)

    def close(self):
        try:
            self.word.Quit(0)
        except Exception:
            pass
        # only quit an Outlook we started, and only if the user hasn't opened a window in it since
        if self.started and self.ol.Explorers.Count == 0:
            self.ol.Quit()


def convert(path, printer=None):
    """Returns (pdf path, 'exact' or 'close')."""
    path = Path(path)
    pdf = path.with_suffix('.pdf')
    if printer and path.suffix.lower() == '.msg':   # Outlook can't open .eml files this way
        pdf.unlink(missing_ok=True)
        try:
            printer.convert(path, pdf)
            return pdf, 'exact'
        except Exception:
            pass   # e.g. a protected or odd .msg: still produce a close match
    mail = load_msg(path) if path.suffix.lower() == '.msg' else load_eml(path)
    to_pdf(build_html(mail), pdf)
    return pdf, 'close'


def expand(paths):
    for p in map(Path, paths):
        if p.is_dir():
            yield from sorted(f for f in p.rglob('*') if f.suffix.lower() in EXTS)
        elif p.suffix.lower() in EXTS:
            yield p


def dark_mode():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize') as k:
            return winreg.QueryValueEx(k, 'AppsUseLightTheme')[0] == 0
    except OSError:
        return False


def style_titlebar(win, dark):
    try:   # Windows 11 dark title bar to match the theme
        import ctypes
        win.update()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(ctypes.c_int(dark)), 4)
    except Exception:
        pass


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # crisp text on high-DPI screens
    except Exception:
        pass
    import sv_ttk
    from tkinter import ttk

    dark = dark_mode()
    root = tk.Tk()
    root.title('Email to PDF')
    root.geometry('820x600')
    root.minsize(640, 460)
    sv_ttk.set_theme('dark' if dark else 'light')
    style_titlebar(root, dark)
    green, amber, red, muted = (('#6ccb5f', '#fce100', '#ff99a4', '#9a9a9a') if dark
                                else ('#0f7b0f', '#9d5d00', '#c42b1c', '#5f5f5f'))
    font = 'Segoe UI Variable Text'
    style = ttk.Style()
    style.configure('Title.TLabel', font=('Segoe UI Variable Display', 22, 'bold'))
    style.configure('Muted.TLabel', foreground=muted)
    style.configure('CardTitle.TLabel', font=(font, 11, 'bold'))
    style.configure('Treeview', rowheight=32)

    page = ttk.Frame(root, padding=(28, 24))
    page.pack(fill='both', expand=True)

    ttk.Label(page, text='Email to PDF', style='Title.TLabel').pack(anchor='w')
    ttk.Label(page, style='Muted.TLabel', text='Turn Outlook emails (.msg or .eml) into PDFs. '
              'Each PDF is saved next to its email.').pack(anchor='w', pady=(2, 18))

    # mode card
    card = ttk.Frame(page, style='Card.TFrame', padding=(18, 14))
    card.pack(fill='x')
    dot = ttk.Label(card, text='●', font=(font, 16))
    dot.pack(side='left', padx=(0, 12))
    texts = ttk.Frame(card, style='Card.TFrame')
    texts.pack(side='left', fill='x', expand=True)
    mode_title = ttk.Label(texts, style='CardTitle.TLabel')
    mode_title.pack(anchor='w')
    mode_desc = ttk.Label(texts, style='Muted.TLabel')
    mode_desc.pack(anchor='w')
    setup_btn = ttk.Button(card, text='Set up')
    setup_btn.pack(side='right')

    # actions
    actions = ttk.Frame(page)
    actions.pack(fill='x', pady=(18, 12))
    pick = ttk.Button(actions, text='Choose emails', style='Accent.TButton')
    pick.pack(side='left', ipadx=10)
    folder = ttk.Button(actions, text='Choose folder')
    folder.pack(side='left', padx=8)
    open_dir = ttk.Button(actions, text='Open PDF folder', state='disabled')
    open_dir.pack(side='right')

    # progress + results
    progress = ttk.Progressbar(page, mode='determinate')
    progress.pack(fill='x')
    summary = ttk.Label(page, style='Muted.TLabel', text='Tip: you can also drop emails or folders onto the app icon.')
    summary.pack(anchor='w', pady=(6, 10))
    table = ttk.Frame(page)
    table.pack(fill='both', expand=True)
    tree = ttk.Treeview(table, columns=('file', 'result'), show='headings', selectmode='browse')
    tree.heading('file', text='Email', anchor='w')
    tree.heading('result', text='Result', anchor='w')
    tree.column('file', anchor='w', stretch=True)
    tree.column('result', anchor='w', width=230, stretch=False)
    tree.tag_configure('exact', foreground=green)
    tree.tag_configure('close', foreground=amber)
    tree.tag_configure('failed', foreground=red)
    scroll = ttk.Scrollbar(table, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    tree.pack(side='left', fill='both', expand=True)
    scroll.pack(side='right', fill='y')
    pdfs, errors, last_dir = {}, {}, []

    def on_open(_):   # double-click: open the PDF, or show why it failed
        row = tree.focus()
        if row in pdfs:
            os.startfile(pdfs[row])
        elif row in errors:
            messagebox.showerror('Could not convert', errors[row], parent=root)

    tree.bind('<Double-1>', on_open)

    def refresh():
        ready = all(ok for _, ok in setup_status())
        dot.configure(foreground=green if ready else amber)
        mode_title.configure(text='Exact copy is on' if ready else 'Close match')
        mode_desc.configure(text='Outlook emails (.msg) are printed by Outlook itself, exactly as Outlook prints them.'
                            if ready else 'Set up classic Outlook on this PC for PDFs that are an exact copy.')
        setup_btn.configure(text='Check setup' if ready else 'Set up')
        return ready

    def run(paths):
        files = list(expand(paths))
        if not files:
            if paths:
                summary.configure(text='No .msg or .eml files found there.')
            return
        for b in (pick, folder, setup_btn):
            b.configure(state='disabled')
        root.configure(cursor='watch')
        progress.configure(maximum=len(files), value=0)
        printer = None
        if refresh() and any(f.suffix.lower() == '.msg' for f in files):
            summary.configure(text='Starting Outlook and Word...')
            root.update()
            try:
                printer = OutlookPrinter()
            except Exception:
                pass   # every file then falls back to close match, which the Result column shows
        ok = 0
        try:
            for i, f in enumerate(files, 1):
                summary.configure(text=f'Converting {i} of {len(files)}: {f.name}')
                root.update()
                try:
                    pdf, mode = convert(f, printer)
                    row = tree.insert('', 0, values=(f.name, '✔  Exact copy' if mode == 'exact' else '✔  Close match'), tags=(mode,))
                    pdfs[row] = pdf
                    last_dir[:] = [pdf.parent]
                    ok += 1
                except Exception as e:
                    row = tree.insert('', 0, values=(f.name, '✖  Failed (double-click for why)'), tags=('failed',))
                    errors[row] = f'{f}\n\n{e}'
                progress.configure(value=i)
        finally:
            if printer:
                printer.close()
            for b in (pick, folder, setup_btn):
                b.configure(state='normal')
            root.configure(cursor='')
        summary.configure(text=f'Done: {ok} of {len(files)} converted. Double-click a row to open its PDF.')
        if last_dir:
            open_dir.configure(state='normal')

    skip_file = Path(os.environ.get('APPDATA', '.'), 'Email to PDF', 'setup-done')

    def setup():
        win = tk.Toplevel(root)
        win.title('Set up exact copy')
        win.transient(root)
        win.resizable(False, False)
        style_titlebar(win, dark)
        body = ttk.Frame(win, padding=24)
        body.pack(fill='both', expand=True)
        ttk.Label(body, text='Set up exact copy', style='CardTitle.TLabel', font=(font, 15, 'bold')).pack(anchor='w')
        ttk.Label(body, style='Muted.TLabel', wraplength=500, justify='left', text=(
            'For PDFs that match exactly what Outlook prints, this PC needs Microsoft Word and classic Outlook '
            '(the Microsoft 365 / Office desktop apps), signed in to an email account once. '
            'Without them the app still works and makes a close match.')).pack(anchor='w', pady=(4, 16))
        rows = ttk.Frame(body, style='Card.TFrame', padding=(16, 10))
        rows.pack(fill='x')
        ttk.Label(body, style='Muted.TLabel', wraplength=500, justify='left', text=(
            'Word or Outlook missing: install the Microsoft 365 / Office desktop apps.\n'
            'Not signed in: click "Open classic Outlook", add your email account, then close Outlook. '
            'If the new Outlook opens instead, turn off its "New Outlook" switch (top right) first.')).pack(anchor='w', pady=16)

        def check():
            for w in rows.winfo_children():
                w.destroy()
            for label, ok in setup_status():
                line = ttk.Frame(rows, style='Card.TFrame')
                line.pack(fill='x', pady=3)
                ttk.Label(line, text='✔' if ok else '✖', foreground=green if ok else red,
                          font=(font, 11, 'bold'), width=3).pack(side='left')
                ttk.Label(line, text=label).pack(side='left')
            refresh()

        def open_outlook():
            try:
                os.startfile('outlook.exe')
            except OSError:
                messagebox.showerror('Email to PDF', 'Classic Outlook is not installed on this PC.', parent=win)

        def close():
            skip_file.parent.mkdir(parents=True, exist_ok=True)
            skip_file.touch()   # don't pop up by itself again; the Set up button still works
            win.destroy()

        btns = ttk.Frame(body)
        btns.pack(fill='x')
        ttk.Button(btns, text='Open classic Outlook', command=open_outlook).pack(side='left')
        ttk.Button(btns, text='Done', style='Accent.TButton', command=close).pack(side='right')
        ttk.Button(btns, text='Check again', command=check).pack(side='right', padx=8)
        win.protocol('WM_DELETE_WINDOW', close)
        check()
        win.grab_set()

    setup_btn.configure(command=setup)
    pick.configure(command=lambda: run(filedialog.askopenfilenames(
        parent=root, filetypes=[('Outlook emails', '*.msg *.eml'), ('All files', '*.*')])))
    folder.configure(command=lambda: run([d] if (d := filedialog.askdirectory(parent=root)) else []))
    open_dir.configure(command=lambda: last_dir and os.startfile(last_dir[0]))

    if not refresh() and not skip_file.exists():
        root.after(300, setup)   # first launch on a PC that isn't set up yet
    if len(sys.argv) > 1:
        root.after(200, run, sys.argv[1:])
    root.mainloop()


if __name__ == '__main__':
    main()
