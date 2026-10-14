"""Email to PDF: turns Outlook emails (.msg / .eml) into PDFs that look like they do in Outlook.

Exact copy (.msg, when classic Outlook + Word are set up on this PC): Outlook saves the
email as MHT and Word, the engine classic Outlook renders and prints mail with, makes the PDF.

Close match (everything else): the email's own HTML plus an Outlook-style header block is
printed to PDF by Microsoft Edge in headless mode, the same Chromium engine that new
Outlook and Outlook on the web render mail with. No Outlook install needed.

Image and PDF attachments can be added as extra pages after the email.
Run it and pick emails, drag emails or folders into the window (also straight from
classic Outlook), or drop them onto the app icon.
"""
import base64
import datetime
import email
import email.policy
import html
import io
import json
import mimetypes
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from urllib.parse import unquote
import winreg

import extract_msg
from bs4 import UnicodeDammit
from pypdf import PdfWriter

VERSION = '1.1.1'
EXTS = {'.msg', '.eml'}
APP_DIR = Path(os.environ.get('APPDATA', '.'), 'Email to PDF')
DEFAULT_OUT = Path.home() / 'Documents' / 'Email PDFs'   # for emails dragged out of Outlook (they have no folder)
NAMINGS = {'file': 'Same name as the email file', 'date-subject': 'Date - Subject', 'subject': 'Subject'}
DEFAULTS = {'out_dir': '', 'naming': 'file', 'attachments': True}
PAGE_IMAGES = {'image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp', 'image/svg+xml'}
CID_RE = r'cid:([^"\'\s)>]+)'

# Print backgrounds/colors exactly, 0.5" margins like Outlook's printing. Scripts never run, as in Outlook.
PAGE_HEAD = ('<meta http-equiv="Content-Security-Policy" content="script-src \'none\'">'
             '<style>@page{margin:0.5in}'
             '*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}</style>')
TD = 'font:11pt Calibri,Arial,sans-serif;color:#000;background:#fff;border:0;padding:1pt 8pt 1pt 0;vertical-align:top;text-align:left'


def load_settings():
    try:
        return {**DEFAULTS, **json.loads((APP_DIR / 'settings.json').read_text(encoding='utf-8'))}
    except (OSError, ValueError):
        return dict(DEFAULTS)


def save_settings(s):
    APP_DIR.mkdir(parents=True, exist_ok=True)
    (APP_DIR / 'settings.json').write_text(json.dumps(s, indent=2), encoding='utf-8')


def norm_cid(cid):
    return unquote(str(cid or '')).strip().strip('<>').lower()


def load_eml(path):
    with open(path, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=email.policy.default)
    body = msg.get_body(('html', 'plain'))
    text = body.get_content() if body else ''
    if body and body.get_content_type() == 'text/plain':
        text = plain_to_html(text)
    files = []
    for part in msg.walk():
        name, cid = part.get_filename(), norm_cid(part['Content-ID'])
        if part.is_multipart() or part is body or not (name or cid):
            continue
        files.append(dict(name=name or cid, cid=cid, mime=part.get_content_type(), data=part.get_payload(decode=True) or b''))
    date = msg['Date']
    return dict(sender=msg['From'], sent=getattr(date, 'datetime', None) or date, to=msg['To'], cc=msg['Cc'],
                subject=msg['Subject'], html=text, files=files)


def load_msg(path):
    m = extract_msg.openMsg(str(path))
    try:
        # htmlBody falls back to HTML wrapped in the RTF body, then to the plain-text body
        text = UnicodeDammit(m.htmlBody or b'', is_html=True).unicode_markup or ''
        files = []
        for a in m.attachments:
            name = getattr(a, 'longFilename', None) or getattr(a, 'shortFilename', None) or getattr(a, 'displayName', None) or 'attachment'
            data = getattr(a, 'data', None)   # an attached email is a Message object, not bytes
            files.append(dict(name=name, cid=norm_cid(getattr(a, 'cid', None)), data=data if isinstance(data, bytes) else None,
                              mime=getattr(a, 'mimetype', None) or mimetypes.guess_type(name)[0]))
        return dict(sender=m.sender, sent=m.date, to=m.to, cc=m.cc, subject=m.subject, html=text, files=files)
    finally:
        m.close()


def load(path):
    return load_msg(path) if Path(path).suffix.lower() == '.msg' else load_eml(path)


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


def real_attachments(mail):
    """Attachments that aren't pictures shown inside the email body."""
    shown = {norm_cid(c) for c in re.findall(CID_RE, mail['html'] or '', re.I)}
    return [f for f in mail['files'] if not (f['cid'] and f['cid'] in shown)]


def build_html(mail):
    inline = {f['cid']: f for f in mail['files'] if f['cid'] and f['data']}

    def embed(m):
        f = inline.get(norm_cid(m.group(1)))
        if not f:
            return m.group(0)
        return f'data:{f["mime"] or "image/png"};base64,{base64.b64encode(f["data"]).decode()}'

    page = re.sub(CID_RE, embed, mail['html'] or '', flags=re.I)
    rows = [('From', mail['sender']), ('Sent', fmt_date(mail['sent'])), ('To', mail['to']), ('Cc', mail['cc']),
            ('Subject', mail['subject']), ('Attachments', '; '.join(f['name'] for f in real_attachments(mail)))]
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
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    if not pdf.exists():
        raise RuntimeError('Edge did not produce a PDF')


def append_attachments(pdf, mail):
    """Adds image and PDF attachments as pages after the email. Returns (added, skipped)."""
    images, pdfs = [], []
    for f in real_attachments(mail):
        mime = (f['mime'] or mimetypes.guess_type(f['name'])[0] or '').lower()
        if not f['data']:
            continue
        if mime in PAGE_IMAGES:
            images.append((f, mime))
        elif mime == 'application/pdf' or f['name'].lower().endswith('.pdf'):
            pdfs.append(f)
    if not images and not pdfs:
        return 0, 0
    writer = PdfWriter(clone_from=str(pdf))
    added = skipped = 0
    if images:   # one page per picture, captioned with its file name
        page = ''.join(f'<section><p>{html.escape(f["name"])}</p>'
                       f'<img src="data:{mime};base64,{base64.b64encode(f["data"]).decode()}"></section>' for f, mime in images)
        shots = pdf.with_name(pdf.stem + '.attachments.tmp.pdf')
        to_pdf('<!DOCTYPE html><style>@page{margin:0.5in}body{margin:0;font:10pt Calibri,Arial,sans-serif;color:#555}'
               'section{break-after:page}section:last-child{break-after:auto}p{margin:0 0 6pt}'
               'img{display:block;max-width:100%;max-height:9.2in;object-fit:contain}</style>' + page, shots)
        try:
            writer.append(str(shots))
            added += len(images)
        finally:
            shots.unlink(missing_ok=True)
    for f in pdfs:
        try:
            writer.append(io.BytesIO(f['data']))
            added += 1
        except Exception:   # password-protected or broken PDF
            skipped += 1
    tmp = pdf.with_name(pdf.stem + '.tmp.pdf')
    with open(tmp, 'wb') as out:
        writer.write(out)
    writer.close()
    os.replace(tmp, pdf)
    return added, skipped


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
    # Some setups leave that key empty; a signed-in classic Outlook always has a mailbox file (.ost / .pst)
    mail_dirs = (Path(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Outlook'), Path.home() / 'Documents' / 'Outlook Files')
    signed_in = signed_in or any(f.suffix.lower() in ('.ost', '.pst') for d in mail_dirs if d.is_dir() for f in d.iterdir())
    return [('Microsoft Word is installed', _reg_has(r'Word.Application\CLSID')),
            ('Classic Outlook is installed', _reg_has(r'Outlook.Application\CLSID')),
            ('Classic Outlook is signed in to an email account', signed_in)]


def outlook():
    """Returns (Outlook app, whether we started it)."""
    import win32com.client
    try:
        win32com.client.GetActiveObject('Outlook.Application')
        started = False
    except Exception:
        started = True
    return win32com.client.Dispatch('Outlook.Application'), started


class OutlookPrinter:
    """Exact copy: Outlook saves the email (its header block + body + images) as MHT, Word prints it."""

    def __init__(self):
        import win32com.client
        self.ol, self.started = outlook()
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


def safe_name(text, limit=120):
    text = re.sub(r'[\\/:*?"<>|\x00-\x1f]', '_', str(text or ''))[:limit].strip(' .')
    return text or '(no subject)'


def pdf_target(src, mail, settings, taken, fallback_dir=None):
    """Where the PDF goes, following the Save-to and naming settings. `taken` avoids clashes within one batch."""
    folder = Path(settings['out_dir'] or fallback_dir or src.parent)
    naming, sent = settings['naming'], mail['sent']
    if naming == 'file':
        stem = src.stem
    else:
        stem = safe_name(mail['subject'])
        if naming == 'date-subject':
            if not isinstance(sent, datetime.datetime):
                sent = datetime.datetime.fromtimestamp(src.stat().st_mtime)
            stem = f'{sent.astimezone() if sent.tzinfo else sent:%Y-%m-%d %H%M} - {stem}'
    pdf, n = folder / f'{stem}.pdf', 2
    while str(pdf).lower() in taken:
        pdf, n = folder / f'{stem} ({n}).pdf', n + 1
    taken.add(str(pdf).lower())
    return pdf


def convert(path, settings=None, printer=None, taken=None, fallback_dir=None):
    """Returns (pdf path, 'exact' or 'close', attachments added, attachments skipped)."""
    path, settings = Path(path), settings or DEFAULTS
    mail = load(path)
    pdf = pdf_target(path, mail, settings, set() if taken is None else taken, fallback_dir)
    pdf.parent.mkdir(parents=True, exist_ok=True)
    mode = 'close'
    if printer and path.suffix.lower() == '.msg':   # Outlook can't open .eml files this way
        pdf.unlink(missing_ok=True)
        try:
            printer.convert(path, pdf)
            mode = 'exact'
        except Exception:
            pass   # e.g. a protected or odd .msg: still produce a close match
    if mode == 'close':
        to_pdf(build_html(mail), pdf)
    added, skipped = append_attachments(pdf, mail) if settings['attachments'] else (0, 0)
    return pdf, mode, added, skipped


def expand(paths):
    for p in map(Path, paths):
        if p.is_dir():
            yield from sorted(f for f in p.rglob('*') if f.suffix.lower() in EXTS)
        elif p.suffix.lower() in EXTS:
            yield p


def export_outlook_selection(folder):
    """Saves the emails selected in classic Outlook (what was just dragged) as .msg files into folder."""
    ol, _ = outlook()
    exp = ol.ActiveExplorer()
    if not exp:
        return []
    sel, out = exp.Selection, []
    for i in range(1, sel.Count + 1):
        item = sel.Item(i)
        p, n = Path(folder, safe_name(getattr(item, 'Subject', ''), 100) + '.msg'), 2
        while p.exists():
            p, n = Path(folder, f'{safe_name(item.Subject, 100)} ({n}).msg'), n + 1
        item.SaveAs(str(p), 9)   # olMSGUnicode
        out.append(p)
    return out


def run_batch(job, settings, ready, events, cancel):
    """Worker thread: converts the job and reports progress to the window through `events`."""
    import pythoncom
    pythoncom.CoInitialize()
    temp = printer = None
    ok = 0
    files = []
    try:
        fallback, empty = None, 'No .msg or .eml files found there.'
        if job == 'outlook':
            events.put(('status', 'Getting the emails from Outlook...'))
            temp = tempfile.mkdtemp(prefix='email-to-pdf-')
            empty = 'Nothing to convert: select emails in classic Outlook, then drag them here.'
            try:
                files = export_outlook_selection(temp)
            except Exception as e:
                empty = f'Could not get the emails from Outlook: {e}'
            fallback = DEFAULT_OUT
        else:
            files = list(expand(job))
        if not files:
            events.put(('status', empty))
            return
        events.put(('start', len(files)))
        if ready and any(f.suffix.lower() == '.msg' for f in files):
            events.put(('status', 'Starting Outlook and Word...'))
            try:
                printer = OutlookPrinter()
            except Exception:
                pass   # every file then falls back to close match, which the Result column shows
        taken = set()
        for i, f in enumerate(files, 1):
            if cancel.is_set():
                break
            events.put(('status', f'Converting {i} of {len(files)}: {f.name}'))
            try:
                pdf, mode, added, skipped = convert(f, settings, printer, taken, fallback)
                note = 'Exact copy' if mode == 'exact' else 'Close match'
                if added:
                    note += f' + {added} attachment{"s" * (added > 1)}'
                if skipped:
                    note += f' ({skipped} skipped)'
                events.put(('row', f.name, '✔  ' + note, mode, str(pdf)))
                ok += 1
            except Exception as e:
                events.put(('row', f.name, '✖  Failed (double-click for why)', 'failed', f'{f}\n\n{e}'))
            events.put(('progress', i))
    finally:
        if printer:
            printer.close()
        if temp:
            shutil.rmtree(temp, ignore_errors=True)
        if files:
            word = 'Cancelled' if cancel.is_set() else 'Done'
            events.put(('status', f'{word}: {ok} of {len(files)} converted. Double-click a row to open its PDF.'))
        events.put(('done',))
        pythoncom.CoUninitialize()


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
    from tkinterdnd2 import COPY, DND_ALL, TkinterDnD

    dark = dark_mode()
    settings = load_settings()
    root = TkinterDnD.Tk()
    root.title('Email to PDF')
    root.geometry('860x680')
    root.minsize(680, 540)
    try:
        root.iconbitmap(default=str(Path(getattr(sys, '_MEIPASS', Path(__file__).parent), 'icon.ico')))
    except tk.TclError:
        pass
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

    head = ttk.Frame(page)
    head.pack(fill='x')
    ttk.Label(head, text='Email to PDF', style='Title.TLabel').pack(side='left')
    ttk.Label(head, text=f'v{VERSION}', style='Muted.TLabel').pack(side='left', padx=10, pady=(12, 0))
    ttk.Label(page, style='Muted.TLabel', text='Turn Outlook emails (.msg or .eml) into PDFs. '
              'Drag emails or folders into this window, straight from Outlook too.').pack(anchor='w', pady=(2, 16))

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

    # options card: where to save, how to name, attachments
    opts = ttk.Frame(page, style='Card.TFrame', padding=(18, 12))
    opts.pack(fill='x', pady=(12, 0))
    opts.columnconfigure(1, weight=1)
    ttk.Label(opts, text='Save PDFs to').grid(row=0, column=0, sticky='w', padx=(0, 14), pady=4)
    out_label = ttk.Label(opts, style='Muted.TLabel')
    out_label.grid(row=0, column=1, sticky='w')
    out_btns = ttk.Frame(opts, style='Card.TFrame')
    out_btns.grid(row=0, column=2, sticky='e')
    change_out = ttk.Button(out_btns, text='Change...')
    change_out.pack(side='left')
    reset_out = ttk.Button(out_btns, text='Next to each email')
    reset_out.pack(side='left', padx=(6, 0))
    ttk.Label(opts, text='File names').grid(row=1, column=0, sticky='w', padx=(0, 14), pady=4)
    naming = ttk.Combobox(opts, state='readonly', width=30, values=list(NAMINGS.values()))
    naming.set(NAMINGS.get(settings['naming'], NAMINGS['file']))
    naming.grid(row=1, column=1, sticky='w')
    attach_var = tk.BooleanVar(value=settings['attachments'])
    attach = ttk.Checkbutton(opts, text='Add image and PDF attachments as pages after the email', variable=attach_var)
    attach.grid(row=2, column=0, columnspan=3, sticky='w', pady=(6, 0))

    def show_out():
        out_label.configure(text=settings['out_dir'] or 'Next to each email (dragged from Outlook: Documents\\Email PDFs)')
        reset_out.configure(state='normal' if settings['out_dir'] else 'disabled')

    def set_out(value):
        settings['out_dir'] = value
        save_settings(settings)
        show_out()

    def on_naming(_):
        settings['naming'] = next(k for k, v in NAMINGS.items() if v == naming.get())
        save_settings(settings)

    def on_attach():
        settings['attachments'] = attach_var.get()
        save_settings(settings)

    change_out.configure(command=lambda: (d := filedialog.askdirectory(parent=root)) and set_out(str(Path(d))))
    reset_out.configure(command=lambda: set_out(''))
    naming.bind('<<ComboboxSelected>>', on_naming)
    attach.configure(command=on_attach)
    show_out()

    # actions
    actions = ttk.Frame(page)
    actions.pack(fill='x', pady=(16, 12))
    pick = ttk.Button(actions, text='Choose emails', style='Accent.TButton')
    pick.pack(side='left', ipadx=10)
    folder = ttk.Button(actions, text='Choose folder')
    folder.pack(side='left', padx=8)
    cancel_btn = ttk.Button(actions, text='Cancel')
    open_dir = ttk.Button(actions, text='Open PDF folder', state='disabled')
    open_dir.pack(side='right')

    # progress + results
    progress = ttk.Progressbar(page, mode='determinate')
    progress.pack(fill='x')
    summary = ttk.Label(page, style='Muted.TLabel', text='Drop emails or folders anywhere in this window to convert them.')
    summary.pack(anchor='w', pady=(6, 10))
    table = ttk.Frame(page)
    table.pack(fill='both', expand=True)
    tree = ttk.Treeview(table, columns=('file', 'result'), show='headings', selectmode='browse')
    tree.heading('file', text='Email', anchor='w')
    tree.heading('result', text='Result', anchor='w')
    tree.column('file', anchor='w', stretch=True)
    tree.column('result', anchor='w', width=300, stretch=False)
    tree.tag_configure('exact', foreground=green)
    tree.tag_configure('close', foreground=amber)
    tree.tag_configure('failed', foreground=red)
    scroll = ttk.Scrollbar(table, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    tree.pack(side='left', fill='both', expand=True)
    scroll.pack(side='right', fill='y')
    pdfs, errors, last_dir = {}, {}, []
    state = {'busy': False, 'closing': False}
    events, cancel = queue.Queue(), threading.Event()

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

    def set_busy(busy):
        state['busy'] = busy
        for w in (pick, folder, setup_btn, change_out, attach):
            w.configure(state='disabled' if busy else 'normal')
        naming.configure(state='disabled' if busy else 'readonly')
        if not busy:
            show_out()
        if busy:
            cancel_btn.configure(state='normal', text='Cancel')
            cancel_btn.pack(side='left')
        else:
            cancel_btn.pack_forget()
        root.configure(cursor='watch' if busy else '')

    def start(job):
        if state['busy'] or not job:
            return
        cancel.clear()
        set_busy(True)
        progress.configure(value=0)
        threading.Thread(target=run_batch, args=(job, dict(settings), refresh(), events, cancel), daemon=True).start()

    def poll():
        try:
            while True:
                kind, *data = events.get_nowait()
                if kind == 'status':
                    summary.configure(text=data[0])
                elif kind == 'start':
                    progress.configure(maximum=data[0], value=0)
                elif kind == 'progress':
                    progress.configure(value=data[0])
                elif kind == 'row':
                    name, result, tag, target = data
                    row = tree.insert('', 0, values=(name, result), tags=(tag,))
                    if tag == 'failed':
                        errors[row] = target
                    else:
                        pdfs[row] = target
                        last_dir[:] = [Path(target).parent]
                        open_dir.configure(state='normal')
                elif kind == 'done':
                    set_busy(False)
                    if state['closing']:
                        root.destroy()
                        return
        except queue.Empty:
            pass
        root.after(100, poll)

    def on_cancel():
        cancel.set()
        cancel_btn.configure(state='disabled', text='Cancelling...')

    def on_close():
        if state['busy']:   # finish the current email and close Outlook/Word cleanly first
            state['closing'] = True
            on_cancel()
        else:
            root.destroy()

    # drag and drop: files/folders from Explorer, or emails straight out of classic Outlook
    def on_drop_enter(event):
        summary.configure(text='Drop to convert')
        return COPY

    def on_drop(event):
        paths = [p for p in root.tk.splitlist(event.data) if os.path.exists(p)] if event.data else []
        # Outlook drags carry no files; the dragged emails are Outlook's current selection
        root.after(50, start, paths or 'outlook')
        return COPY   # never MOVE, so Outlook keeps its emails

    root.drop_target_register(DND_ALL)
    root.dnd_bind('<<DropEnter>>', on_drop_enter)
    root.dnd_bind('<<DropLeave>>', lambda e: summary.configure(text='Drop emails or folders anywhere in this window to convert them.'))
    root.dnd_bind('<<Drop>>', on_drop)

    skip_file = APP_DIR / 'setup-done'

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
    pick.configure(command=lambda: start(list(filedialog.askopenfilenames(
        parent=root, filetypes=[('Outlook emails', '*.msg *.eml'), ('All files', '*.*')]))))
    folder.configure(command=lambda: start([d] if (d := filedialog.askdirectory(parent=root)) else []))
    open_dir.configure(command=lambda: last_dir and os.startfile(last_dir[0]))
    cancel_btn.configure(command=on_cancel)
    root.protocol('WM_DELETE_WINDOW', on_close)

    if not refresh() and not skip_file.exists():
        root.after(300, setup)   # first launch on a PC that isn't set up yet
    if len(sys.argv) > 1:
        root.after(200, start, sys.argv[1:])
    root.after(100, poll)
    root.mainloop()


if __name__ == '__main__':
    main()
