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

VERSION = '1.2.0'
EXTS = {'.msg', '.eml'}
APP_DIR = Path(os.environ.get('APPDATA', '.'), 'Email to PDF')
DEFAULT_OUT = Path.home() / 'Documents' / 'Email PDFs'   # for emails dragged out of Outlook (they have no folder)
NAMINGS = {'file': 'Same name as the email file', 'date-subject': 'Date - Subject', 'subject': 'Subject'}
DEFAULTS = {'out_dir': '', 'naming': 'file', 'attachments': True, 'theme': 'system'}
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
        events.put(('queue', [f.name for f in files]))
        if ready and any(f.suffix.lower() == '.msg' for f in files):
            events.put(('status', 'Starting Outlook and Word...'))
            try:
                printer = OutlookPrinter()
            except Exception:
                pass   # every file then falls back to close match, which the Result column shows
        taken = set()
        for i, f in enumerate(files):
            if cancel.is_set():
                break
            events.put(('working', i, f'Converting {i + 1} of {len(files)}: {f.name}'))
            src = None if temp else str(f)   # emails dragged from Outlook only existed temporarily
            try:
                pdf, mode, added, skipped = convert(f, settings, printer, taken, fallback)
                note = 'Exact copy' if mode == 'exact' else 'Close match'
                if added:
                    note += f' + {added} attachment{"s" * (added > 1)}'
                if skipped:
                    note += f' ({skipped} skipped)'
                events.put(('row', i, '✔  ' + note, mode, str(pdf), src))
                ok += 1
            except Exception as e:
                events.put(('row', i, '✖  Failed (double-click for why)', 'failed', f'{f}\n\n{e}', src))
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
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(ctypes.c_int(dark)), 4)
    except Exception:
        pass


def mix(a, b, t):
    """Blend two #rrggbb colors; t=0 gives a, t=1 gives b."""
    rgb = lambda c: [int(c[i:i + 2], 16) for i in (1, 3, 5)]
    return '#' + ''.join(f'{round(x + (y - x) * t):02x}' for x, y in zip(rgb(a), rgb(b)))


THEMES = {'system': 'Match Windows', 'light': 'Light', 'dark': 'Dark'}
SPINNER = '◐◓◑◒'


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # crisp text on high-DPI screens
    except Exception:
        pass
    import sv_ttk
    from tkinter import ttk
    from tkinterdnd2 import COPY, DND_ALL, TkinterDnD

    settings = load_settings()
    here = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
    root = TkinterDnD.Tk()
    root.title('Email to PDF')
    root.geometry('900x700')
    root.minsize(700, 560)
    try:
        root.iconbitmap(default=str(here / 'icon.ico'))
    except tk.TclError:
        pass
    logo = tk.PhotoImage(file=str(here / 'icon.png'))
    font = 'Segoe UI Variable Text'
    style = ttk.Style()
    pal = {}

    page = ttk.Frame(root, padding=(28, 22))
    page.pack(fill='both', expand=True)

    # header
    head = ttk.Frame(page)
    head.pack(fill='x')
    ttk.Label(head, text='Email to PDF', style='Title.TLabel').pack(side='left')
    ttk.Label(head, text=f'v{VERSION}', style='Muted.TLabel').pack(side='left', padx=10, pady=(12, 0))
    gear = ttk.Button(head, text='⚙  Settings')
    gear.pack(side='right')

    # mode card
    card = ttk.Frame(page, style='Card.TFrame', padding=(18, 14))
    card.pack(fill='x', pady=(14, 14))
    dot = ttk.Label(card, text='●', font=(font, 16))
    dot.pack(side='left', padx=(0, 12))
    texts = ttk.Frame(card)
    texts.pack(side='left', fill='x', expand=True)
    mode_title = ttk.Label(texts, style='CardTitle.TLabel')
    mode_title.pack(anchor='w')
    mode_desc = ttk.Label(texts, style='Muted.TLabel')
    mode_desc.pack(anchor='w')
    setup_btn = ttk.Button(card, text='Set up')
    setup_btn.pack(side='right')

    content = ttk.Frame(page)
    content.pack(fill='both', expand=True)

    # --- empty state: the drop zone ---
    zone = tk.Canvas(content, highlightthickness=0, bd=0)
    zone_state = {'hot': False}
    pick = ttk.Button(zone, text='Choose emails', style='Accent.TButton')
    folder = ttk.Button(zone, text='Choose folder')

    def draw_zone(_=None):
        zone.delete('all')
        w, h = zone.winfo_width(), zone.winfo_height()
        if w < 50 or h < 50:
            return
        hot = zone_state['hot']
        fill = pal['zone_hot'] if hot else pal['zone']
        line = pal['accent'] if hot else pal['dash']
        zone.configure(bg=pal['bg'])
        x0, y0, x1, y1, r = 2, 2, w - 3, h - 3, 18
        zone.create_polygon(x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1, x1 - r, y1,
                            x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0, smooth=True, fill=fill, outline='')
        dash = (8, 6)   # dashed border drawn by hand: straight edges + arc corners
        for a, b in (((x0 + r, y0), (x1 - r, y0)), ((x0 + r, y1), (x1 - r, y1)),
                     ((x0, y0 + r), (x0, y1 - r)), ((x1, y0 + r), (x1, y1 - r))):
            horizontal = a[1] == b[1]
            length = (b[0] - a[0]) if horizontal else (b[1] - a[1])
            pos = 0
            while pos < length:
                end = min(pos + dash[0], length)
                if horizontal:
                    zone.create_line(a[0] + pos, a[1], a[0] + end, a[1], fill=line, width=2)
                else:
                    zone.create_line(a[0], a[1] + pos, a[0], a[1] + end, fill=line, width=2)
                pos += sum(dash)
        for cx, cy, start in ((x0, y0, 90), (x1 - 2 * r, y0, 0), (x0, y1 - 2 * r, 180), (x1 - 2 * r, y1 - 2 * r, 270)):
            zone.create_arc(cx, cy, cx + 2 * r, cy + 2 * r, start=start, extent=90, style='arc', outline=line, width=2)
        cx, cy = w / 2, h / 2
        zone.create_image(cx, cy - 78, image=logo)
        zone.create_text(cx, cy - 12, text='Drop emails here' if hot else 'Drag emails here',
                         fill=pal['accent'] if hot else pal['fg'], font=(font, 16, 'bold'))
        zone.create_text(cx, cy + 16, fill=pal['muted'], font=(font, 10),
                         text='.msg or .eml files, folders, or emails straight from classic Outlook')
        zone.create_window(cx - 8, cy + 62, window=pick, anchor='e')
        zone.create_window(cx + 8, cy + 62, window=folder, anchor='w')

    def set_hot(hot):
        if zone_state['hot'] != hot:
            zone_state['hot'] = hot
            draw_zone()

    zone.bind('<Configure>', draw_zone)
    zone.bind('<Enter>', lambda e: set_hot(True))
    zone.bind('<Leave>', lambda e: set_hot(False) if not zone.winfo_containing(e.x_root, e.y_root) in (pick, folder) else None)

    # --- results view ---
    results = ttk.Frame(content)
    bar = ttk.Frame(results)
    bar.pack(fill='x')
    add_more = ttk.Button(bar, text='Add emails', style='Accent.TButton')
    add_more.pack(side='left')
    cancel_btn = ttk.Button(bar, text='Cancel')
    clear_btn = ttk.Button(bar, text='Clear')
    clear_btn.pack(side='right')
    open_dir = ttk.Button(bar, text='Open PDF folder', state='disabled')
    open_dir.pack(side='right', padx=8)
    progress = ttk.Progressbar(results, mode='determinate')
    progress.pack(fill='x', pady=(12, 0))
    summary = ttk.Label(results, style='Muted.TLabel', text='')
    summary.pack(anchor='w', pady=(6, 10))
    table = ttk.Frame(results)
    table.pack(fill='both', expand=True)
    tree = ttk.Treeview(table, columns=('file', 'result'), show='headings', selectmode='browse')
    tree.heading('file', text='Email', anchor='w')
    tree.heading('result', text='Result', anchor='w')
    tree.column('file', anchor='w', stretch=True)
    tree.column('result', anchor='w', width=320, stretch=False)
    scroll = ttk.Scrollbar(table, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    tree.pack(side='left', fill='both', expand=True)
    scroll.pack(side='right', fill='y')
    ttk.Label(results, style='Muted.TLabel', text='Tip: drop more emails anywhere in the window. '
              'Right-click a row for more options.').pack(anchor='w', pady=(8, 0))

    rows = {}        # tree row id -> {'pdf': path, 'src': path, 'error': text}
    batch = []       # tree row ids of the running batch, in order
    state = {'busy': False, 'closing': False, 'working': None, 'spin': 0}
    events, cancel = queue.Queue(), threading.Event()
    last_dir = []

    def show(view):
        zone.pack_forget()
        results.pack_forget()
        (results if view == 'results' else zone).pack(fill='both', expand=True)

    # --- theme ---
    def apply_theme():
        dark = dark_mode() if settings['theme'] == 'system' else settings['theme'] == 'dark'
        sv_ttk.set_theme('dark' if dark else 'light')
        style_titlebar(root, dark)
        pal.update(
            bg=style.lookup('TFrame', 'background') or ('#1c1c1c' if dark else '#fafafa'),
            fg='#ffffff' if dark else '#1c1c1c', muted='#9a9a9a' if dark else '#5f5f5f',
            accent='#4cc2ff' if dark else '#005fb8', dash='#5a5a5a' if dark else '#b8b8b8',
            zone='#232323' if dark else '#f3f3f3', zone_hot='#1a3346' if dark else '#e5f1fb',
            exact='#6ccb5f' if dark else '#0f7b0f', close='#fce100' if dark else '#9d5d00',
            failed='#ff99a4' if dark else '#c42b1c')
        # ttk styles are per theme, so (re)configure them after every switch
        style.configure('Title.TLabel', font=('Segoe UI Variable Display', 22, 'bold'))
        style.configure('Muted.TLabel', foreground=pal['muted'])
        style.configure('CardTitle.TLabel', font=(font, 11, 'bold'))
        style.configure('Treeview', rowheight=34)
        for tag in ('exact', 'close', 'failed'):
            tree.tag_configure(tag, foreground=pal[tag])
            for k in range(6):   # fade-in steps from muted to the final color
                tree.tag_configure(f'{tag}{k}', foreground=mix(pal['muted'], pal[tag], k / 6))
        tree.tag_configure('pending', foreground=pal['muted'])
        tree.tag_configure('working', foreground=pal['accent'])
        refresh()
        draw_zone()

    # --- mode card ---
    def refresh():
        ready = all(ok for _, ok in setup_status())
        dot.configure(foreground=pal['exact'] if ready else pal['close'])
        mode_title.configure(text='Exact copy is on' if ready else 'Close match')
        mode_desc.configure(text='Outlook emails (.msg) are printed by Outlook itself, exactly as Outlook prints them.'
                            if ready else 'Set up classic Outlook on this PC for PDFs that are an exact copy.')
        setup_btn.configure(text='Check setup' if ready else 'Set up')
        return ready

    # --- running a batch ---
    def set_busy(busy):
        state['busy'] = busy
        for w in (pick, folder, add_more, setup_btn, clear_btn):
            w.configure(state='disabled' if busy else 'normal')
        if busy:
            cancel_btn.configure(state='normal', text='Cancel')
            cancel_btn.pack(side='left', padx=8)
        else:
            cancel_btn.pack_forget()
        root.configure(cursor='watch' if busy else '')

    def start(job):
        if state['busy'] or not job:
            return
        cancel.clear()
        batch.clear()
        set_busy(True)
        show('results')
        progress.configure(value=0)
        summary.configure(text='Starting...')
        threading.Thread(target=run_batch, args=(job, dict(settings), refresh(), events, cancel), daemon=True).start()

    def fade(row, tag, k=1):
        if not tree.exists(row):
            return
        tree.item(row, tags=(f'{tag}{k}',) if k < 6 else (tag,))
        if k < 6:
            root.after(40, fade, row, tag, k + 1)

    def spin():
        row = state['working']
        if row and tree.exists(row):
            state['spin'] = (state['spin'] + 1) % len(SPINNER)
            tree.set(row, 'result', f'{SPINNER[state["spin"]]}  Converting...')
        root.after(120, spin)

    def poll():
        try:
            while True:
                kind, *data = events.get_nowait()
                if kind == 'status':
                    summary.configure(text=data[0])
                elif kind == 'queue':   # all rows appear at once, newest batch on top
                    progress.configure(maximum=len(data[0]), value=0)
                    for i, name in enumerate(data[0]):
                        batch.append(tree.insert('', i, values=(name, '…  Waiting'), tags=('pending',)))
                    tree.see(batch[0])
                elif kind == 'working':
                    i, text = data
                    state['working'] = batch[i]
                    tree.item(batch[i], tags=('working',))
                    summary.configure(text=text)
                elif kind == 'row':
                    i, result, tag, target, src = data
                    row = batch[i]
                    state['working'] = None
                    tree.set(row, 'result', result)
                    rows[row] = {'src': src, 'error' if tag == 'failed' else 'pdf': target}
                    if tag != 'failed':
                        last_dir[:] = [Path(target).parent]
                        open_dir.configure(state='normal')
                    progress.configure(value=i + 1)
                    fade(row, tag)
                elif kind == 'done':
                    state['working'] = None
                    for row in batch:   # anything never reached (cancelled)
                        if row not in rows and tree.exists(row):
                            tree.set(row, 'result', '–  Cancelled')
                    set_busy(False)
                    if state['closing']:
                        root.destroy()
                        return
        except queue.Empty:
            pass
        root.after(80, poll)

    def on_cancel():
        cancel.set()
        cancel_btn.configure(state='disabled', text='Cancelling...')

    def on_close():
        if state['busy']:   # finish the current email and close Outlook/Word cleanly first
            state['closing'] = True
            on_cancel()
        else:
            root.destroy()

    def clear():
        tree.delete(*tree.get_children())
        rows.clear()
        last_dir.clear()
        open_dir.configure(state='disabled')
        show('zone')

    # --- row actions: double-click and right-click menu ---
    def open_row(row):
        info = rows.get(row, {})
        if info.get('pdf'):
            os.startfile(info['pdf'])
        elif info.get('error'):
            messagebox.showerror('Could not convert', info['error'], parent=root)

    def reveal(path):
        subprocess.Popen(['explorer', '/select,', str(Path(path))])

    menu = tk.Menu(root, tearoff=False)

    def on_right_click(e):
        row = tree.identify_row(e.y)
        if not row or row not in rows:
            return
        tree.selection_set(row)
        info = rows[row]
        pdf, src = info.get('pdf'), info.get('src')
        menu.delete(0, 'end')
        menu.add_command(label='Open PDF', command=lambda: os.startfile(pdf), state='normal' if pdf else 'disabled')
        menu.add_command(label='Show in folder', command=lambda: reveal(pdf), state='normal' if pdf else 'disabled')
        menu.add_command(label='Open original email', command=lambda: os.startfile(src),
                         state='normal' if src and os.path.exists(src) else 'disabled')
        menu.add_separator()
        menu.add_command(label='Copy PDF path', state='normal' if pdf else 'disabled',
                         command=lambda: (root.clipboard_clear(), root.clipboard_append(pdf)))
        if info.get('error'):
            menu.add_command(label='Why did it fail?', command=lambda: open_row(row))
        menu.tk_popup(e.x_root, e.y_root)

    tree.bind('<Double-1>', lambda e: open_row(tree.identify_row(e.y)))
    tree.bind('<Button-3>', on_right_click)

    # --- drag and drop: files/folders from Explorer, or emails straight out of classic Outlook ---
    def on_drop(event):
        set_hot(False)
        paths = [p for p in root.tk.splitlist(event.data) if os.path.exists(p)] if event.data else []
        # Outlook drags carry no files; the dragged emails are Outlook's current selection
        root.after(50, start, paths or 'outlook')
        return COPY   # never MOVE, so Outlook keeps its emails

    def on_drop_enter(_):
        set_hot(True)
        return COPY

    root.drop_target_register(DND_ALL)
    root.dnd_bind('<<DropEnter>>', on_drop_enter)
    root.dnd_bind('<<DropPosition>>', lambda e: COPY)
    root.dnd_bind('<<DropLeave>>', lambda e: set_hot(False))
    root.dnd_bind('<<Drop>>', on_drop)

    # --- settings panel, slides in from the right ---
    panel = ttk.Frame(root, style='Card.TFrame', padding=(22, 18))
    panel_w = 360
    panel_state = {'open': False, 'x': 0}

    def place_panel():
        panel.place(relx=1.0, x=panel_w - panel_state['x'], y=0, relheight=1.0, width=panel_w, anchor='ne')

    def slide(target, step=0):
        panel_state['x'] = round(panel_w * ((step / 8) if target else (1 - step / 8)))
        place_panel()
        if step < 8:
            root.after(15, slide, target, step + 1)
        elif not target:
            panel.place_forget()

    def toggle_panel(_=None):
        panel_state['open'] = not panel_state['open']
        panel.lift()
        slide(panel_state['open'])

    top = ttk.Frame(panel)
    top.pack(fill='x', pady=(0, 14))
    ttk.Label(top, text='Settings', font=(font, 15, 'bold')).pack(side='left')
    ttk.Button(top, text='✕', width=3, command=toggle_panel).pack(side='right')

    def section(title):
        ttk.Label(panel, text=title, style='CardTitle.TLabel').pack(anchor='w', pady=(12, 4))

    section('Save PDFs to')
    out_label = ttk.Label(panel, style='Muted.TLabel', wraplength=panel_w - 50, justify='left')
    out_label.pack(anchor='w')
    out_btns = ttk.Frame(panel)
    out_btns.pack(anchor='w', pady=(6, 0))
    change_out = ttk.Button(out_btns, text='Change...')
    change_out.pack(side='left')
    reset_out = ttk.Button(out_btns, text='Next to each email')
    reset_out.pack(side='left', padx=(6, 0))

    section('File names')
    naming = ttk.Combobox(panel, state='readonly', values=list(NAMINGS.values()))
    naming.set(NAMINGS.get(settings['naming'], NAMINGS['file']))
    naming.pack(fill='x')

    section('Attachments')
    attach_var = tk.BooleanVar(value=settings['attachments'])
    attach = ttk.Checkbutton(panel, text='Add image and PDF attachments\nas pages after the email', variable=attach_var)
    attach.pack(anchor='w')

    section('Theme')
    theme_var = tk.StringVar(value=settings['theme'])
    for key, label in THEMES.items():
        ttk.Radiobutton(panel, text=label, value=key, variable=theme_var,
                        command=lambda: (settings.update(theme=theme_var.get()), save_settings(settings), apply_theme())
                        ).pack(anchor='w', pady=2)

    def show_out():
        out_label.configure(text=settings['out_dir'] or 'Next to each email.\nEmails dragged from Outlook go to Documents\\Email PDFs.')
        reset_out.configure(state='normal' if settings['out_dir'] else 'disabled')

    def set_out(value):
        settings['out_dir'] = value
        save_settings(settings)
        show_out()

    def on_naming(_):
        settings['naming'] = next(k for k, v in NAMINGS.items() if v == naming.get())
        save_settings(settings)

    change_out.configure(command=lambda: (d := filedialog.askdirectory(parent=root)) and set_out(str(Path(d))))
    reset_out.configure(command=lambda: set_out(''))
    naming.bind('<<ComboboxSelected>>', on_naming)
    attach.configure(command=lambda: (settings.update(attachments=attach_var.get()), save_settings(settings)))
    show_out()

    # --- setup dialog ---
    skip_file = APP_DIR / 'setup-done'

    def setup():
        dark = sv_ttk.get_theme() == 'dark'
        win = tk.Toplevel(root)
        win.title('Set up exact copy')
        win.transient(root)
        win.resizable(False, False)
        style_titlebar(win, dark)
        body = ttk.Frame(win, padding=24)
        body.pack(fill='both', expand=True)
        ttk.Label(body, text='Set up exact copy', font=(font, 15, 'bold')).pack(anchor='w')
        ttk.Label(body, style='Muted.TLabel', wraplength=500, justify='left', text=(
            'For PDFs that match exactly what Outlook prints, this PC needs Microsoft Word and classic Outlook '
            '(the Microsoft 365 / Office desktop apps), signed in to an email account once. '
            'Without them the app still works and makes a close match.')).pack(anchor='w', pady=(4, 16))
        checklist = ttk.Frame(body, style='Card.TFrame', padding=(16, 10))
        checklist.pack(fill='x')
        ttk.Label(body, style='Muted.TLabel', wraplength=500, justify='left', text=(
            'Word or Outlook missing: install the Microsoft 365 / Office desktop apps.\n'
            'Not signed in: click "Open classic Outlook", add your email account, then close Outlook. '
            'If the new Outlook opens instead, turn off its "New Outlook" switch (top right) first.')).pack(anchor='w', pady=16)

        def check():
            for w in checklist.winfo_children():
                w.destroy()
            for label, ok in setup_status():
                line = ttk.Frame(checklist)
                line.pack(fill='x', pady=3)
                ttk.Label(line, text='✔' if ok else '✖', foreground=pal['exact'] if ok else pal['failed'],
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

    def choose_emails():
        start(list(filedialog.askopenfilenames(parent=root, filetypes=[('Outlook emails', '*.msg *.eml'), ('All files', '*.*')])))

    setup_btn.configure(command=setup)
    gear.configure(command=toggle_panel)
    pick.configure(command=choose_emails)
    add_more.configure(command=choose_emails)
    folder.configure(command=lambda: start([d] if (d := filedialog.askdirectory(parent=root)) else []))
    open_dir.configure(command=lambda: last_dir and os.startfile(last_dir[0]))
    cancel_btn.configure(command=on_cancel)
    clear_btn.configure(command=clear)
    root.bind('<Escape>', lambda e: toggle_panel() if panel_state['open'] else None)
    root.protocol('WM_DELETE_WINDOW', on_close)

    apply_theme()
    show('zone')
    if not refresh() and not skip_file.exists():
        root.after(300, setup)   # first launch on a PC that isn't set up yet
    if len(sys.argv) > 1:
        root.after(200, start, sys.argv[1:])
    root.after(80, poll)
    root.after(120, spin)
    root.mainloop()


if __name__ == '__main__':
    main()
