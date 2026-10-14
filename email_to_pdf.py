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
import locale
import mimetypes
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import socket
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from email.utils import parseaddr
from urllib.parse import unquote
import winreg

import extract_msg
from bs4 import UnicodeDammit
from pypdf import PdfWriter

VERSION = '1.5.0'
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


# Light, polite Edge: no extras, below-normal CPU priority so a slow PC stays usable while converting.
EDGE_FLAGS = ['--headless', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-extensions',
              '--disable-background-networking', '--disable-sync', '--disable-component-update', '--mute-audio',
              '--disable-features=Translate,OptimizationHints,MediaRouter']
QUIET = dict(stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
             creationflags=subprocess.CREATE_NO_WINDOW | subprocess.BELOW_NORMAL_PRIORITY_CLASS)


def find_browser():
    for base in ('ProgramFiles(x86)', 'ProgramFiles', 'LOCALAPPDATA'):
        for rel in (r'Microsoft\Edge\Application\msedge.exe', r'Google\Chrome\Application\chrome.exe'):
            if os.environ.get(base) and (p := Path(os.environ[base], rel)).exists():
                return p
    raise RuntimeError('Microsoft Edge (or Google Chrome) is needed but was not found')


class EdgePrinter:
    """One headless Edge for a whole batch, driven over the DevTools protocol (a tiny stdlib WebSocket client).

    Starting Edge costs seconds on a slow PC; reusing one makes each email after the first much faster.
    It starts on first use; if anything about it fails, to_pdf falls back to a one-off Edge per email."""

    def __init__(self):
        self.proc = self.sock = None
        self.broken = False
        self.n = 0

    def _start(self):
        self.tmp = tempfile.mkdtemp(prefix='email-to-pdf-edge-')
        self.proc = subprocess.Popen([str(find_browser()), *EDGE_FLAGS, '--remote-debugging-port=0',
                                      f'--user-data-dir={self.tmp}', 'about:blank'], **QUIET)
        port_file, deadline = Path(self.tmp, 'DevToolsActivePort'), time.time() + 90   # slow PCs start slowly
        while True:
            lines = port_file.read_text().split() if port_file.exists() else []
            if len(lines) >= 2:
                break
            if self.proc.poll() is not None or time.time() > deadline:
                raise RuntimeError('Edge did not start')
            time.sleep(0.1)
        port, path = lines[:2]
        self.sock = socket.create_connection(('127.0.0.1', int(port)), timeout=180)
        key = base64.b64encode(os.urandom(16)).decode()
        self.sock.sendall(f'GET {path} HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\nUpgrade: websocket\r\nConnection: Upgrade\r\n'
                          f'Sec-WebSocket-Key: {key}\r\nSec-WebSocket-Version: 13\r\n\r\n'.encode())
        reply = b''
        while b'\r\n\r\n' not in reply:
            chunk = self.sock.recv(4096)
            if not chunk:
                raise RuntimeError('Edge closed the connection')
            reply += chunk
        if b' 101 ' not in reply.split(b'\r\n', 1)[0]:
            raise RuntimeError('Edge refused the connection')
        self.buf, self.events = reply.split(b'\r\n\r\n', 1)[1], []
        target = self.call('Target.createTarget', url='about:blank')['targetId']
        self.session = self.call('Target.attachToTarget', targetId=target, flatten=True)['sessionId']
        self.call('Page.enable', in_page=True)

    # --- minimal WebSocket framing (RFC 6455): client frames masked, server frames not ---
    def _send(self, text):
        data, head = text.encode(), bytearray([0x81])
        n = len(data)
        head += bytes([0x80 | n]) if n < 126 else bytes([0x80 | 126]) + n.to_bytes(2, 'big') if n < 65536 \
            else bytes([0x80 | 127]) + n.to_bytes(8, 'big')
        mask = os.urandom(4)
        self.sock.sendall(bytes(head) + mask + bytes(b ^ mask[i % 4] for i, b in enumerate(data)))

    def _read(self, n):
        while len(self.buf) < n:
            chunk = self.sock.recv(max(65536, n - len(self.buf)))
            if not chunk:
                raise RuntimeError('Edge closed the connection')
            self.buf += chunk
        out, self.buf = self.buf[:n], self.buf[n:]
        return out

    def _recv(self):
        message = b''
        while True:
            b0, b1 = self._read(2)
            n = b1 & 0x7F
            n = int.from_bytes(self._read(2 if n == 126 else 8), 'big') if n >= 126 else n
            payload = self._read(n)
            if b0 & 0x0F == 8:
                raise RuntimeError('Edge closed the connection')
            if b0 & 0x0F in (9, 10):   # ping / pong: not used by Edge here
                continue
            message += payload
            if b0 & 0x80:
                return json.loads(message)

    def call(self, method, in_page=False, **params):
        self.n += 1
        msg = {'id': self.n, 'method': method, 'params': params}
        if in_page:
            msg['sessionId'] = self.session
        self._send(json.dumps(msg))
        while True:
            reply = self._recv()
            if reply.get('id') == self.n:
                if 'error' in reply:
                    raise RuntimeError(reply['error'].get('message', 'Edge error'))
                return reply.get('result', {})
            self.events.append(reply.get('method'))

    def print(self, page, pdf):
        if self.proc is None:
            self._start()
        src = Path(self.tmp, f'email-{self.n}.html')
        src.write_text(page, encoding='utf-8-sig')   # BOM beats whatever charset the email's HTML declares
        try:
            self.events.clear()
            self.call('Page.navigate', in_page=True, url=src.as_uri())
            deadline = time.time() + 30   # online images may be slow; print what has loaded by then
            self.sock.settimeout(1)
            while 'Page.loadEventFired' not in self.events and time.time() < deadline:
                try:
                    self.events.append(self._recv().get('method'))
                except socket.timeout:
                    pass
            self.sock.settimeout(180)
            data = self.call('Page.printToPDF', in_page=True, printBackground=True, preferCSSPageSize=True,
                             displayHeaderFooter=False)['data']
            pdf.write_bytes(base64.b64decode(data))
        finally:
            src.unlink(missing_ok=True)

    def close(self):
        if self.proc is None:
            return
        try:
            self.call('Browser.close')
        except Exception:
            pass
        try:
            self.proc.wait(10)
        except subprocess.TimeoutExpired:
            self.proc.kill()
        if self.sock:
            self.sock.close()
        shutil.rmtree(self.tmp, ignore_errors=True)


def to_pdf(page, pdf, edge=None):
    pdf.unlink(missing_ok=True)
    if edge and not edge.broken:
        try:
            return edge.print(page, pdf)
        except Exception:
            edge.broken = True   # e.g. blocked by security software: use a one-off Edge for the rest
            pdf.unlink(missing_ok=True)
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        src = Path(tmp, 'email.html')
        src.write_text(page, encoding='utf-8-sig')
        subprocess.run([str(find_browser()), *EDGE_FLAGS, '--no-pdf-header-footer', '--timeout=30000',
                        f'--user-data-dir={Path(tmp, "profile")}', f'--print-to-pdf={pdf}', src.as_uri()],
                       timeout=300, **QUIET)
    if not pdf.exists():
        raise RuntimeError('Edge did not produce a PDF')


def append_attachments(pdf, mail, edge=None):
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
               'img{display:block;max-width:100%;max-height:9.2in;object-fit:contain}</style>' + page, shots, edge)
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


def convert(path, settings=None, printer=None, taken=None, fallback_dir=None, skip_existing=False, info=None, edge=None):
    """Returns (pdf path, 'exact' / 'close' / 'skipped', attachments added, attachments skipped).

    `info`, if given, is filled with the email's subject, sender and sent date for display."""
    path, settings = Path(path), settings or DEFAULTS
    mail = load(path)
    if info is not None:
        info.update(subject=mail['subject'], sender=mail['sender'], sent=mail['sent'])
    pdf = pdf_target(path, mail, settings, set() if taken is None else taken, fallback_dir)
    if skip_existing and pdf.exists():
        return pdf, 'skipped', 0, 0
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
        to_pdf(build_html(mail), pdf, edge)
    added, skipped = append_attachments(pdf, mail, edge) if settings['attachments'] else (0, 0)
    return pdf, mode, added, skipped


# Folders never holding the user's emails: skipping them keeps a scan of Desktop or a whole drive fast.
SKIP_DIRS = {'appdata', 'node_modules', 'windows', 'program files', 'program files (x86)', 'programdata',
             'site-packages', '_internal', '__pycache__', '.git', '.venv', 'venv'}
HIDDEN = 0x2 | 0x4   # FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM
EML_HEADER = re.compile(rb'^(from|received|subject|date|to|message-id|mime-version|return-path|delivered-to):', re.I | re.M)


def skip_dir(top, name):
    if name.startswith(('$', '.')) or name.lower() in SKIP_DIRS:
        return True
    full = os.path.join(top, name)
    try:   # hidden/system folders, and junctions (they can loop back on themselves)
        return bool(os.stat(full, follow_symlinks=False).st_file_attributes & HIDDEN) or os.path.isjunction(full)
    except OSError:   # can't read it: nothing of the user's in there
        return True


def is_email(path):
    """Checks the content, not just the name: a .msg must be an Outlook (OLE) file, an .eml must start with mail headers."""
    try:
        with open(path, 'rb') as f:
            head = f.read(4096)
    except OSError:
        return False
    if path.suffix.lower() == '.msg':
        return head.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')
    return bool(EML_HEADER.search(head))


def expand(paths, skipped=None, cancel=None, on_folder=None):
    """Email files in the given files/folders. Non-email lookalikes are counted in skipped['n'], if given.

    on_folder(n) reports progress while scanning big folders; setting `cancel` stops the scan."""
    skipped = skipped if skipped is not None else {}
    skipped.setdefault('n', 0)

    def keep(p):
        if p.suffix.lower() not in EXTS:
            return False
        if is_email(p):
            return True
        skipped['n'] += 1
        return False

    for p in map(Path, paths):
        if p.is_dir():
            found, n = [], 0
            for top, dirs, files in os.walk(p):
                n += 1
                if cancel is not None and cancel.is_set():
                    break
                if on_folder and n % 200 == 0:
                    on_folder(n)
                dirs[:] = [d for d in dirs if not skip_dir(top, d)]
                found += [f for f in (Path(top, name) for name in files) if keep(f)]
            yield from sorted(found)
        elif keep(p):
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


MAIL_CLASSES = {43, 53, 54, 55, 56, 57}   # olMail and meeting request/cancel/response items


def collect_outlook_folder(folder, subfolders, start, end, rel=Path()):
    """(entry id, store id, file stem, relative output folder) for each email in an Outlook folder."""
    # ponytail: walks every item over COM (~1 ms each); use Items.Restrict on dates if 50k+ folders get slow
    rel = rel / safe_name(folder.Name, 80)
    out, items = [], folder.Items
    for i in range(1, items.Count + 1):
        it = items.Item(i)
        if getattr(it, 'Class', 0) not in MAIL_CLASSES:
            continue
        rt = it.ReceivedTime
        day = datetime.date(rt.year, rt.month, rt.day)
        if (start and day < start) or (end and day > end):
            continue
        out.append((it.EntryID, folder.StoreID, f'{rt:%Y-%m-%d %H%M} - {safe_name(it.Subject, 100)}', rel))
    if subfolders:
        for j in range(1, folder.Folders.Count + 1):
            out += collect_outlook_folder(folder.Folders.Item(j), True, start, end, rel)
    return out


def friendly_error(e):
    """The failure in plain words plus what to do about it; the raw error is shown underneath as details."""
    text = str(e)
    if isinstance(e, PermissionError):
        return ("Couldn't save the PDF. It may be open in another program, or the folder is read-only. "
                'Close the PDF and try again, or pick another Save-to folder in Settings.')
    if isinstance(e, FileNotFoundError):
        return "The email file couldn't be found. It may have been moved, renamed or deleted."
    if isinstance(e, subprocess.TimeoutExpired):
        return 'This email took more than 2 minutes to convert, so it was stopped. Try it again on its own.'
    if 'Edge' in text and 'not found' in text:
        return 'Microsoft Edge is needed to make the PDF, but it is not installed on this PC.'
    if 'did not produce a PDF' in text:
        return "Edge couldn't turn this email into a PDF. Try again; if it keeps failing, the email may be damaged."
    if type(e).__name__ == 'com_error':
        return "Outlook couldn't open or save this email. Make sure Outlook isn't showing a dialog, then retry."
    if 'not a valid' in text.lower() or 'olefile' in text.lower() or 'invalid' in text.lower():
        return "This file doesn't look like a valid Outlook email."
    return 'Something went wrong while converting this email.'


def parse_date(text, day_first=None):
    """2026-10-14, 14/10/2026, 10/14/2026, 14 Oct 2026. Ambiguous a/b/yyyy follows the PC's date order."""
    text = text.strip()
    if not text:
        return None
    try:
        return datetime.date.fromisoformat(text)
    except ValueError:
        pass
    m = re.fullmatch(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})', text)
    if m:
        a, b, y = int(m[1]), int(m[2]), int(m[3])
        y += 2000 if y < 100 else 0
        if day_first is None:
            sample = datetime.date(2026, 12, 31).strftime('%x')
            day_first = sample.find('31') < sample.find('12')
        d, mo = (a, b) if (a > 12 or (day_first and b <= 12)) else (b, a)
        return datetime.date(y, mo, d)   # raises ValueError for impossible dates
    for fmt in ('%d %b %Y', '%d %B %Y', '%b %d %Y', '%B %d %Y'):
        try:
            return datetime.datetime.strptime(text.replace(',', ''), fmt).date()
        except ValueError:
            pass
    raise ValueError(text)


def short_date(d):
    if not isinstance(d, datetime.datetime):
        return ''
    if d.tzinfo:
        d = d.astimezone()
    return f'{d:%x}  {d:%H:%M}'


def sender_name(sender):
    name, addr = parseaddr(str(sender or '').split(';')[0])
    return name or addr


def run_batch(job, settings, ready, events, cancel):
    """Worker thread: converts the job and reports progress to the window through `events`.

    job: list of paths, 'outlook' (emails just dragged from Outlook), or a dict for a whole Outlook folder.
    """
    import pythoncom
    pythoncom.CoInitialize()
    temp = printer = None
    edge = EdgePrinter()
    ok = failed = done_before = 0
    not_email = {'n': 0}
    tasks = []   # (row label, open() -> (.msg/.eml path, output folder or None, original path or None))
    skip_existing = False
    try:
        empty = 'No .msg or .eml files found there.'
        if job == 'outlook':
            events.put(('status', 'Getting the emails from Outlook\u2026'))
            temp = tempfile.mkdtemp(prefix='email-to-pdf-')
            empty = 'Nothing came through from Outlook. Select the emails in classic Outlook, then drag them here.'
            try:
                tasks = [(f.name, lambda f=f: (f, DEFAULT_OUT, None)) for f in export_outlook_selection(temp)]
            except Exception as e:
                empty = f"Couldn't get the emails from Outlook. {friendly_error(e)}"
        elif isinstance(job, dict):   # a whole Outlook folder
            events.put(('status', 'Reading the Outlook folder\u2026'))
            temp = tempfile.mkdtemp(prefix='email-to-pdf-')
            skip_existing = job['skip']
            empty = 'No emails in that folder for the chosen dates.'
            try:
                ns = outlook()[0].Session
                top = ns.GetFolderFromID(job['entry'], job['store'])
                found = collect_outlook_folder(top, job['sub'], job['from'], job['to'])
            except Exception as e:
                found, empty = [], f"Couldn't read the Outlook folder. {friendly_error(e)}"
            base = Path(settings['out_dir'] or DEFAULT_OUT)
            if found:
                events.put(('root', str(base / found[0][3].parts[0])))

            def save(eid, sid, stem, rel):
                p = Path(temp, stem + '.msg')
                ns.GetItemFromID(eid, sid).SaveAs(str(p), 9)   # olMSGUnicode
                return p, base / rel, None

            tasks = [(f'{t[3].as_posix()}  \u203A  {t[2]}', lambda t=t: save(*t)) for t in found]
            settings = {**settings, 'out_dir': ''}   # each email goes to its mirrored folder under `base`
        else:
            events.put(('status', 'Looking for emails\u2026'))
            tasks = [(f.name, lambda f=f: (f, None, str(f))) for f in expand(job, not_email, cancel, lambda n: events.put(('status', f'Looking for emails… {n:,} folders checked')))]
            if not tasks and not_email['n']:
                empty = (f"No Outlook emails there. {not_email['n']} file{'s' * (not_email['n'] > 1)} ending in .msg/.eml "
                         "aren't emails (for example, program language files), so they were skipped.")
        if not tasks:
            events.put(('empty', empty))
            return
        events.put(('queue', [label for label, _ in tasks]))
        if ready and (not isinstance(job, list) or any(label.lower().endswith('.msg') for label, _ in tasks)):
            events.put(('status', 'Starting Outlook and Word\u2026'))
            try:
                printer = OutlookPrinter()
            except Exception:
                pass   # every file then falls back to close match, and its row says why
        taken = set()
        for i, (label, open_task) in enumerate(tasks):
            if cancel.is_set():
                break
            events.put(('working', i, f'Converting {i + 1} of {len(tasks)}: {label}'))
            src = path = None
            info = {}
            try:
                path, out_dir, src = open_task()
                pdf, mode, added, skipped = convert(path, settings, printer, taken, out_dir, skip_existing, info, edge)
                meta = (info.get('subject') or '(no subject)', sender_name(info.get('sender')), short_date(info.get('sent')))
                if mode == 'skipped':
                    events.put(('row', i, 'skipped', 'Already converted', str(pdf), src, meta))
                    done_before += 1
                    continue
                if mode == 'exact':
                    note = 'Exact copy'
                else:
                    note = 'Close match \u00B7 ' + (".eml can't be exact" if path.suffix.lower() == '.eml' else
                                                   'set up Outlook for exact' if not ready else
                                                   "Outlook didn't start" if not printer else
                                                   "Outlook couldn't print this one")
                if added:
                    note += f' \u00B7 +{added} attachment{"s" * (added > 1)}'
                if skipped:
                    note += f' ({skipped} skipped)'
                events.put(('row', i, mode, note, str(pdf), src, meta))
                ok += 1
            except Exception as e:
                failed += 1
                details = f'{friendly_error(e)}\n\nEmail: {label}\nDetails: {e}'
                meta = (info.get('subject') or label, sender_name(info.get('sender')), short_date(info.get('sent')))
                events.put(('row', i, 'failed', 'Failed \u00B7 double-click for details', details, src, meta))
            finally:
                if path and isinstance(job, dict):   # folder exports can be huge: don't keep the temp copies
                    path.unlink(missing_ok=True)
    finally:
        if printer:
            printer.close()
        edge.close()
        if temp:
            shutil.rmtree(temp, ignore_errors=True)
        if isinstance(job, dict) and job.get('quit_outlook'):
            try:
                ol = outlook()[0]
                if ol.Explorers.Count == 0:
                    ol.Quit()
            except Exception:
                pass
        if tasks:
            events.put(('finished', ok, failed, done_before, len(tasks), cancel.is_set(), not_email['n']))
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


def nice_path(p, limit=64):
    """C:\\Users\\me\\Documents\\Email PDFs -> Documents\\Email PDFs; very long paths keep their last two folders."""
    p = Path(p)
    try:
        text = str(p.relative_to(Path.home()))
    except ValueError:
        text = str(p)
    return text if len(text) <= limit else '\u2026\\' + '\\'.join(Path(text).parts[-2:])


THEMES = {'system': 'Match Windows', 'light': 'Light', 'dark': 'Dark'}
RANGES = ['All emails', 'Last 7 days', 'Last 30 days', 'Last 12 months', 'This year', 'Custom range\u2026']


def range_dates(choice, today=None):
    today = today or datetime.date.today()
    return {'Last 7 days': (today - datetime.timedelta(days=6), today),
            'Last 30 days': (today - datetime.timedelta(days=29), today),
            'Last 12 months': (today.replace(year=today.year - 1) + datetime.timedelta(days=1), today),
            'This year': (today.replace(month=1, day=1), today)}.get(choice, (None, None))


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # crisp text on high-DPI screens
    except Exception:
        pass
    try:
        locale.setlocale(locale.LC_TIME, '')   # dates in the PC's own format
    except locale.Error:
        pass
    import sv_ttk
    from tkinter import ttk
    from tkinterdnd2 import COPY, DND_ALL, TkinterDnD

    settings = load_settings()
    here = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
    root = TkinterDnD.Tk()
    root.title('Email to PDF')
    root.geometry('980x720')
    root.minsize(760, 580)
    try:
        root.iconbitmap(default=str(here / 'icon.ico'))
    except tk.TclError:
        pass
    scale = root.winfo_fpixels('1i') / 96
    size = min((18, 24, 32), key=lambda s: abs(s - 18 * scale))
    icons = {k: tk.PhotoImage(file=str(here / 'icons' / f'{k}-{size}.png'))
             for k in ('exact', 'close', 'failed', 'skipped', 'pending', 'working')}
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
    ttk.Label(head, text=f'v{VERSION}', style='Muted.TLabel').pack(side='left', padx=10, pady=(10, 0))
    gear = ttk.Button(head, text='Settings')
    gear.pack(side='right')

    # mode card
    card = ttk.Frame(page, style='Card.TFrame', padding=(18, 14))
    card.pack(fill='x', pady=(14, 14))
    dot = ttk.Label(card)
    dot.pack(side='left', padx=(0, 14))
    texts = ttk.Frame(card)
    texts.pack(side='left', fill='x', expand=True)
    mode_title = ttk.Label(texts, style='CardTitle.TLabel')
    mode_title.pack(anchor='w')
    mode_desc = ttk.Label(texts, style='Muted.TLabel')
    mode_desc.pack(anchor='w')
    setup_btn = ttk.Button(card)
    setup_btn.pack(side='right')

    content = ttk.Frame(page)
    content.pack(fill='both', expand=True)

    # --- empty state: the drop zone (click anywhere in it to choose emails) ---
    zone = tk.Canvas(content, highlightthickness=0, bd=0, cursor='hand2', takefocus=0)
    zone_state = {'mode': 'idle'}   # idle / hover / drag
    inner = tk.Frame(zone, cursor='hand2')
    z_icon = tk.Label(inner, image=logo, cursor='hand2', bd=0)
    z_icon.pack()
    z_title = tk.Label(inner, text='Drag emails here', font=(font, 16, 'bold'), cursor='hand2')
    z_title.pack(pady=(14, 2))
    z_sub = tk.Label(inner, text='.msg or .eml files, folders, or emails straight from classic Outlook',
                     font=(font, 10), cursor='hand2')
    z_sub.pack()
    z_btns = tk.Frame(inner)
    z_btns.pack(pady=(20, 0))
    pick = ttk.Button(z_btns, text='Choose emails', style='Accent.TButton')
    pick.pack(side='left')
    folder = ttk.Button(z_btns, text='Choose a folder')
    folder.pack(side='left', padx=8)
    from_ol = ttk.Button(z_btns, text='Export from Outlook')
    from_ol.pack(side='left')
    z_note = tk.Label(inner, font=(font, 10), wraplength=520, justify='center')   # shown only with a message

    def draw_zone(_=None):
        zone.delete('all')
        w, h = zone.winfo_width(), zone.winfo_height()
        if w < 50 or h < 50 or not pal:
            return
        mode = zone_state['mode']
        fill = pal['zone_hot'] if mode == 'drag' else pal['zone']
        line = {'drag': pal['accent'], 'hover': pal['dash_hover']}.get(mode, pal['dash'])
        zone.configure(bg=pal['bg'])
        for w_ in (inner, z_btns, z_icon):
            w_.configure(bg=fill)
        z_title.configure(bg=fill, fg=pal['accent'] if mode == 'drag' else pal['fg'])
        z_sub.configure(bg=fill, fg=pal['muted'])
        z_note.configure(bg=fill, fg=pal['failed'])
        z_title.configure(text='Drop to convert' if mode == 'drag' else 'Drag emails here')
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
        for cx, cy, start_ in ((x0, y0, 90), (x1 - 2 * r, y0, 0), (x0, y1 - 2 * r, 180), (x1 - 2 * r, y1 - 2 * r, 270)):
            zone.create_arc(cx, cy, cx + 2 * r, cy + 2 * r, start=start_, extent=90, style='arc', outline=line, width=2)
        zone.create_window(w / 2, h / 2, window=inner)

    def set_zone(mode):
        if zone_state['mode'] != mode:
            zone_state['mode'] = mode
            draw_zone()

    def zone_hover(e):
        if zone_state['mode'] != 'drag':
            inside = zone.winfo_containing(e.x_root, e.y_root)
            set_zone('hover' if inside is not None and str(inside).startswith(str(zone)) else 'idle')

    zone.bind('<Configure>', draw_zone)
    for w_ in (zone, inner, z_icon, z_title, z_sub, z_note):
        w_.bind('<Enter>', zone_hover)
        w_.bind('<Leave>', zone_hover)
        w_.bind('<Button-1>', lambda e: choose_emails())

    # --- results view ---
    results = ttk.Frame(content)
    bar = ttk.Frame(results)
    bar.pack(fill='x')
    add_more = ttk.Button(bar, text='Add emails', style='Accent.TButton')
    add_more.pack(side='left')
    add_ol = ttk.Button(bar, text='Export from Outlook')
    add_ol.pack(side='left', padx=(8, 0))
    cancel_btn = ttk.Button(bar, text='Cancel')
    clear_btn = ttk.Button(bar, text='Clear list')
    clear_btn.pack(side='right')
    open_dir = ttk.Button(bar, text='Open folder', state='disabled')
    open_dir.pack(side='right', padx=8)
    retry_btn = ttk.Button(bar, text='Retry failed')
    progress = ttk.Progressbar(results, mode='determinate')
    summary = ttk.Label(results, style='Muted.TLabel', text='', wraplength=round(880 * scale), justify='left')
    summary.pack(anchor='w', pady=(14, 10))
    table = ttk.Frame(results)
    table.pack(fill='both', expand=True)
    tree = ttk.Treeview(table, columns=('subject', 'sender', 'date', 'result'), show='tree headings', selectmode='browse')
    tree.column('#0', width=round(44 * scale), stretch=False, anchor='center')
    for col, label, width, stretch in (('subject', 'Email', 260, True), ('sender', 'From', 150, False),
                                       ('date', 'Date', 140, False), ('result', 'Result', 300, True)):
        tree.heading(col, text=label, anchor='w')
        tree.column(col, anchor='w', width=round(width * scale), stretch=stretch, minwidth=60)
    scroll = ttk.Scrollbar(table, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    tree.pack(side='left', fill='both', expand=True)
    scroll.pack(side='right', fill='y')
    ttk.Label(results, style='Muted.TLabel', text='Enter or double-click opens a PDF. Right-click (or Shift+F10) for more. '
              'Drop more emails anywhere in the window.').pack(anchor='w', pady=(8, 0))

    rows = {}        # tree row id -> {'pdf': path, 'src': path, 'error': text, 'file': row label}
    batch = []       # tree row ids of the running batch, in order
    state = {'busy': False, 'closing': False, 'working': None, 'job': None, 'root': None}
    events, cancel = queue.Queue(), threading.Event()

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
            accent='#4cc2ff' if dark else '#005fb8',
            dash='#737373' if dark else '#8a8a8a',            # >= 3:1 against the zone fill
            dash_hover='#a0a0a0' if dark else '#5f5f5f',
            zone='#232323' if dark else '#f3f3f3', zone_hot='#1a3346' if dark else '#e5f1fb',
            failed='#ff99a4' if dark else '#c42b1c', border='#3a3a3a' if dark else '#d9d9d9')
        # ttk styles are per theme, so (re)configure them after every switch
        style.configure('Title.TLabel', font=('Segoe UI Variable Display', 22, 'bold'))
        style.configure('Muted.TLabel', foreground=pal['muted'])
        style.configure('CardTitle.TLabel', font=(font, 11, 'bold'))
        style.configure('Done.TLabel', font=(font, 11, 'bold'), foreground=pal['fg'])
        style.configure('Error.TLabel', foreground=pal['failed'])
        style.configure('Treeview', rowheight=round(36 * scale))
        for k in range(6):   # finished rows fade in from muted to normal text
            tree.tag_configure(f'fade{k}', foreground=mix(pal['muted'], pal['fg'], k / 6))
        for tag in ('done', 'failed', 'working'):
            tree.tag_configure(tag, foreground=pal['fg'])
        for tag in ('pending', 'skipped', 'cancelled'):
            tree.tag_configure(tag, foreground=pal['muted'])
        panel_outer.configure(bg=pal['border'])
        refresh()
        draw_zone()

    # --- mode card ---
    def refresh():
        ready = all(ok for _, ok in setup_status())
        dot.configure(image=icons['exact' if ready else 'close'])
        mode_title.configure(text='Exact copy is on' if ready else 'Exact copy is off')
        mode_desc.configure(text='Outlook emails (.msg) are printed by Outlook itself. .eml files get a close match.'
                            if ready else 'PDFs closely match Outlook. Set up classic Outlook on this PC for exact copies.')
        setup_btn.configure(text='Check setup' if ready else 'Set up exact copy')
        return ready

    # --- running a batch ---
    def set_busy(busy):
        state['busy'] = busy
        for w_ in (pick, folder, from_ol, add_more, add_ol, setup_btn, clear_btn, retry_btn, change_out, attach):
            w_.configure(state='disabled' if busy else 'normal')
        naming.configure(state='disabled' if busy else 'readonly')   # settings apply to the next batch
        if busy:
            show_out(disable=True)
            cancel_btn.configure(state='normal', text='Cancel')
            cancel_btn.pack(side='left', padx=8)
            retry_btn.pack_forget()
            progress.pack(fill='x', pady=(14, 0), before=summary)
            summary.configure(style='Muted.TLabel')
        else:
            show_out()
            cancel_btn.pack_forget()
            progress.pack_forget()

    def start(job):
        if state['busy'] or not job:
            return
        z_note.pack_forget()
        cancel.clear()
        batch.clear()
        state.update(job=job, root=None)
        add_more.configure(style='Accent.TButton')
        open_dir.configure(style='TButton')
        set_busy(True)
        show('results')
        progress.configure(value=0, maximum=1)
        summary.configure(text='Starting\u2026')
        threading.Thread(target=run_batch, args=(job, dict(settings), refresh(), events, cancel), daemon=True).start()

    def fade(row, k=1):
        if not tree.exists(row):
            return
        tree.item(row, tags=(f'fade{k}',) if k < 6 else ('done',))
        if k < 6:
            root.after(40, fade, row, k + 1)

    def batch_root():
        if state['root']:
            return state['root']
        dirs = [str(Path(rows[r]['pdf']).parent) for r in batch if rows.get(r, {}).get('pdf')]
        try:
            return os.path.commonpath(dirs) if dirs else None
        except ValueError:   # PDFs on different drives
            return dirs[-1]

    def finish(ok, failed, before, total, cancelled, not_email=0):
        where = batch_root()
        parts = []
        if ok:
            parts.append(f'{ok} PDF{"s" * (ok != 1)} saved to {nice_path(where)}' if where else f'{ok} PDFs saved')
        if before:
            parts.append(f'{before} already done')
        if failed:
            parts.append(f'{failed} failed')
        if cancelled:
            parts.append(f'cancelled after {ok + failed + before} of {total}')
        if not_email:
            parts.append(f"{not_email} non-email file{'s' * (not_email > 1)} skipped")
        summary.configure(text=' \u00B7 '.join(parts) or 'Nothing was converted.', style='Done.TLabel')
        if where:
            open_dir.configure(state='normal', style='Accent.TButton')
            add_more.configure(style='TButton')
        retriable = failed and state['job'] != 'outlook'   # dragged Outlook emails only existed temporarily
        if retriable:
            retry_btn.pack(side='right', before=open_dir)
            retry_btn.configure(state='normal')

    def poll():
        try:
            while True:
                kind, *data = events.get_nowait()
                if kind == 'status':
                    summary.configure(text=data[0])
                elif kind == 'empty':
                    if not tree.get_children():
                        show('zone')
                        z_note.configure(text=data[0])
                        z_note.pack(pady=(14, 0))
                    else:
                        summary.configure(text=data[0], style='Muted.TLabel')
                elif kind == 'root':
                    state['root'] = data[0]
                elif kind == 'queue':   # all rows appear at once, newest batch on top
                    progress.configure(maximum=len(data[0]), value=0)
                    for i, name in enumerate(data[0]):
                        row = tree.insert('', i, image=icons['pending'], values=(name, '', '', 'Waiting'), tags=('pending',))
                        rows[row] = {'file': name}
                        batch.append(row)
                    tree.see(batch[0])
                elif kind == 'working':
                    i, text = data
                    state['working'] = batch[i]
                    tree.item(batch[i], image=icons['working'], tags=('working',))
                    tree.set(batch[i], 'result', 'Converting\u2026')
                    summary.configure(text=text)
                elif kind == 'row':
                    i, status, note, target, src, meta = data
                    row = batch[i]
                    state['working'] = None
                    subject, sender, date = meta
                    tree.item(row, image=icons[status], values=(subject, sender, date, note))
                    rows[row].update({'src': src, 'error' if status == 'failed' else 'pdf': target})
                    progress.configure(value=i + 1)
                    if status == 'skipped':
                        tree.item(row, tags=('skipped',))
                    else:
                        fade(row)
                elif kind == 'finished':
                    for row in batch:   # anything never reached (cancelled)
                        if tree.exists(row) and 'pdf' not in rows[row] and 'error' not in rows[row]:
                            tree.item(row, image=icons['skipped'], tags=('cancelled',))
                            tree.set(row, 'result', 'Cancelled')
                    finish(*data)
                elif kind == 'done':
                    state['working'] = None
                    set_busy(False)
                    if state['closing']:
                        root.destroy()
                        return
        except queue.Empty:
            pass
        root.after(80, poll)

    def on_cancel():
        cancel.set()
        cancel_btn.configure(state='disabled', text='Cancelling\u2026')
        summary.configure(text='Finishing the current email, then stopping\u2026')

    def on_close():
        if state['busy']:   # finish the current email and close Outlook/Word cleanly first
            state['closing'] = True
            on_cancel()
            summary.configure(text='Closing after the current email\u2026')
        else:
            root.destroy()

    def clear():
        tree.delete(*tree.get_children())
        rows.clear()
        state['root'] = None
        open_dir.configure(state='disabled', style='TButton')
        retry_btn.pack_forget()
        show('zone')

    def retry():
        job = state['job']
        if isinstance(job, dict):   # re-run the folder: finished emails are skipped, failed ones are tried again
            start({**job, 'skip': True})
        else:
            start([rows[r]['src'] for r in batch if 'error' in rows.get(r, {}) and rows[r].get('src')])

    # --- row actions: Enter / double-click, and a menu by right-click, Shift+F10 or the Menu key ---
    def open_row(row):
        info = rows.get(row, {})
        if info.get('pdf'):
            os.startfile(info['pdf'])
        elif info.get('error'):
            messagebox.showerror("Couldn't convert this email", info['error'], parent=root)

    def reveal(path):
        subprocess.Popen(['explorer', '/select,', str(Path(path))])

    menu = tk.Menu(root, tearoff=False)

    def show_menu(row, x, y):
        if not row or not ({'pdf', 'error'} & rows.get(row, {}).keys()):
            return
        tree.selection_set(row)
        tree.focus(row)
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
        menu.tk_popup(x, y)

    def menu_from_keyboard(_):
        row = tree.focus()
        box = tree.bbox(row, 'subject') if row else None
        if box:
            show_menu(row, tree.winfo_rootx() + box[0] + 20, tree.winfo_rooty() + box[1] + box[3])
        return 'break'

    tree.bind('<Double-1>', lambda e: open_row(tree.identify_row(e.y)))
    tree.bind('<Return>', lambda e: open_row(tree.focus()))
    tree.bind('<Button-3>', lambda e: show_menu(tree.identify_row(e.y), e.x_root, e.y_root))
    tree.bind('<Shift-F10>', menu_from_keyboard)
    tree.bind('<App>', menu_from_keyboard)

    # hovering an email shows its file name
    tip = {'win': None, 'row': None}

    def hide_tip(_=None):
        if tip['win']:
            tip['win'].destroy()
        tip.update(win=None, row=None)

    def on_motion(e):
        row = tree.identify_row(e.y)
        if row == tip['row']:
            return
        hide_tip()
        name = rows.get(row, {}).get('file')
        if row and name and tree.identify_column(e.x) in ('#1', '#0'):
            tip['row'] = row
            tip['win'] = w_ = tk.Toplevel(root)
            w_.wm_overrideredirect(True)
            w_.wm_geometry(f'+{e.x_root + 14}+{e.y_root + 18}')
            ttk.Label(w_, text=name, padding=(8, 4)).pack()

    tree.bind('<Motion>', on_motion)
    tree.bind('<Leave>', hide_tip)

    # --- drag and drop: files/folders from Explorer, or emails straight out of classic Outlook ---
    def on_drop(event):
        set_zone('idle')
        paths = [p for p in root.tk.splitlist(event.data) if os.path.exists(p)] if event.data else []
        # Outlook drags carry no files; the dragged emails are Outlook's current selection
        root.after(50, start, paths or 'outlook')
        return COPY   # never MOVE, so Outlook keeps its emails

    def on_drop_enter(_):
        set_zone('drag')
        return COPY

    root.drop_target_register(DND_ALL)
    root.dnd_bind('<<DropEnter>>', on_drop_enter)
    root.dnd_bind('<<DropPosition>>', lambda e: COPY)
    root.dnd_bind('<<DropLeave>>', lambda e: set_zone('idle'))
    root.dnd_bind('<<Drop>>', on_drop)

    # --- settings panel, slides in from the right ---
    panel_w = round(360 * scale)
    panel_outer = tk.Frame(root)   # its 1px left edge is the panel's border
    panel = ttk.Frame(panel_outer, style='Card.TFrame', padding=(22, 18))
    panel.pack(fill='both', expand=True, padx=(1, 0))
    panel_state = {'open': False, 'x': 0}

    def place_panel():
        panel_outer.place(relx=1.0, x=panel_w - panel_state['x'], y=0, relheight=1.0, width=panel_w, anchor='ne')

    def slide(target, step=0):
        panel_state['x'] = round(panel_w * ((step / 8) if target else (1 - step / 8)))
        place_panel()
        if step < 8:
            root.after(15, slide, target, step + 1)
        elif not target:
            panel_outer.place_forget()

    def toggle_panel(_=None):
        panel_state['open'] = not panel_state['open']
        panel_outer.lift()
        slide(panel_state['open'])
        (close_panel if panel_state['open'] else gear).focus_set()   # keyboard focus follows the panel

    top = ttk.Frame(panel)
    top.pack(fill='x', pady=(0, 14))
    ttk.Label(top, text='Settings', font=(font, 15, 'bold')).pack(side='left')
    close_panel = ttk.Button(top, text='Close', command=toggle_panel)
    close_panel.pack(side='right')

    def section(title):
        ttk.Label(panel, text=title, style='CardTitle.TLabel').pack(anchor='w', pady=(12, 4))

    section('Save PDFs to')
    out_label = ttk.Label(panel, style='Muted.TLabel', wraplength=panel_w - 50, justify='left')
    out_label.pack(anchor='w')
    out_btns = ttk.Frame(panel)
    out_btns.pack(anchor='w', pady=(6, 0))
    change_out = ttk.Button(out_btns, text='Choose a folder\u2026')
    change_out.pack(side='left')
    reset_out = ttk.Button(out_btns, text='Use next to each email')

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
    panel_note = ttk.Label(panel, style='Muted.TLabel', wraplength=panel_w - 50, justify='left')
    panel_note.pack(anchor='w', pady=(16, 0))

    def show_out(disable=False):
        if settings['out_dir']:
            out_label.configure(text=nice_path(settings['out_dir']))
            reset_out.pack(side='left', padx=(6, 0))
        else:
            out_label.configure(text='Next to each email.\nEmails from Outlook go to Documents\\Email PDFs.')
            reset_out.pack_forget()
        reset_out.configure(state='disabled' if disable else 'normal')
        panel_note.configure(text='Changes apply to the next batch.' if disable else '')

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

    def dialog(title):
        win = tk.Toplevel(root)
        win.title(title)
        win.transient(root)
        win.resizable(False, False)
        style_titlebar(win, sv_ttk.get_theme() == 'dark')
        body = ttk.Frame(win, padding=24)
        body.pack(fill='both', expand=True)
        ttk.Label(body, text=title, font=(font, 15, 'bold')).pack(anchor='w')
        return win, body

    # --- setup dialog ---
    def setup():
        win, body = dialog('Set up exact copy')
        ttk.Label(body, style='Muted.TLabel', wraplength=500, justify='left', text=(
            'For PDFs that match exactly what Outlook prints, this PC needs Microsoft Word and classic Outlook '
            '(the Microsoft 365 / Office desktop apps), signed in to an email account once. '
            'Without them the app still works and makes a close match.')).pack(anchor='w', pady=(4, 16))
        checklist = ttk.Frame(body, style='Card.TFrame', padding=(16, 10))
        checklist.pack(fill='x')
        howto = ttk.Label(body, style='Muted.TLabel', wraplength=500, justify='left', text=(
            'Word or Outlook missing: install the Microsoft 365 / Office desktop apps.\n'
            'Not signed in: click "Open classic Outlook", add your email account, then close Outlook. '
            'If the new Outlook opens instead, turn off its "New Outlook" switch (top right) first.'))
        btns = ttk.Frame(body)
        btns.pack(fill='x', side='bottom', pady=(16, 0))

        def check():
            for w_ in checklist.winfo_children():
                w_.destroy()
            results_ = setup_status()
            for label, ok in results_:
                line = ttk.Frame(checklist)
                line.pack(fill='x', pady=3)
                ttk.Label(line, image=icons['exact' if ok else 'failed']).pack(side='left', padx=(0, 10))
                ttk.Label(line, text=label).pack(side='left')
            if all(ok for _, ok in results_):
                howto.pack_forget()
            else:
                howto.pack(anchor='w', pady=(16, 0), before=btns)
            refresh()

        def open_outlook():
            try:
                os.startfile('outlook.exe')
            except OSError:
                messagebox.showerror('Email to PDF', 'Classic Outlook is not installed on this PC.', parent=win)

        ttk.Button(btns, text='Open classic Outlook', command=open_outlook).pack(side='left')
        done = ttk.Button(btns, text='Done', style='Accent.TButton', command=win.destroy)
        done.pack(side='right')
        ttk.Button(btns, text='Check again', command=check).pack(side='right', padx=8)
        win.bind('<Escape>', lambda e: win.destroy())
        check()
        win.grab_set()
        done.focus_set()

    # --- export a whole Outlook folder ---
    def outlook_folder_dialog():
        if state['busy']:
            return
        if not refresh():
            messagebox.showinfo('Email to PDF', 'Exporting from Outlook needs classic Outlook signed in on this PC.\n\n'
                                'Click "Set up exact copy" to see what is missing.', parent=root)
            return
        try:
            ol, started = outlook()
            picked = ol.Session.PickFolder()   # Outlook's own folder picker
        except Exception as e:
            messagebox.showerror('Email to PDF', f"Couldn't open Outlook. {friendly_error(e)}", parent=root)
            return

        def quit_ol():   # only an Outlook we started just for this, with no window open
            try:
                if started and ol.Explorers.Count == 0:
                    ol.Quit()
            except Exception:
                pass

        if picked is None:
            return quit_ol()
        base = Path(settings['out_dir'] or DEFAULT_OUT)
        win, body = dialog('Export from Outlook')
        box = ttk.Frame(body, style='Card.TFrame', padding=(16, 12))
        box.pack(fill='x', pady=(12, 0))
        ttk.Label(box, text=picked.Name, style='CardTitle.TLabel').pack(anchor='w')
        ttk.Label(box, style='Muted.TLabel', text=f'{picked.FolderPath.lstrip(chr(92))}  \u00B7  {picked.Items.Count} items').pack(anchor='w')
        sub_var, skip_var = tk.BooleanVar(value=False), tk.BooleanVar(value=True)
        n_subs = picked.Folders.Count
        if n_subs:
            ttk.Checkbutton(body, text=f'Include its {n_subs} subfolder{"s" * (n_subs > 1)}',
                            variable=sub_var).pack(anchor='w', pady=(14, 0))
        rng = ttk.Frame(body)
        rng.pack(anchor='w', pady=(14, 0))
        ttk.Label(rng, text='Emails received').pack(side='left')
        period = ttk.Combobox(rng, state='readonly', values=RANGES, width=18)
        period.set(RANGES[0])
        period.pack(side='left', padx=(10, 0))
        custom = ttk.Frame(body)
        ttk.Label(custom, text='From').pack(side='left')
        d_from = ttk.Entry(custom, width=13)
        d_from.pack(side='left', padx=8)
        ttk.Label(custom, text='to').pack(side='left')
        d_to = ttk.Entry(custom, width=13)
        d_to.pack(side='left', padx=8)
        hint = ttk.Label(body, style='Muted.TLabel', text=f'For example {datetime.date.today():%x}. Leave "to" empty for up to today.')
        err = ttk.Label(body, style='Error.TLabel')
        after_dates = ttk.Frame(body)
        after_dates.pack(fill='x')
        ttk.Checkbutton(after_dates, text='Skip emails that already have a PDF (handy for running it again later)',
                        variable=skip_var).pack(anchor='w', pady=(14, 0))
        ttk.Label(after_dates, style='Muted.TLabel', wraplength=480, justify='left',
                  text=f'PDFs go to {nice_path(base / safe_name(picked.Name, 80))}, with subfolders matching Outlook.'
                  ).pack(anchor='w', pady=(12, 0))

        def on_period(_=None):
            if period.get() == RANGES[-1]:
                custom.pack(anchor='w', pady=(10, 0), before=after_dates)
                hint.pack(anchor='w', pady=(4, 0), before=after_dates)
                d_from.focus_set()
            else:
                for w_ in (custom, hint, err):
                    w_.pack_forget()

        period.bind('<<ComboboxSelected>>', on_period)

        def go():
            for e_ in (d_from, d_to):
                e_.state(['!invalid'])
            if period.get() == RANGES[-1]:
                first = last = None
                for e_ in (d_from, d_to):
                    try:
                        value = parse_date(e_.get())
                    except ValueError:
                        e_.state(['invalid'])
                        err.configure(text=f"\"{e_.get().strip()}\" isn't a date. Try {datetime.date.today():%x}.")
                        err.pack(anchor='w', pady=(4, 0), before=after_dates)
                        return e_.focus_set()
                    first, last = (value, last) if e_ is d_from else (first, value)
                if first and last and first > last:
                    d_from.state(['invalid'])
                    err.configure(text='The "from" date is after the "to" date.')
                    err.pack(anchor='w', pady=(4, 0), before=after_dates)
                    return d_from.focus_set()
            else:
                first, last = range_dates(period.get())
            job = dict(entry=picked.EntryID, store=picked.StoreID, sub=sub_var.get(), skip=skip_var.get(),
                       quit_outlook=started, **{'from': first, 'to': last})
            win.destroy()
            start(job)

        def cancel_dialog():
            win.destroy()
            quit_ol()

        btns = ttk.Frame(body)
        btns.pack(fill='x', pady=(18, 0))
        ttk.Button(btns, text='Convert', style='Accent.TButton', command=go).pack(side='right')
        ttk.Button(btns, text='Cancel', command=cancel_dialog).pack(side='right', padx=8)
        win.protocol('WM_DELETE_WINDOW', cancel_dialog)
        # Enter presses the focused button, otherwise Convert; Esc cancels
        win.bind('<Return>', lambda e: e.widget.invoke() if isinstance(e.widget, ttk.Button) else go())
        win.bind('<Escape>', lambda e: cancel_dialog())
        win.grab_set()
        period.focus_set()

    def choose_emails():
        if not state['busy']:
            start(list(filedialog.askopenfilenames(parent=root, filetypes=[('Outlook emails', '*.msg *.eml'), ('All files', '*.*')])))

    setup_btn.configure(command=setup)
    gear.configure(command=toggle_panel)
    pick.configure(command=choose_emails)
    add_more.configure(command=choose_emails)
    from_ol.configure(command=outlook_folder_dialog)
    add_ol.configure(command=outlook_folder_dialog)
    folder.configure(command=lambda: start([d] if (d := filedialog.askdirectory(parent=root)) else []))
    open_dir.configure(command=lambda: (where := batch_root()) and os.startfile(where))
    retry_btn.configure(command=retry)
    cancel_btn.configure(command=on_cancel)
    clear_btn.configure(command=clear)
    root.bind('<Escape>', lambda e: toggle_panel() if panel_state['open'] else None)
    root.bind('<Control-o>', lambda e: choose_emails())
    root.protocol('WM_DELETE_WINDOW', on_close)

    apply_theme()
    show('zone')
    if len(sys.argv) > 1:
        root.after(200, start, sys.argv[1:])
    root.after(80, poll)
    try:   # the packaged app shows a splash while it unpacks; the window is ready now
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass
    root.mainloop()


if __name__ == '__main__':
    main()
