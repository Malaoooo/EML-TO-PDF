"""Self-check: builds a sample .eml (HTML + inline image + attachments) and converts it. Run: python test_email_to_pdf.py"""
import datetime
import io
import struct
import tempfile
import zlib
from email.message import EmailMessage
from email.utils import formatdate
from pathlib import Path

from pypdf import PdfReader, PdfWriter

from email_to_pdf import build_html, convert, load_eml, parse_date, range_dates, setup_status


def png(w, h, rgb):
    chunk = lambda t, d: struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d))
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))


def two_page_pdf():
    w = PdfWriter()
    w.add_blank_page(612, 792)
    w.add_blank_page(612, 792)
    buf = io.BytesIO()
    w.write(buf)
    return buf.getvalue()


msg = EmailMessage()
msg['From'] = 'Ana Café <ana@example.com>'
msg['To'] = 'Team <team@example.com>'
msg['Cc'] = 'Boss <boss@example.com>'
msg['Subject'] = 'Q3 numbers: café ☕'
msg['Date'] = formatdate(1791565500, localtime=True)
msg.set_content('plain fallback')
msg.add_alternative('<!DOCTYPE html><html><head><meta charset="windows-1252"></head><body style="font-family:Georgia">'
                    '<img src="cid:logo1@x"><h2 style="color:#c0392b">Quarterly numbers — 日本語</h2>'
                    '<table style="border-collapse:collapse"><tr style="background:#2c3e50;color:#fff"><th>Region</th><th>Revenue</th></tr>'
                    '<tr style="background:#ecf0f1"><td>North</td><td>$1.2M</td></tr></table></body></html>', subtype='html')
msg.get_payload()[1].add_related(png(160, 40, (70, 130, 180)), 'image', 'png', cid='<logo1@x>')
msg.add_attachment(two_page_pdf(), maintype='application', subtype='pdf', filename='report.pdf')
msg.add_attachment(png(800, 600, (200, 60, 60)), maintype='image', subtype='png', filename='photo.png')
msg.add_attachment(b'%PDF-1.4 broken', maintype='application', subtype='pdf', filename='broken.pdf')
msg.add_attachment(b'a,b\n1,2', maintype='text', subtype='csv', filename='data.csv')

with tempfile.TemporaryDirectory() as tmp:
    eml = Path(tmp, 'sample.eml')
    eml.write_bytes(bytes(msg))
    page = build_html(load_eml(eml))
    assert page.startswith('<!DOCTYPE html>'), 'doctype must stay first'
    assert 'cid:' not in page and 'data:image/png;base64,' in page, 'inline image not embedded'
    assert 'report.pdf; photo.png; broken.pdf; data.csv' in page and 'logo1' not in page, 'attachment list wrong'
    assert 'Ana Café &lt;ana@example.com&gt;' in page and 'café ☕' in page, 'header missing'

    plain, mode, added, skipped = convert(eml, {'out_dir': '', 'naming': 'file', 'attachments': False})
    assert mode == 'close' and plain == eml.with_suffix('.pdf') and plain.read_bytes()[:5] == b'%PDF-', 'no real PDF produced'
    base_pages = len(PdfReader(plain).pages)

    out = Path(tmp, 'out')
    pdf, mode, added, skipped = convert(eml, {'out_dir': str(out), 'naming': 'date-subject', 'attachments': True})
    assert pdf.parent == out and pdf.name.endswith(' - Q3 numbers_ café ☕.pdf') and pdf.name[:4].isdigit(), pdf.name
    # photo.png -> 1 page, report.pdf -> 2 pages; broken.pdf skipped; data.csv only listed
    assert (added, skipped) == (2, 1), (added, skipped)
    assert len(PdfReader(pdf).pages) == base_pages + 3, 'attachment pages missing'
    assert not list(out.glob('*.tmp.pdf')), 'temp files left behind'
    assert len(setup_status()) == 3, 'setup checklist broken'

# dates typed into "Export from Outlook"
d = datetime.date
assert parse_date('2026-10-14') == d(2026, 10, 14) and parse_date('') is None
assert parse_date('14/10/2026') == d(2026, 10, 14) and parse_date('10/14/2026') == d(2026, 10, 14), 'unambiguous day/month'
assert parse_date('03/04/2026', day_first=True) == d(2026, 4, 3) and parse_date('03/04/2026', day_first=False) == d(2026, 3, 4)
assert parse_date('14 Oct 2026') == d(2026, 10, 14) and parse_date('5/1/26', day_first=True) == d(2026, 1, 5)
for bad in ('31/31/2026', 'yesterday', '2026-13-01'):
    try:
        parse_date(bad)
        raise AssertionError(f'{bad!r} should be rejected')
    except ValueError:
        pass
assert range_dates('Last 7 days', d(2026, 10, 14)) == (d(2026, 10, 8), d(2026, 10, 14))
assert range_dates('This year', d(2026, 10, 14)) == (d(2026, 1, 1), d(2026, 10, 14))
assert range_dates('All emails') == (None, None)
print('ok')
