"""Self-check: builds a sample .eml (HTML + inline image + attachment) and converts it. Run: python test_email_to_pdf.py"""
import struct
import tempfile
import zlib
from email.message import EmailMessage
from email.utils import formatdate
from pathlib import Path

from email_to_pdf import build_html, convert, load_eml, setup_status


def png(w, h, rgb):
    chunk = lambda t, d: struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d))
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))


msg = EmailMessage()
msg['From'] = 'Ana Café <ana@example.com>'
msg['To'] = 'Team <team@example.com>'
msg['Cc'] = 'Boss <boss@example.com>'
msg['Subject'] = 'Q3 numbers — café ☕'
msg['Date'] = formatdate(1791565500, localtime=True)
msg.set_content('plain fallback')
msg.add_alternative('<!DOCTYPE html><html><head><meta charset="windows-1252"></head><body style="font-family:Georgia">'
                    '<img src="cid:logo1@x"><h2 style="color:#c0392b">Quarterly numbers — 日本語</h2>'
                    '<table style="border-collapse:collapse"><tr style="background:#2c3e50;color:#fff"><th>Region</th><th>Revenue</th></tr>'
                    '<tr style="background:#ecf0f1"><td>North</td><td>$1.2M</td></tr></table></body></html>', subtype='html')
msg.get_payload()[1].add_related(png(160, 40, (70, 130, 180)), 'image', 'png', cid='<logo1@x>')
msg.add_attachment(b'%PDF-1.4 fake', maintype='application', subtype='pdf', filename='report.pdf')

with tempfile.TemporaryDirectory() as tmp:
    eml = Path(tmp, 'sample.eml')
    eml.write_bytes(bytes(msg))
    page = build_html(load_eml(eml))
    assert page.startswith('<!DOCTYPE html>'), 'doctype must stay first'
    assert 'cid:' not in page and 'data:image/png;base64,' in page, 'inline image not embedded'
    assert 'Attachments:' in page and 'report.pdf' in page and 'logo1' not in page, 'attachment list wrong'
    assert 'Ana Café &lt;ana@example.com&gt;' in page and 'café ☕' in page, 'header missing'
    pdf, mode = convert(eml)
    assert mode == 'close' and pdf.read_bytes()[:5] == b'%PDF-' and pdf.stat().st_size > 5000, 'no real PDF produced'
    assert len(setup_status()) == 3, 'setup checklist broken'
print('ok')
