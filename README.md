# Email to PDF

Turns Outlook emails (`.msg` and `.eml`) into PDFs that look like they do in Outlook. Windows 10/11.

**Download** from the [Releases](../../releases) page:
- `Email-to-PDF-Setup.exe`: installs the app with Start menu and desktop shortcuts (no admin rights needed).
- `Email to PDF.exe`: the app on its own, no install.

## Using it

- Drag emails or folders into the window, including straight out of classic Outlook, or click the drop zone, **Choose emails** / **Choose a folder** (Ctrl+O also works), or drop them onto the app icon.
- **Export from Outlook** (needs classic Outlook signed in): pick any Outlook folder, optionally with its subfolders and a date range (All, Last 7/30 days, Last 12 months, This year, or custom dates in your usual format). PDFs go to `Documents\Email PDFs\<folder>` (or your Save-to folder) with subfolders matching Outlook. Emails that already have a PDF are skipped, so re-running only converts new mail.
- In **Settings** (top right): **Save PDFs to**: next to each email (default) or a folder you pick. Emails dragged from Outlook go to `Documents\Email PDFs` unless you pick a folder.
- **File names**: same as the email file, *Date - Subject*, or *Subject*.
- **Attachments**: image and PDF attachments are added as pages after the email (can be turned off). Other attachments are listed by name.
- Converting runs in the background and can be cancelled. Each row shows the email's subject, sender and date, with an icon for the result. Enter or double-click opens a PDF; right-click (or Shift+F10) for Show in folder, Open original email and Copy PDF path. When a batch finishes you get where the PDFs went, **Open folder**, and **Retry failed** if anything failed.
- Theme: Light, Dark, or match Windows (in Settings).

Settings are remembered in `%APPDATA%\Email to PDF`.

To get email files out of Outlook without the app, drag them from Outlook into a folder (classic Outlook makes `.msg` files) or use *Save as* (new Outlook / Outlook on the web make `.eml` files).

## Two modes

| Mode | When | How |
|---|---|---|
| **Exact copy** | `.msg` files on a PC with classic Outlook + Word, signed in to an email account once | Outlook and Word print the email themselves, exactly as Outlook prints it |
| **Close match** | everything else (`.eml` files, or no classic Outlook) | The email's own HTML is printed by Microsoft Edge, the engine new Outlook and Outlook on the web use |

On first launch the app checks the PC and opens a short setup checklist if exact copy isn't available. It can be reopened with the **Set up** button.

Close-match limits: Bcc isn't shown, old rich-text (RTF-only) emails lose formatting, and online images are downloaded during conversion.

## First run warning

The `.exe` isn't code-signed, so Windows shows "Windows protected your PC". Click **More info → Run anyway**.

## Building it yourself

Needs Python 3.10+ and Microsoft Edge. Run:

```
build.bat
```

It installs the dependencies into `.venv`, runs `test_email_to_pdf.py`, and writes `dist\Email to PDF.exe`. If [Inno Setup 6](https://jrsoftware.org/isinfo.php) is installed it also writes `dist\Email-to-PDF-Setup.exe`.

To release: bump `VERSION` in `email_to_pdf.py`, then push a matching tag like `v1.2.0`. GitHub builds both files and attaches them to a release (see `.github/workflows/build.yml`).

## License

Uses [extract-msg](https://github.com/TeamMsgExtractor/msg-extractor), which is GPL-3.0, so this app is distributed under GPL-3.0 too.
