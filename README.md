# Email to PDF

Turns Outlook emails (`.msg` and `.eml`) into PDFs that look like they do in Outlook. Windows 10/11, one `.exe`, no install.

**Download:** grab `Email to PDF.exe` from the [Releases](../../releases) page.

## Using it

Open the app and click **Choose emails** or **Choose folder**, or drop email files or folders onto the app icon. Each PDF is saved next to its email. Double-click a row in the results list to open its PDF.

To get emails out of Outlook, drag them from Outlook into a folder (classic Outlook makes `.msg` files) or use *Save as* (new Outlook / Outlook on the web make `.eml` files).

## Two modes

| Mode | When | How |
|---|---|---|
| **Exact copy** | `.msg` files on a PC with classic Outlook + Word, signed in to an email account once | Outlook and Word print the email themselves, exactly as Outlook prints it |
| **Close match** | everything else (`.eml` files, or no classic Outlook) | The email's own HTML is printed by Microsoft Edge, the engine new Outlook and Outlook on the web use |

On first launch the app checks the PC and opens a short setup checklist if exact copy isn't available. It can be reopened with the **Set up** button.

Close-match limits: attachments are listed by name but not included, Bcc isn't shown, old rich-text (RTF-only) emails lose formatting, and online images are downloaded during conversion.

## First run warning

The `.exe` isn't code-signed, so Windows shows "Windows protected your PC". Click **More info → Run anyway**.

## Building it yourself

Needs Python 3.10+ and Microsoft Edge. Run:

```
build.bat
```

It installs the dependencies into `.venv`, runs `test_email_to_pdf.py`, and writes `dist\Email to PDF.exe`.

Pushing a tag like `v1.0` builds the `.exe` on GitHub and attaches it to a release automatically (see `.github/workflows/build.yml`).

## License

Uses [extract-msg](https://github.com/TeamMsgExtractor/msg-extractor), which is GPL-3.0, so this app is distributed under GPL-3.0 too.
