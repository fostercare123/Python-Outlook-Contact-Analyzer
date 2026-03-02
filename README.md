# Python Email Extractor

Extract unique email addresses from Outlook and save them to Excel with country lookup by TLD and optional WHOIS registrant info. Includes email send dates.

## Quick start

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Configure

Edit `.env` with your credentials (used only for IMAP mode):

```
EMAIL_USER=your-email@outlook.com
EMAIL_PASS=your-app-password-here
```

## PST mode (no login)

Use this if you already exported a .pst file from Outlook.

In `config_classic.py`:

```python
USE_PST = True
PST_PATH = r"C:\path\to\mailbox.pst"
PST_FOLDERS = []  # Empty = scan all folders, or set ["Inbox"]
```

## IMAP mode (login)

In `config_classic.py`:

```python
USE_PST = False
```

## Run

```bash
python run_classic.py
```

Output: `extracted_contacts.xlsx` with columns:
- **Email Address**
- **Country** (detected by TLD, or WHOIS for .com/.net domains)
- **Last Email Date** (when the email was sent/received)

**Sorted by:** Country, then Email Address, then Date

## Features

- Scans **all emails** in the PST (no time limit)
- Detects country by TLD (e.g., .de = Germany)
- For unknown TLDs (.com, .net, etc.), attempts WHOIS lookup on the domain registrant
- Extracts the date each email was sent
- Deduplicates and exports to Excel
