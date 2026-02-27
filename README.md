# Python Email Extractor

Extract unique email addresses from Outlook and save them to Excel with a simple country lookup by TLD.

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

Output: `extracted_contacts.xlsx` (sorted by Country, then Email Address)
