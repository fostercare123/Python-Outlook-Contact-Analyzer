# Outlook Contact Analyzer

Extract unique email addresses from Outlook and save them to Excel with country lookup by TLD and optional WHOIS registrant info. Includes email send dates and engagement frequency.

## Preview Output

`extracted_contacts.xlsx` will look like this:

| Email Address | Country | Last Email Date | Count |
|---|---|---|---:|
| customer@novonordisk.com | Denmark | 2025-11-04 | 14 |
| support@firma.de | Germany | 2024-08-19 | 6 |
| sales@company.fr | France | 2023-03-12 | 3 |
| info@example.net | Unknown | 2022-09-01 | 1 |

## Quick start

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Mode 1: PST (no login)

**What is PST?** PST (Personal Storage Table) is Outlook's file format for storing emails, calendar items, and contacts. You can export your mailbox to a `.pst` file from Outlook: File → Open & Export → Import/Export → Export to a file → Outlook Data File (.pst).

Use this mode if you already have a PST file. You do not need `.env` for this mode.

In `config.py`:

```python
USE_PST = True
PST_PATH = r"C:\path\to\mailbox.pst"
PST_FOLDERS = []  # Empty = scan all folders, or set ["Inbox"]
```

## Mode 2: IMAP (login)

Use this mode if you want to connect directly to your mailbox.

Create or edit `.env`:

```
EMAIL_USER=your-email@outlook.com
EMAIL_PASS=your-app-password-here
```

In `config.py`:

```python
USE_PST = False
```

## Run

```bash
python main.py
```

Output: `extracted_contacts.xlsx` with columns:
- **Email Address**
- **Country** (detected by TLD, or WHOIS for .com/.net domains; codes mapped to full names)
- **Last Email Date** (when the email was sent/received)
- **Count** (how many times the email appears in all folders)

**Sorted by:** Country (A-Z), then by Count (most frequent first)

## WHOIS Caching

WHOIS lookups are cached locally in `whois_cache.json` to avoid redundant network requests. The first run performs all WHOIS queries (slower), but subsequent runs reuse the cache (much faster).

- **First run:** WHOIS queries all unknown domains, saves results to `whois_cache.json`
- **Subsequent runs:** Loads cache automatically, only queries new domains
- **To reset cache:** Delete `whois_cache.json` and re-run to refresh all lookups

This cache persists across runs, so you only pay the WHOIS performance cost once.

## Summary

The script prints statistics after completion:
```
Total unique emails: 1,234
Total emails found:  3,456
Countries:           42
Unknown country:     156
```

## Features

- Scans **all emails** in the PST (no time limit)
- Detects country by TLD (e.g., .de = Germany)
- For unknown TLDs (.com, .net, etc.), attempts WHOIS registrant lookup
- WHOIS country codes automatically mapped to full names (e.g., DK → Denmark)
- Tracks email frequency (engagement metric)
- Real-time progress bar during country detection
- Deduplicates and exports to Excel
