# Python Email Extractor

Extract customer email addresses from your Outlook mailbox and automatically identify their countries based on email domain TLDs. Perfect for analyzing customer geographic distribution.

## Features

- ✉️ **Connects via IMAP** to Outlook/Microsoft 365 mailboxes
- 🚀 **Efficient Header-Only Scanning** - fetches only headers, not full email bodies
- 🌍 **Automatic Country Detection** - identifies 30+ countries based on domain TLDs
- 🔄 **Automatic Deduplication** - removes duplicate email addresses
- 📊 **Excel Export** - organized by country, sorted alphabetically
- 📈 **Progress Tracking** - real-time scan progress display

## Requirements

- Python 3.7 or higher
- Access to Outlook/Microsoft 365 account with IMAP enabled
- An app password (for Office 365 accounts with 2-factor authentication)

## Installation

### 1. Clone or download this project
```bash
git clone <repository-url>
cd Python-Email-Extractor
```

### 2. Create a virtual environment (recommended)
```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

## Configuration

### 1. Generate an Outlook App Password

If you use Office 365 with 2-factor authentication:
1. Go to https://account.microsoft.com/account/manage-my-microsoft-account
2. Select **Security** → **Advanced security options**
3. Create an **App password** for "Mail"
4. Microsoft will generate a 16-character password

### 2. Update `config_classic.py`

Edit `config_classic.py` with your credentials:

```python
EMAIL_USER = 'your-email@outlook.com'
EMAIL_PASS = 'your-app-password-here'
```

**Security Note:** Never commit `config_classic.py` with real credentials to version control. Consider using environment variables or `.gitignore` for sensitive data.

## PST (No Authentication) Mode

If you already exported a local Outlook Data File (.pst), you can scan it without any login:

1. Install the Windows dependency:
```bash
pip install pywin32
```

2. Update `config_classic.py`:
```python
USE_PST = True
PST_PATH = r"C:\path\to\mailbox.pst"
PST_FOLDERS = []  # Optional: e.g. ["Inbox", "Sent Items"]
```

3. Run the script:
```bash
python run_classic.py
```

Notes:
- This uses Outlook on Windows via COM and does not require email authentication.
- Leave `PST_FOLDERS` empty to scan all folders in the PST.

## Usage

Simply run the script:

```bash
python run_classic.py
```

The script will:
1. Connect to your Outlook inbox
2. Scan all email headers
3. Extract unique email addresses
4. Detect country based on domain TLD
5. Save results to `extracted_contacts.xlsx`

### Example Output

```
Starting scan of 1250 emails...
Progress: [1250/1250]
Scan complete.
Success! 487 addresses saved to extracted_contacts.xlsx
```

## Output Format

The Excel file (`extracted_contacts.xlsx`) contains two columns:

| Email Address | Country |
|---|---|
| customer@company.de | Germany |
| info@example.fr | France |
| support@startup.com | Unknown |
| hello@company.uk | United Kingdom |

**Sorted by:** Country first, then Email Address (alphabetically)

## Country Detection

The script maps email domain TLDs to 30+ countries:

### Supported Country Codes
- **Europe:** DE, UK, FR, NL, BE, SE, NO, DK, FI, ES, IT, CH, AT, PL, CZ, HU, PT, GR, IE
- **North America:** US, CA, MX
- **Asia-Pacific:** AU, NZ, JP, CN, IN, TH, SG, HK, KR
- **Other:** BR, AR, ZA, AE, IL

**Generic TLDs** (.com, .org, .net, .edu, .gov, .biz, .info, .co) are marked as "Unknown" since they don't indicate a specific country.

## Troubleshooting

### "Login failed" error
- Check that email and password are correct in `config.py`
- Verify IMAP is enabled in your Outlook settings
- For Office 365, ensure you're using an app password, not your regular password

### No emails found
- Verify your inbox contains emails
- Check that the correct mailbox folder is selected (currently set to "inbox")
- To scan other folders, edit `mail.select("inbox")` in the code

### Missing dependencies
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

## Customization

### Scan Different Folder
Edit line 137 in `IMAP-extractor.py`:
```python
mail.select("inbox")  # Change to "Sent Items", "Archive", etc.
```

### Add More Countries
Edit the `COUNTRY_TLD_MAP` dictionary at the top of `IMAP-extractor.py`:
```python
'nz': 'New Zealand',  # Example entry
'tw': 'Taiwan',       # Add new countries here
```

## How It Works

1. **IMAP Connection:** Connects securely to Outlook's IMAP server (outlook.office365.com)
2. **Header Scanning:** Fetches only email headers (RFC822 format) to minimize bandwidth
3. **Email Extraction:** Uses regex pattern matching to find all email addresses in From, To, and Cc headers
4. **Deduplication:** Uses a Python set to automatically remove duplicates
5. **Country Detection:** Extracts the TLD from each domain and maps it to a country
6. **Excel Generation:** Creates a sorted, organized spreadsheet with pandas and openpyxl
