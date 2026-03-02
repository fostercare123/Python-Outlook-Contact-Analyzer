"""
Outlook Contact Analyzer
Extracts unique email addresses from Outlook and analyzes geographic distribution.
Author: Vasilije Niko Nikolic
"""

import imaplib
import email
import os
import re
import pandas as pd
import config  # Local file for secrets
from datetime import datetime
from collections import defaultdict

try:
    from tqdm import tqdm
except Exception:
    tqdm = None

try:
    import win32com.client  # Requires pywin32 on Windows
except Exception:
    win32com = None

try:
    import whois
except Exception:
    whois = None

# Configuration constants
IMAP_SERVER = 'outlook.office365.com'

# Detailed Email Regex using VERBOSE mode for clarity
EMAIL_REGEX = re.compile(r"""
    [a-zA-Z0-9._%+-]+    # Local part: Letters, numbers, and common symbols
    @                    # Separator: The literal @ symbol
    [a-zA-Z0-9.-]+       # Domain: The provider (e.g., outlook, gmail)
    \.                   # Dot: The literal period before the TLD
    [a-zA-Z]{2,}         # TLD: The extension (e.g., com, net, org, edu)
    """, re.VERBOSE)

# Country mapping based on TLD and domain patterns
COUNTRY_TLD_MAP = {
    # European countries
    'de': 'Germany',
    'uk': 'United Kingdom',
    'fr': 'France',
    'nl': 'Netherlands',
    'be': 'Belgium',
    'se': 'Sweden',
    'no': 'Norway',
    'da': 'Denmark',
    'dk': 'Denmark',
    'fi': 'Finland',
    'es': 'Spain',
    'it': 'Italy',
    'ch': 'Switzerland',
    'at': 'Austria',
    'pl': 'Poland',
    'cz': 'Czech Republic',
    'hu': 'Hungary',
    'pt': 'Portugal',
    'gr': 'Greece',
    'ie': 'Ireland',
    # North America
    'ca': 'Canada',
    'mx': 'Mexico',
    'us': 'United States',
    # Asia Pacific
    'au': 'Australia',
    'nz': 'New Zealand',
    'jp': 'Japan',
    'cn': 'China',
    'in': 'India',
    'th': 'Thailand',
    'sg': 'Singapore',
    'hk': 'Hong Kong',
    'kr': 'South Korea',
    # Other regions
    'br': 'Brazil',
    'ar': 'Argentina',
    'za': 'South Africa',
    'ae': 'United Arab Emirates',
    'il': 'Israel',
}

# Country code to full name mapping (for WHOIS results)
COUNTRY_CODE_MAP = {
    'DK': 'Denmark',
    'DE': 'Germany',
    'SE': 'Sweden',
    'NO': 'Norway',
    'NL': 'Netherlands',
    'BE': 'Belgium',
    'FR': 'France',
    'IT': 'Italy',
    'ES': 'Spain',
    'PT': 'Portugal',
    'AT': 'Austria',
    'CH': 'Switzerland',
    'PL': 'Poland',
    'CZ': 'Czech Republic',
    'HU': 'Hungary',
    'GR': 'Greece',
    'IE': 'Ireland',
    'FI': 'Finland',
    'GB': 'United Kingdom',
    'UK': 'United Kingdom',
    'US': 'United States',
    'CA': 'Canada',
    'MX': 'Mexico',
    'BR': 'Brazil',
    'AR': 'Argentina',
    'AU': 'Australia',
    'NZ': 'New Zealand',
    'JP': 'Japan',
    'CN': 'China',
    'IN': 'India',
    'SG': 'Singapore',
    'HK': 'Hong Kong',
    'TH': 'Thailand',
    'KR': 'South Korea',
    'ZA': 'South Africa',
    'AE': 'United Arab Emirates',
    'IL': 'Israel',
}

# Common generic TLDs (map to 'Unknown' as they don't indicate country)
GENERIC_TLDS = {'com', 'org', 'net', 'edu', 'gov', 'biz', 'info', 'co'}


def detect_country(email_address):
    """
    Detects the country based on the email domain's TLD (top-level domain).
    
    Args:
        email_address (str): The email address to analyze (e.g., user@company.de)
    
    Returns:
        str: The country name if detected, otherwise 'Unknown'.
    """
    try:
        # Extract the domain part (everything after @)
        domain = email_address.split('@')[1]
        
        # Extract the TLD (last part after the final dot)
        tld = domain.split('.')[-1].lower()
        
        # Check if TLD exists in our country mapping
        if tld in COUNTRY_TLD_MAP:
            return COUNTRY_TLD_MAP[tld]
        
        # Generic TLDs don't indicate a specific country
        if tld in GENERIC_TLDS:
            return 'Unknown'
        
        # For any unmapped TLD, return it as unknown
        return 'Unknown'
    except Exception:
        # If parsing fails, return unknown
        return 'Unknown'


def whois_lookup_country(domain):
    """
    Attempts to detect country via WHOIS registrant info for generic TLDs.
    
    Args:
        domain (str): Domain name (e.g., "novonordisk.com")
    
    Returns:
        str: Country name or code if found, otherwise None.
    """
    if not whois:
        return None
    
    try:
        w = whois.whois(domain)
        registrant_country = getattr(w, "registrant_country", None)
        if registrant_country:
            # Try to map the code to full name
            country_upper = registrant_country.upper().strip()
            return COUNTRY_CODE_MAP.get(country_upper, registrant_country)
    except Exception:
        pass
    
    return None


def detect_country_with_whois(email_address):
    """
    Detects country: first by TLD, then by WHOIS for generic TLDs.
    
    Args:
        email_address (str): Email address (e.g., "user@novonordisk.com")
    
    Returns:
        tuple: (country_name, source) where source is "TLD" or "WHOIS"
    """
    country = detect_country(email_address)
    source = "TLD"
    
    # If unknown, try WHOIS on generic TLDs
    if country == "Unknown":
        try:
            domain = email_address.split("@")[1]
            whois_country = whois_lookup_country(domain)
            if whois_country:
                country = whois_country
                source = "WHOIS"
        except Exception:
            pass
    
    return (country, source)


def extract_emails():
    """
    Connects to Outlook, scans headers, and collects unique email addresses
    with their detected countries.
    
    Returns:
        list: A list of tuples (email, country) found in the mailbox.
    """
    try:
        # Establish secure connection to the IMAP server
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(config.EMAIL_USER, config.EMAIL_PASS)
        # Select the 'inbox' folder; use "Sent Items" for sent mail
        mail.select("inbox")
    except Exception as e:
        print(f"Login failed: {e}")
        return []

    # Search for all message IDs in the selected folder
    _, data = mail.search(None, 'ALL')
    mail_ids = data[0].split()
    total_emails = len(mail_ids)
    
    # Use a set to automatically handle deduplication of email addresses
    found_addresses = set()

    print(f"Starting scan of {total_emails} emails...")

    for index, num in enumerate(mail_ids):
        # Update the terminal line with current progress
        if (index + 1) % 10 == 0 or (index + 1) == total_emails:
            print(f"Progress: [{index + 1}/{total_emails}]", end='\r')

        # Fetch only the header data to minimize bandwidth and time
        _, msg_data = mail.fetch(num, '(RFC822.HEADER)')
        
        for response_part in msg_data:
            # Check if the part contains the actual header bytes
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                
                # Iterate through key headers where addresses are found
                for header in ['From', 'To', 'Cc']:
                    header_value = msg.get(header, '')
                    # Find all matches within the header string
                    matches = re.findall(EMAIL_REGEX, header_value)
                    found_addresses.update(matches)

    print("\nScan complete.")
    mail.logout()
    
    # Convert to list of tuples with country detection and sort alphabetically
    email_country_pairs = [
        (email_addr, detect_country(email_addr))
        for email_addr in sorted(found_addresses)
    ]
    
    return email_country_pairs


def _iter_folders(folder):
    yield folder
    for subfolder in folder.Folders:
        for nested in _iter_folders(subfolder):
            yield nested


def extract_emails_from_pst(pst_path, allowed_folder_names=None):
    """
    Scans a local PST file via Outlook and collects unique email addresses
    with their detected countries.

    Args:
        pst_path (str): Full path to the PST file.
        allowed_folder_names (list | None): Folder names to scan. If None or empty,
            scans all folders.

    Returns:
        list: A list of tuples (email, country) found in the PST file.
    """
    if not win32com:
        print("pywin32 is required for PST scanning. Install it with: pip install pywin32")
        return []

    if not pst_path:
        print("PST path is empty. Set PST_PATH in config.py")
        return []

    pst_path = os.path.abspath(pst_path)
    if not os.path.exists(pst_path):
        print(f"PST file not found: {pst_path}")
        return []

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    namespace.AddStore(pst_path)

    store_root = None
    try:
        for store in namespace.Stores:
            store_path = getattr(store, "FilePath", "")
            if store_path and os.path.abspath(store_path).lower() == pst_path.lower():
                store_root = store.GetRootFolder()
                break

        if not store_root:
            print("Could not locate the PST store after attaching it.")
            return []

        allowed = set(name.lower() for name in (allowed_folder_names or []) if name)
        found_addresses = {}  # Dict to store {email: {"date": date, "count": count}}

        for folder in _iter_folders(store_root):
            if allowed and folder.Name.lower() not in allowed:
                continue

            try:
                items = folder.Items
            except Exception:
                continue

            for item in items:
                if getattr(item, "Class", None) != 43:
                    continue

                # Capture email send date
                try:
                    sent_date = getattr(item, "SentOn", None)
                    if not sent_date:
                        sent_date = getattr(item, "ReceivedTime", None)
                    if sent_date:
                        sent_date = sent_date.strftime("%Y-%m-%d")
                    else:
                        sent_date = "Unknown"
                except Exception:
                    sent_date = "Unknown"

                for header_value in [
                    getattr(item, "SenderEmailAddress", ""),
                    getattr(item, "To", ""),
                    getattr(item, "CC", ""),
                    getattr(item, "BCC", ""),
                ]:
                    if not header_value:
                        continue
                    matches = re.findall(EMAIL_REGEX, header_value)
                    for match in matches:
                        if match not in found_addresses:
                            found_addresses[match] = {"date": sent_date, "count": 1}
                        else:
                            found_addresses[match]["count"] += 1
                            # Keep the most recent date
                            if sent_date != "Unknown":
                                found_addresses[match]["date"] = sent_date

        # Convert dict to list of tuples with country detection
        email_country_pairs = []
        sorted_emails = sorted(found_addresses.items())
        
        # Show progress bar if tqdm is available
        if tqdm:
            iterator = tqdm(sorted_emails, desc="Detecting countries", unit="email")
        else:
            iterator = sorted_emails
        
        for email_addr, data in iterator:
            country, _ = detect_country_with_whois(email_addr)
            email_country_pairs.append((email_addr, country, data["date"], data["count"]))

        return email_country_pairs
    finally:
        if store_root:
            try:
                namespace.RemoveStore(store_root)
            except Exception:
                pass



def save_to_excel(email_country_list):
    """
    Converts the list of (email, country, date, count) tuples into a sorted Excel spreadsheet.
    
    Args:
        email_country_list (list): List of tuples (email, country, date, count).
    """
    if not email_country_list:
        print("No emails found. Skipping file creation.")
        return

    # Create a DataFrame from the email-country-date-count tuples
    df = pd.DataFrame(
        email_country_list,
        columns=['Email Address', 'Country', 'Last Email Date', 'Count']
    )
    
    # Sort by country first, then by count (descending), then by email address
    df.sort_values(by=['Country', 'Count'], ascending=[True, False], inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Export using openpyxl as the underlying engine
    filename = "extracted_contacts.xlsx"
    df.to_excel(filename, index=False)
    
    # Print summary statistics
    print("\n" + "="*60)
    print("SUMMARY STATISTICS")
    print("="*60)
    total_unique = len(email_country_list)
    total_emails = df['Count'].sum()
    countries = df['Country'].nunique()
    unknown = (df['Country'] == 'Unknown').sum()
    print(f"Total unique emails: {total_unique}")
    print(f"Total emails found:  {total_emails}")
    print(f"Countries:           {countries}")
    print(f"Unknown country:     {unknown}")
    print("="*60)
    print(f"\nSuccess! Saved to {filename}")



if __name__ == "__main__":
    # Extract email addresses and their corresponding countries from mailbox
    if getattr(config, "USE_PST", False):
        email_country_data = extract_emails_from_pst(
            getattr(config, "PST_PATH", ""),
            getattr(config, "PST_FOLDERS", None),
        )
    else:
        email_country_data = extract_emails()
    # Save the results to an Excel file, organized by country
    save_to_excel(email_country_data)
