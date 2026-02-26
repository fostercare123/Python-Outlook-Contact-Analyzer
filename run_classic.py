"""
Outlook Email Extractor
Extracts unique email addresses from Outlook headers and saves to Excel.
"""

import imaplib
import email
import re
import pandas as pd
import config_classic  # Local file for secrets

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
        mail.login(config_classic.EMAIL_USER, config_classic.EMAIL_PASS)
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



def save_to_excel(email_country_list):
    """
    Converts the list of email-country pairs into a sorted Excel spreadsheet.
    
    Args:
        email_country_list (list): List of tuples (email, country).
    """
    if not email_country_list:
        print("No emails found. Skipping file creation.")
        return

    # Create a DataFrame from the email-country pairs with proper column names
    df = pd.DataFrame(
        email_country_list,
        columns=['Email Address', 'Country']
    )
    
    # Sort by country first, then by email address for better organization
    df.sort_values(by=['Country', 'Email Address'], inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Export using openpyxl as the underlying engine
    filename = "extracted_contacts.xlsx"
    df.to_excel(filename, index=False)
    print(f"Success! {len(email_country_list)} addresses saved to {filename}")



if __name__ == "__main__":
    # Extract email addresses and their corresponding countries from mailbox
    email_country_data = extract_emails()
    # Save the results to an Excel file, organized by country
    save_to_excel(email_country_data)
