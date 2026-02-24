"""
Outlook Email Extractor
Extracts unique email addresses from Outlook headers and saves to Excel.
"""

import imaplib
import email
import re
import pandas as pd
import config  # Local file for secrets

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


def extract_emails():
    """
    Connects to Outlook, scans headers, and collects unique email addresses.
    
    Returns:
        list: A list of unique email addresses found.
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
    
    # Use a set to automatically handle deduplication
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
    return list(found_addresses)


def save_to_excel(address_list):
    """
    Converts the list of addresses into a sorted Excel spreadsheet.
    
    Args:
        address_list (list): The list of unique email addresses.
    """
    if not address_list:
        print("No emails found. Skipping file creation.")
        return

    # Initialize DataFrame and sort for better readability
    df = pd.DataFrame(address_list, columns=['Email Address'])
    df.sort_values(by='Email Address', inplace=True)

    # Export using openpyxl as the underlying engine
    filename = "extracted_contacts.xlsx"
    df.to_excel(filename, index=False)
    print(f"Success! {len(address_list)} addresses saved to {filename}")


if __name__ == "__main__":
    # Execute the extraction and save process
    unique_emails = extract_emails()
    save_to_excel(unique_emails)