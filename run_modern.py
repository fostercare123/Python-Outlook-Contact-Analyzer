import re
import pandas as pd
from O365 import Account
import modern_config as config # Note the rename here

COUNTRY_TLD_MAP = {'dk': 'Denmark', 'de': 'Germany', 'no': 'Norway', 'se': 'Sweden'} # Add more as needed
GENERIC_TLDS = {'com', 'org', 'net'}

def detect_country(email_addr):
    try:
        tld = email_addr.split('.')[-1].lower()
        return COUNTRY_TLD_MAP.get(tld, 'Unknown') if tld not in GENERIC_TLDS else 'Unknown'
    except: return 'Unknown'

def run_extractor():
    # 'public' flow is easiest for desktop scripts
    credentials = (config.CLIENT_ID, None)
    account = Account(credentials, tenant_id=config.TENANT_ID, auth_flow_type='public')
    
    if not account.is_authenticated:
        # This prints a link to your terminal. Visit it, log in, paste back the URL.
        account.authenticate(scopes=['https://graph.microsoft.com/Mail.Read'])

    mailbox = account.mailbox(resource=config.TARGET_EMAIL)
    inbox = mailbox.inbox_folder()
    
    print(f"Connecting to {config.TARGET_EMAIL}...")
    found_addresses = set()
    
    for msg in inbox.get_messages(limit=200): # Start with 200 to test
        if msg.sender.address:
            found_addresses.add(msg.sender.address)

    results = [(addr, detect_country(addr)) for addr in sorted(found_addresses)]
    df = pd.DataFrame(results, columns=['Email Address', 'Country'])
    df.to_excel("extracted_contacts_MODERN.xlsx", index=False)
    print(f"Success! Found {len(results)} addresses.")

if __name__ == "__main__":
    run_extractor()