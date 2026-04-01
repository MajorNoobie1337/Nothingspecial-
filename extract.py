import re
import zipfile

OLM_FILE = "ibm_test.olm"

ACCOUNT_PATTERN = r'account:\s*([A-Za-z0-9\-]+)\s*\(Account ID:\s*([a-f0-9]+)\)'
IBM_INVITE_PATTERN = r'https://cloud\.ibm\.com/registration/accept-invite-start\?token=[A-Za-z0-9\-_]+'

results = []

with zipfile.ZipFile(OLM_FILE, 'r') as z:
    xml_files = [n for n in z.namelist() if n.endswith('.xml')]
    for xml_file in xml_files:
        with z.open(xml_file) as f:
            content = f.read().decode('utf-8', errors='replace')
            links = re.findall(IBM_INVITE_PATTERN, content)
            accounts = re.findall(ACCOUNT_PATTERN, content)
            for i, link in enumerate(links):
                account_name = accounts[i][0] if i < len(accounts) else "Unknown"
                account_id = accounts[i][1] if i < len(accounts) else "Unknown"
                results.append((account_name, account_id, link))

# Save to file
with open("links_with_accounts.txt", "w") as f:
    for account_name, account_id, link in results:
        f.write(f"{account_name} | {account_id} | {link}\n")

print(f"Saved {len(results)} entries to links_with_accounts.txt")
```

Run this as a separate script — it will generate `links_with_accounts.txt` with each line like:
```
dMZR-PRD-ac002i000243 | 89345d283bf... | https://cloud.ibm.com/...
