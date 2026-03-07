"""
Downloads the loan spreadsheet from SharePoint using Microsoft credentials.
Credentials are stored as GitHub Secrets — never in this file.
"""
import os, sys, re, subprocess

subprocess.check_call([sys.executable, "-m", "pip", "install", "Office365-REST-Python-Client", "-q"])

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

user = os.environ.get("SHAREPOINT_USER")
pwd  = os.environ.get("SHAREPOINT_PASS")
url  = os.environ.get("SHAREPOINT_FILE_URL", "")

if not all([user, pwd, url]):
    print("ERROR: One or more secrets are missing (SHAREPOINT_USER, SHAREPOINT_PASS, SHAREPOINT_FILE_URL)")
    sys.exit(1)

# Extract base hostname e.g. https://apexfunding.sharepoint.com
host_match = re.match(r"(https://[^/]+)", url)
if not host_match:
    print(f"ERROR: URL doesn't look like a SharePoint URL")
    sys.exit(1)

host = host_match.group(1)
site_url = f"{host}/sites/ApexFunding"

print(f"Connecting to: {site_url}")
print(f"As user: {user}")

try:
    credentials = UserCredential(user, pwd)
    ctx = ClientContext(site_url).with_credentials(credentials)

    # Try common file paths
    paths_to_try = [
        "/sites/ApexFunding/Shared Documents/Loan Pipeline Checklist.xlsx",
        "/sites/ApexFunding/Shared Documents/General/Loan Pipeline Checklist.xlsx",
        "/sites/ApexFunding/Documents/Loan Pipeline Checklist.xlsx",
        "/sites/ApexFunding/Shared Documents/Loan%20Pipeline%20Checklist.xlsx",
    ]

    downloaded = False
    for path in paths_to_try:
        try:
            print(f"Trying: {path}")
            f_obj = ctx.web.get_file_by_server_relative_url(path)
            with open("spreadsheet.xlsx", "wb") as f:
                f_obj.download(f)
            ctx.execute_query()
            size = os.path.getsize("spreadsheet.xlsx")
            if size > 5000:
                print(f"Success! Downloaded {size:,} bytes → spreadsheet.xlsx")
                downloaded = True
                break
            else:
                print(f"  File too small ({size} bytes), trying next path...")
        except Exception as e:
            print(f"  Failed: {e}")

    if not downloaded:
        print("\nERROR: Could not find spreadsheet. Trying to list available files...")
        try:
            folder = ctx.web.get_folder_by_server_relative_url("/sites/ApexFunding/Shared Documents")
            files = folder.files
            ctx.load(files)
            ctx.execute_query()
            print("Files in Shared Documents:")
            for f in files:
                print(f"  - {f.properties['Name']}")
        except Exception as e:
            print(f"Could not list files: {e}")
        sys.exit(1)

except Exception as e:
    print(f"Connection failed: {e}")
    print("Check that your Microsoft email and password are correct.")
    sys.exit(1)
