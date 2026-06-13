import os
import re
import msal
import requests
import pdfkit
from email import message_from_bytes
from email.policy import default

# --- Configuration ---
CLIENT_ID = "YOUR_AZURE_CLIENT_ID"
TENANT_ID = "common"
SCOPES = ["Mail.Read"]

# 1. Define your predetermined Outlook categories here
TARGET_CATEGORIES = ["Invoices", "Project Alpha", "Receipts"]
DOWNLOAD_BASE_DIR = "./Outlook_Sorted_Downloads"

def clean_filename(name):
    """Removes invalid file/folder characters from strings."""
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()

def download_categorized_emails():
    # Authenticate with Microsoft OAuth2 (Device Code Flow)
    app = msal.PublicClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}")
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print("Failed to initialize authentication.")
        return

    print("*" * 60)
    print(flow["message"])
    print("*" * 60)

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        print("Authentication failed.")
        return

    headers = {"Authorization": f"Bearer {result['access_token']}"}

    # 2. Iterate through your target categories
    for category in TARGET_CATEGORIES:
        print(f"\n=== Processing Category: [{category}] ===")

        # Dynamically create separate target folder for this category
        category_folder_name = clean_filename(category)
        category_dir = os.path.join(DOWNLOAD_BASE_DIR, category_folder_name)
        os.makedirs(category_dir, exist_ok=True)

        # 3. Query Graph API filtering for the specific category
        messages_url = "https://graph.microsoft.com/v1.0/me/messages"
        params = {
            "$filter": f"categories/any(c:c eq '{category}')",
            "$top": 10,  # Limits payload to latest 10 emails per category
            "$select": "id,subject"
        }

        # Using 'params' in requests automatically handles URL-encoding for category spaces
        response = requests.get(messages_url, headers=headers, params=params).json()
        messages = response.get("value", [])

        if not messages:
            print(f"No emails found matching category: '{category}'")
            continue

        # 4. Download and process individual emails
        for msg_info in messages:
            msg_id = msg_info["id"]
            subject = msg_info.get("subject") or "No Subject"

            # Generate clean unique file names using cleaned subject + short ID snippet
            safe_subject = clean_filename(subject)[:50]
            file_base_name = f"{safe_subject}_{msg_id[:8]}"

            eml_path = os.path.join(category_dir, f"{file_base_name}.eml")
            pdf_path = os.path.join(category_dir, f"{file_base_name}.pdf")

            print(f" -> Downloading: {subject}")

            # Fetch raw .eml contents
            eml_url = f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}/$value"
            eml_resp = requests.get(eml_url, headers=headers)
            if eml_resp.status_code != 200:
                print(f"    Failed to download raw data for ID: {msg_id[:8]}")
                continue

            # Save EML
            with open(eml_path, "wb") as f:
                f.write(eml_resp.content)

            # Parse and render HTML content into PDF
            msg = message_from_bytes(eml_resp.content, policy=default)
            body_part = msg.get_body(preferencelist=('html', 'plain'))

            if body_part:
                body_content = body_part.get_content()
                if body_part.get_content_type() == 'text/plain':
                    body_content = f"<html><body><pre style='white-space: pre-wrap;'>{body_content}</pre></body></html>"
                try:
                    pdfkit.from_string(body_content, pdf_path)
                except Exception as e:
                    print(f"    PDF rendering failed: {e}")
            else:
                print("    Skipping PDF: No extractable text/HTML found.")

if __name__ == "__main__":
    download_categorized_emails()