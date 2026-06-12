import msal
import requests
import pdfkit
from email import message_from_bytes
from email.policy import default

# --- Configuration ---
# You need to register a quick, free app in the Azure Portal to get a Client ID (Steps below)
CLIENT_ID = "YOUR_AZURE_CLIENT_ID"
TENANT_ID = "common"  # 'common' works perfectly for all personal Hotmail/Outlook accounts
SCOPES = ["Mail.Read"]

def get_outlook_email():
    # 1. Authenticate with Microsoft OAuth2 (Device Code Flow)
    app = msal.PublicClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}")
    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        print("Failed to create device flow authentication.")
        return

    print("*" * 60)
    print(flow["message"]) # Displays the URL and the code to enter
    print("*" * 60)

    # Blocks and waits until you successfully authenticate in your browser
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print(f"Authentication failed: {result.get('error_description')}")
        return

    token = result["access_token"]
    headers = {"Authorization": f"Bearer {token}"}

    # 2. Fetch the ID of the most recent email
    print("Fetching latest email info...")
    messages_url = "https://graph.microsoft.com/v1.0/me/messages?$top=1&$select=id"
    msg_list_response = requests.get(messages_url, headers=headers).json()

    if not msg_list_response.get("value"):
        print("No emails found in your inbox.")
        return

    latest_message_id = msg_list_response["value"][0]["id"]

    # 3. Download the raw email content as .eml
    # Adding '/$value' tells Microsoft Graph to return the raw MIME data (EML format)
    print("Downloading raw email (.eml)...")
    eml_url = f"https://graph.microsoft.com/v1.0/me/messages/{latest_message_id}/$value"
    eml_response = requests.get(eml_url, headers=headers)

    if eml_response.status_code != 200:
        print("Failed to download the EML format.")
        return

    raw_eml = eml_response.content

    # Save EML File
    eml_filename = "outlook_email.eml"
    with open(eml_filename, "wb") as f:
        f.write(raw_eml)
    print(f" Saved: {eml_filename}")

    # 4. Extract HTML/Text content for PDF conversion
    msg = message_from_bytes(raw_eml, policy=default)
    body_part = msg.get_body(preferencelist=('html', 'plain'))

    if body_part:
        body_content = body_part.get_content()
        if body_part.get_content_type() == 'text/plain':
            body_content = f"<html><body><pre style='white-space: pre-wrap;'>{body_content}</pre></body></html>"

        # 5. Convert to PDF
        pdf_filename = "outlook_email.pdf"
        print("Generating PDF...")
        try:
            pdfkit.from_string(body_content, pdf_filename)
            print(f" Saved: {pdf_filename}")
        except Exception as e:
            print(f"PDF generation failed: {e}")
    else:
        print("Could not extract a readable body.")

if __name__ == "__main__":
    get_outlook_email()