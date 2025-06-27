import smtplib
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import time
import json

# === SMTP/IMAP Credentials ===
SMTP_SERVER = "mail.b2bgrowthexpo.com"
SMTP_PORT = 587
SMTP_EMAIL = "speakersengagement@b2bgrowthexpo.com"
SMTP_PASSWORD = "jH!Ra[9q[f68"

IMAP_SERVER = "mail.b2bgrowthexpo.com"
IMAP_PORT = 143
IMAP_EMAIL = SMTP_EMAIL
IMAP_PASSWORD = SMTP_PASSWORD

SENDER_NAME = "Nagendra Mishra"

# === Email Template ===
EMAIL_TEMPLATE = """
<html>
  <body style="font-family: Arial, sans-serif; font-size: 15px; color: #333; background-color: #ffffff; padding: 20px;">
    <p>Dear {%name%},</p>

    <p>
      I hope this message finds you well.<br><br>
      Thank you for showing interest in speaking at our upcoming <strong>{%expo%}</strong>.
      This exciting event will bring together industry leaders, innovators, and professionals 
      for a day of connection, collaboration, and the exchange of valuable insights.
      We would be honoured to welcome you as one of our speakers.
    </p>

    <p>
      While this is an unpaid opportunity, speaking at the Expo offers several key benefits:
    </p>
    <ul>
      <li>Increased visibility and recognition within your industry</li>
      <li>Opportunities to expand your professional network</li>
      <li>A platform to showcase your expertise to a diverse and engaged audience</li>
    </ul>

    <p>
      Our previous events have drawn a dynamic mix of participants, including startup founders, 
      SME owners, corporate executives, and other influential figures from across various sectors—
      ensuring a high-quality audience for your session.
    </p>

    <p>
      If you are interested, please let us know your availability at your earliest convenience 
      so we can reserve your speaking slot and discuss any specific needs you may have.
    </p>

    <p>
      Thank you for considering this invitation. I look forward to the possibility of working with you 
      and hope to welcome you as a valued speaker at the Bournemouth Business Expo.
    </p>

    <p>
      If you would like to schedule a meeting with me,<br>
      please use the link below:<br>
      <a href="https://tidycal.com/nagendra/b2b-discovery-call" target="_blank">https://tidycal.com/nagendra/b2b-discovery-call</a>
    </p>

    <p style="margin-top: 30px;">
      Thanks & Regards,<br>
      <strong>Nagendra Mishra</strong><br>
      Director | B2B Growth Hub<br>
      Mo: +44 7913 027482<br>
      Email: <a href="mailto:nagendra@b2bgrowthhub.com">nagendra@b2bgrowthhub.com</a><br>
      <a href="https://www.b2bgrowthhub.com" target="_blank">www.b2bgrowthhub.com</a>
    </p>

    <p style="font-size: 13px; color: #888;">
      If you don’t want to hear from me again, please let me know.
    </p>
  </body>
</html>
"""

# === Google Sheets Auth (local testing) ===
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file("/etc/secrets/google-credentials.json", scopes=SCOPES)
sheets_api = build("sheets", "v4", credentials=creds)
gc = gspread.authorize(creds)
sheet = gc.open("Expo-Sales-Management").worksheet("speakers-1")
spreadsheet_id = sheet.spreadsheet.id

# === Utilities ===
def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return {
        "red": int(hex_color[0:2], 16) / 255,
        "green": int(hex_color[2:4], 16) / 255,
        "blue": int(hex_color[4:6], 16) / 255
    }

def get_row_colors(start=2, end=1000):
    try:
        range_ = f"{sheet.title}!A{start}:A{end}"
        result = sheets_api.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=[range_],
            fields="sheets.data.rowData.values.effectiveFormat.backgroundColor"
        ).execute()
        rows = result['sheets'][0]['data'][0]['rowData']
        colors = []
        for row in rows:
            color = row['values'][0].get('effectiveFormat', {}).get('backgroundColor', {})
            rgb = (
                int(color.get('red', 0) * 255),
                int(color.get('green', 0) * 255),
                int(color.get('blue', 0) * 255)
            )
            colors.append(rgb)
        return colors
    except Exception as e:
        print(f"❌ Failed to get row colors: {e}")
        return []

def batch_update_cells(updates):
    try:
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": updates
        }
        sheets_api.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        print("✅ Sheet updated.")
    except Exception as e:
        print(f"❌ Failed to update cells: {e}")

def color_row(row_idx, color_hex):
    rgb = hex_to_rgb(color_hex)
    request = {
        "requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet._properties['sheetId'],
                    "startRowIndex": row_idx - 1,
                    "endRowIndex": row_idx
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": rgb
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        }]
    }
    sheets_api.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request).execute()

def add_comment_to_cell(row_idx, col_idx, comment_text):
    try:
        request = {
            "requests": [
                {
                    "createDeveloperMetadata": {
                        "developerMetadata": {
                            "metadataKey": "comment",
                            "metadataValue": comment_text[:250],  # comment limit
                            "visibility": "DOCUMENT",
                            "location": {
                                "dimensionRange": {
                                    "sheetId": sheet._properties['sheetId'],
                                    "dimension": "ROWS",
                                    "startIndex": row_idx - 1,
                                    "endIndex": row_idx
                                }
                            }
                        }
                    }
                }
            ]
        }
        note_request = {
            "requests": [
                {
                    "updateCells": {
                        "rows": [
                            {
                                "values": [
                                    {
                                        "note": comment_text
                                    }
                                ]
                            }
                        ],
                        "fields": "note",
                        "range": {
                            "sheetId": sheet._properties['sheetId'],
                            "startRowIndex": row_idx - 1,
                            "endRowIndex": row_idx,
                            "startColumnIndex": col_idx,
                            "endColumnIndex": col_idx + 1
                        }
                    }
                }
            ]
        }
        sheets_api.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=note_request).execute()
    except Exception as e:
        print(f"❌ Failed to add comment: {e}")

def send_email(to_email, subject, body_html):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{SENDER_NAME} <{SMTP_EMAIL}>"
    msg["To"] = to_email
    msg.attach(MIMEText(body_html, "html"))

    try:
        # Send email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, to_email, msg.as_string())

        print(f"✅ Email sent to {to_email}", flush=True)

        # Save to Sent folder
        with imaplib.IMAP4(IMAP_SERVER, IMAP_PORT) as imap:
            imap.login(IMAP_EMAIL, IMAP_PASSWORD)
            imap.append("INBOX.Sent", '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())

        print(f"📬 Saved to INBOX.Sent for {to_email}", flush=True)

    except Exception as e:
        print(f"❌ SMTP/IMAP error: {e}", flush=True)

def get_reply_emails():
    replies = {}
    try:
        with imaplib.IMAP4(IMAP_SERVER, IMAP_PORT) as mail:
            mail.login(IMAP_EMAIL, IMAP_PASSWORD)
            mail.select("INBOX")
            typ, data = mail.search(None, 'UNSEEN')
            for num in data[0].split():
                typ, msg_data = mail.fetch(num, '(RFC822)')
                msg = email.message_from_bytes(msg_data[0][1])
                from_email = email.utils.parseaddr(msg['From'])[1].lower()

                body = ""
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        try:
                            body = part.get_payload(decode=True).decode(errors="ignore").strip()
                        except:
                            body = ""
                        break

                replies[from_email] = body
    except Exception as e:
        print(f"❌ IMAP fetch error: {e}")
    return replies

# === Main Functions ===
def process_speakers_emails():
    print("📤 Processing new speaker emails...")
    rows = sheet.get_all_records()
    row_colors = get_row_colors(2, len(rows) + 1)
    updates = []
    today = datetime.today().strftime("%d-%m-%Y")

    for i, row in enumerate(rows, start=2):
        rgb = row_colors[i - 2]
        if rgb != (255, 255, 255):
            continue
        if row.get("Reply Status") or row.get("Email Sent-Date"):
            continue

        name = row.get("First_Name", "").strip()
        email_addr = row.get("Email", "").strip()
        if not email_addr:
            continue

        expo = row.get("Show", "").strip()
        email_html = EMAIL_TEMPLATE.replace("{%name%}", name).replace("{%expo%}", expo)
        send_email(email_addr, "You Showed Interest in Speaking — Here's What’s Next", email_html)

        updates.append({"range": f"{sheet.title}!F{i}", "values": [["Pending"]]})  # Reply Status (Column F)
        updates.append({"range": f"{sheet.title}!E{i}", "values": [[today]]})     # Email Sent-Date (Column E)

    if updates:
        batch_update_cells(updates)

def process_speaker_replies():
    print("📥 Checking speaker replies...")
    replied_emails = get_reply_emails()
    rows = sheet.get_all_records()
    row_colors = get_row_colors(2, len(rows) + 1)
    updates = []

    for i, row in enumerate(rows, start=2):
        rgb = row_colors[i - 2]
        if rgb != (255, 255, 255):
            continue
        email_addr = row.get("Email", "").strip().lower()
        if row.get("Reply Status") == "Replied":
            continue
        if email_addr in replied_emails:
            updates.append({"range": f"{sheet.title}!F{i}", "values": [["Replied"]]})  # Column F
            color_row(i, "#FFFF00")
            comment = replied_emails[email_addr]
            add_comment_to_cell(i, 2, comment)  # Column C = First Name

    if updates:
        batch_update_cells(updates)

# === Run Loop ===
if __name__ == "__main__":
    print("🚀 Speaker automation started.")
    next_send_time = time.time()

    while True:
        try:
            process_speaker_replies()
            if time.time() >= next_send_time:
                process_speakers_emails()
                next_send_time = time.time() + 43200  # 12 hrs
        except Exception as e:
            print(f"❌ Error: {e}")
        time.sleep(30)
