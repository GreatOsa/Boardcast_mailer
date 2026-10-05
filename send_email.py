import os
import sys
import re
import time
import base64
import mimetypes
import html

import pandas as pd

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage

from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# ============================================================
# SETTINGS
# ============================================================

EXCEL_FILE = "AuroraCreatorRegistration.xlsx"
IMAGE_FILE = "aurora_publication.jpeg"

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"

OUTPUT_FOLDER = "Generated_Posts"
SENT_LOG = "sent_log.txt"

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

# --- Safety / testing options ---------------------------------

# True  = create the Word documents and print what WOULD happen,
#         but do not send any email.
DRY_RUN = False

# Put your own email here to do a test run: only the FIRST row is
# processed and the email goes to this address instead of the
# applicant. Test sends are NOT written to the sent log.
# Leave as "" for the real run.
TEST_EMAIL = ""

# True  = a blank line is placed between paragraphs in the Word
#         document, so the post keeps its spacing when copied into
#         LinkedIn. False = spacing only (no empty lines).
BLANK_LINE_BETWEEN_PARAGRAPHS = True

# Pause between emails (seconds), to stay friendly with Gmail limits.
DELAY_BETWEEN_EMAILS = 1


# ============================================================
# CREATE OUTPUT FOLDER
# ============================================================

os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ============================================================
# GMAIL AUTHENTICATION
# ============================================================

def authenticate():
    creds = None

    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception:
            print("Saved token could not be read. Re-authenticating...")
            creds = None

    if creds and creds.valid:
        return creds

    # Try a silent refresh first
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            print("Saved login has expired. Re-authenticating...")
            if os.path.exists(TOKEN_FILE):
                os.remove(TOKEN_FILE)
            creds = None

    # Full login in the browser
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            CREDENTIALS_FILE, SCOPES
        )
        creds = flow.run_local_server(port=0)

    with open(TOKEN_FILE, "w") as token:
        token.write(creds.to_json())

    return creds


# ============================================================
# CHECK FILES
# ============================================================

if not os.path.exists(EXCEL_FILE):
    print(f"ERROR: Excel file not found: {EXCEL_FILE}")
    sys.exit(1)

if not os.path.exists(IMAGE_FILE):
    print(f"ERROR: Image file not found: {IMAGE_FILE}")
    sys.exit(1)

if not DRY_RUN and not os.path.exists(CREDENTIALS_FILE) and not os.path.exists(TOKEN_FILE):
    print(f"ERROR: {CREDENTIALS_FILE} not found.")
    sys.exit(1)


# ============================================================
# READ THE IMAGE ONCE
# ============================================================

with open(IMAGE_FILE, "rb") as f:
    IMAGE_BYTES = f.read()

guessed_type, _ = mimetypes.guess_type(IMAGE_FILE)
IMAGE_SUBTYPE = (guessed_type or "image/jpeg").split("/")[-1]


# ============================================================
# READ EXCEL
# ============================================================

df = pd.read_excel(EXCEL_FILE)
df.columns = df.columns.astype(str).str.strip()

print("\nColumns found:")
print(df.columns.tolist())

required_columns = ["Full name", "Email Address", "Post"]
missing_columns = [c for c in required_columns if c not in df.columns]

if missing_columns:
    print("\nERROR: Missing columns:")
    print(missing_columns)
    print("\nAvailable columns:")
    print(df.columns.tolist())
    sys.exit(1)


# ============================================================
# SENT LOG (prevents emailing the same person twice)
# ============================================================

def load_sent_log():
    if not os.path.exists(SENT_LOG):
        return set()
    with open(SENT_LOG, "r") as f:
        return {line.strip().lower() for line in f if line.strip()}


def record_sent(email):
    with open(SENT_LOG, "a") as f:
        f.write(email.lower() + "\n")


already_sent = load_sent_log()


# ============================================================
# HELPERS
# ============================================================

def is_blank(value):
    """True for empty cells, NaN, or the text 'nan'."""
    if pd.isna(value):
        return True
    text = str(value).strip()
    return text == "" or text.lower() == "nan"


def clean_filename(name):
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    return name.strip()


def _join_entity_surrogates(match):
    """Turns &#55356;&#57225; (an emoji split into two halves) back into the emoji."""
    def to_int(token):
        token = token.lstrip("#")
        return int(token[1:], 16) if token[:1] in "xX" else int(token)

    hi, lo = to_int(match.group(1)), to_int(match.group(2))
    if 0xD800 <= hi <= 0xDBFF and 0xDC00 <= lo <= 0xDFFF:
        return chr(0x10000 + ((hi - 0xD800) << 10) + (lo - 0xDC00))
    return match.group(0)


def clean_text(text):
    """
    Fixes text that was stored in the Excel file with HTML codes
    (&#39; -> ', &amp; -> &, emoji codes -> emoji) and tidies spaces.
    """
    text = str(text)

    # Emoji stored as pairs of numeric codes
    text = re.sub(
        r'&(#(?:x[0-9a-fA-F]+|\d+));&(#(?:x[0-9a-fA-F]+|\d+));',
        _join_entity_surrogates,
        text
    )

    # Decode HTML entities (repeat in case they were encoded twice)
    for _ in range(3):
        decoded = html.unescape(text)
        if decoded == text:
            break
        text = decoded

    # Join emoji halves that arrived as raw surrogate characters
    try:
        text = text.encode("utf-16", "surrogatepass").decode("utf-16")
    except Exception:
        pass

    # Non-breaking spaces and zero-width characters
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", text)

    return text


def format_post(post):
    """
    Splits a post into paragraphs: every line in the Excel cell
    becomes its own paragraph, exactly as it was written.
    """
    post = clean_text(post).strip()

    post = post.replace("\r\n", "\n")
    post = post.replace("\r", "\n")

    paragraphs = []

    for line in post.split("\n"):
        line = re.sub(r'[ \t]+', ' ', line).strip()
        if line:
            paragraphs.append(line)

    return paragraphs


def create_word_document(row_number, full_name, post_task):
    # Row number keeps filenames unique if two people share a name
    filename = clean_filename(
        f"{full_name}_Aurora_Publication.docx"
    )
    filepath = os.path.join(OUTPUT_FOLDER, filename)

    document = Document()

    section = document.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    # Title
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("AURORA PUBLICATION")
    title_run.bold = True
    title_run.font.size = Pt(18)


    # Heading
    heading = document.add_paragraph()
    heading_run = heading.add_run("Assigned Post")
    heading_run.bold = True
    heading_run.font.size = Pt(14)

    # Post
    paragraphs = format_post(post_task)

    for i, paragraph_text in enumerate(paragraphs):

        paragraph = document.add_paragraph()
        paragraph.paragraph_format.line_spacing = 1.15

        if BLANK_LINE_BETWEEN_PARAGRAPHS:
            paragraph.paragraph_format.space_after = Pt(0)
        else:
            paragraph.paragraph_format.space_after = Pt(10)

        run = paragraph.add_run(paragraph_text)
        run.font.size = Pt(11)

        if BLANK_LINE_BETWEEN_PARAGRAPHS and i < len(paragraphs) - 1:
            spacer = document.add_paragraph()
            spacer.paragraph_format.space_after = Pt(0)

    document.save(filepath)
    return filepath


def build_message(to_address, first_name, docx_file):
    subject = "Aurora Publication Application Update"

    body = f"""Dear {first_name},

Thank you for your application to the Aurora Publication.

We are pleased to inform you that you have been selected to proceed to the next stage of the process.

To publish your assigned post, please follow the instructions provided in the video shared in the WhatsApp group chat. In particular, you are required to:

1. Let the paragraph of the post be well-structured.
2. Attach the accompanying image (included with this email) to your post.
3. Tag all relevant individuals and companies, as demonstrated in the video.

Your assigned post is attached to this email as a Word document. Kindly review it carefully and follow the instructions before publishing.

Should you require any clarification, please contact any of the WhatsApp group administrators.

Yours sincerely,
Osayande
Aurora Robotics
"""

    message = MIMEMultipart()
    message["To"] = to_address
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    # Word document
    with open(docx_file, "rb") as file:
        doc_attachment = MIMEApplication(
            file.read(),
            _subtype="vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    doc_attachment.add_header(
        "Content-Disposition",
        "attachment",
        filename=os.path.basename(docx_file)
    )
    message.attach(doc_attachment)

    # Image (explicit subtype, so it works on Python 3.13+)
    image = MIMEImage(IMAGE_BYTES, _subtype=IMAGE_SUBTYPE)
    image.add_header(
        "Content-Disposition",
        "attachment",
        filename=os.path.basename(IMAGE_FILE)
    )
    message.attach(image)

    return message


def send_message(service, message):
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    service.users().messages().send(
        userId="me",
        body={"raw": raw_message}
    ).execute()


# ============================================================
# SUMMARY AND CONFIRMATION
# ============================================================

total_people = len(df)
already_count = sum(
    1 for e in df["Email Address"]
    if not is_blank(e) and str(e).strip().lower() in already_sent
)

print("\n" + "=" * 60)

if TEST_EMAIL:
    print(f"TEST MODE: first row only, sent to {TEST_EMAIL}")
elif DRY_RUN:
    print(f"DRY RUN: documents will be created for up to {total_people} people, no emails sent.")
else:
    print(f"Rows in sheet: {total_people}")
    print(f"Already sent (from {SENT_LOG}): {already_count}")
    print(f"Emails to send now: up to {total_people - already_count}")

print("=" * 60)

if not DRY_RUN:
    confirmation = input("\nType YES to continue: ").strip()
    if confirmation != "YES":
        print("\nSending cancelled.")
        sys.exit(0)


# ============================================================
# CONNECT TO GMAIL
# ============================================================

service = None

if not DRY_RUN:
    creds = authenticate()
    service = build("gmail", "v1", credentials=creds)
    print("\nGmail connected successfully!")


# ============================================================
# PROCESS EVERY APPLICANT
# ============================================================

successful = 0
failed = 0
skipped = 0

rows = df.head(1) if TEST_EMAIL else df

for index, row in rows.iterrows():

    row_number = index + 1
    email = ""

    try:
        full_name = "" if is_blank(row["Full name"]) else str(row["Full name"]).strip()
        email = "" if is_blank(row["Email Address"]) else str(row["Email Address"]).strip()
        post_task = "" if is_blank(row["Post"]) else str(row["Post"]).strip()

        label = full_name or f"row {row_number}"

        # ---- Validation ------------------------------------
        if not full_name:
            print(f"\nSKIPPED: row {row_number} has no name.")
            failed += 1
            continue

        if not email:
            print(f"\nSKIPPED: {label} has no email address.")
            failed += 1
            continue

        if "@" not in email:
            print(f"\nSKIPPED: {label} has an invalid email: {email}")
            failed += 1
            continue

        if not post_task:
            print(f"\nSKIPPED: {label} has no post assigned.")
            failed += 1
            continue

        if not TEST_EMAIL and email.lower() in already_sent:
            print(f"SKIPPED (already sent): {email}")
            skipped += 1
            continue

        first_name = full_name.split()[0]

        # ---- Create the document ---------------------------
        docx_file = create_word_document(row_number, full_name, post_task)

        # ---- Dry run ---------------------------------------
        if DRY_RUN:
            print(f"[DRY RUN] {first_name} → {email}")
            print(f"    Document → {os.path.basename(docx_file)}")
            successful += 1
            continue

        # ---- Send ------------------------------------------
        to_address = TEST_EMAIL if TEST_EMAIL else email
        message = build_message(to_address, first_name, docx_file)
        send_message(service, message)

        successful += 1

        if not TEST_EMAIL:
            record_sent(email)
            already_sent.add(email.lower())

        print(f"[{successful}] SENT → {first_name} → {to_address}")
        print(f"    Document → {os.path.basename(docx_file)}")

        time.sleep(DELAY_BETWEEN_EMAILS)

    except KeyboardInterrupt:
        print("\n\nStopped by user. Run the script again to continue; "
              "people already emailed will be skipped.")
        break

    except Exception as e:
        failed += 1
        print(f"\nFAILED → {email or f'row {row_number}'}")
        print(f"Error: {e}")


# ============================================================
# FINAL SUMMARY
# ============================================================

print("\n" + "=" * 60)
print("DONE")
print(f"Successful: {successful}")
print(f"Skipped (already sent): {skipped}")
print(f"Failed / skipped for bad data: {failed}")
print(f"\nWord documents saved in: {OUTPUT_FOLDER}")
print("=" * 60)