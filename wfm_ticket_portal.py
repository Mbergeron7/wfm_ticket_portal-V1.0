#!/usr/bin/env python3
"""
wfm_ticket_portal.py - Flask app for WFM request portal.

Features:
- Saves uploaded files to uploads/ and stores the saved path + original filename in Google Sheets
- Automatically expands the Google Sheets header row when new form fields appear
- Keeps old data intact
- Sends confirmation emails (SendGrid) if env vars are present
- Improved logging & error handling
"""

from flask import Flask, render_template, request, redirect, session, url_for, flash
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from zoneinfo import ZoneInfo
import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Email, To, Content, PlainTextContent
from werkzeug.utils import secure_filename
import traceback

# --- Config ---
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "9234b8aa0a7c5f289c4fee35b3153713d22a910f")

BASE_DIR = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Allowed upload extensions
ALLOWED_UPLOAD_EXT = {".png", ".jpg", ".jpeg", ".pdf", ".doc", ".docx", ".xls", ".xlsx"}

# Whitelisted user emails (keep as-is; consider moving to env or config file later)
ALLOWED_USERS = [
    'mbergeron@storagevaultcanada.com',
    'jsauve@storagevaultcanada.com',
    'ddevenny@storagevaultcanada.com',
    'mikebergeron36@gmail.com',
    'szenasni@storagevaultcanada.com'
]

# Google Sheets setup (attempt local path first, then service_account.json in repo)
sheet = None
SHEET_KEY = os.environ.get("GOOGLE_SHEET_KEY", "1gzJ30wmAAcwEJ8H_nte7ZgH6suYZjGX_w86BhPIRndU")

def init_google_sheet():
    global sheet
    try:
        print("📂 Current working directory:", os.getcwd())
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        local_path = r"C:\Users\mikeb\OneDrive - StorageVault Canada Inc\3.  Workforce Management\Mike Files\Power BI Files\Power Automate Schedule Files\Ticketing Tool Flask\wfm_ticket_portal\service_account.json"
        fallback_path = os.path.join(BASE_DIR, "service_account.json")
        key_path = local_path if os.path.exists(local_path) else fallback_path

        if not os.path.exists(key_path):
            raise FileNotFoundError(f"Service account file not found at either path: {local_path} or {fallback_path}")

        print(f"🔑 Using service account file: {key_path}")
        creds = ServiceAccountCredentials.from_json_keyfile_name(key_path, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SHEET_KEY).sheet1
        print("✅ Google Sheets connection established.")
    except Exception as e:
        print(f"❌ Google Sheets setup failed: {e}")
        traceback.print_exc()
        sheet = None

init_google_sheet()

# --- Helpers ---
def send_ticket_email(to_email, subject, html_body, plain_body):
    try:
        api_key = os.environ.get("SENDGRID_API_KEY")
        if not api_key:
            print("⚠️ SENDGRID_API_KEY not set; skipping email send.")
            return
        sg = SendGridAPIClient(api_key=api_key)
        from_addr = os.environ.get("EMAIL_USER")
        if not from_addr:
            print("⚠️ EMAIL_USER not set; skipping email send.")
            return

        mail = Mail(from_email=Email(from_addr), to_emails=To(to_email), subject=subject)
        mail.add_content(Content("text/html", html_body))
        mail.add_content(PlainTextContent(plain_body))

        response = sg.send(mail)
        print(f"✅ Email sent: {response.status_code}")
    except Exception as e:
        print(f"❌ SendGrid error: {e}")
        traceback.print_exc()

def allowed_file(filename):
    if not filename:
        return False
    _, ext = os.path.splitext(filename)
    return ext.lower() in ALLOWED_UPLOAD_EXT

def save_upload_and_return_paths(file_storage):
    """
    Save uploaded file to UPLOAD_FOLDER and return tuple (saved_path, original_filename).
    saved_path is the filesystem relative path that will be stored to Google Sheets.
    """
    if not file_storage:
        return ("", "")
    original_name = file_storage.filename or ""
    if original_name == "":
        return ("", "")
    filename = secure_filename(original_name)
    if not allowed_file(filename):
        print(f"⚠️ Upload blocked (disallowed extension): {original_name}")
        return ("", original_name)

    # prepend timestamp to avoid collisions but retain original name visible
    ts = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y%m%d_%H%M%S")
    saved_filename = f"{ts}_{filename}"
    dest = os.path.join(UPLOAD_FOLDER, saved_filename)
    try:
        file_storage.save(dest)
        print(f"✅ Saved upload to {dest}")
        # Return a path relative to the app base (so it's not an absolute machine path)
        # We'll store the relative path (uploads/...) which you requested.
        rel_path = os.path.join("uploads", saved_filename)
        return (rel_path, original_name)
    except Exception as e:
        print(f"❌ Failed to save upload {original_name}: {e}")
        traceback.print_exc()
        return ("", original_name)

def get_sheet_headers():
    """
    Return list of headers from Google Sheet if available, else empty list.
    """
    if not sheet:
        return []
    try:
        values = sheet.get_all_values()
        if not values:
            return []
        return values[0]
    except Exception as e:
        print(f"❌ Error reading sheet headers: {e}")
        traceback.print_exc()
        return []

def ensure_sheet_headers(include_keys):
    """
    Ensure that every key in include_keys exists in the sheet header row.
    If new headers are found, update the sheet header row (row 1) to include them.
    This preserves existing headers and appends new ones to the end.
    """
    if not sheet:
        print("⚠️ Sheet not configured; cannot ensure headers.")
        return []

    try:
        existing_headers = get_sheet_headers()
        if existing_headers is None:
            existing_headers = []

        # Use current headers list; append any new keys that are missing
        changed = False
        for key in include_keys:
            if key not in existing_headers:
                existing_headers.append(key)
                changed = True

        # Always ensure certain administrative columns exist
        for admin in ["Submitted By", "Submitted At", "Closed At"]:
            if admin not in existing_headers:
                existing_headers.append(admin)
                changed = True

        if changed:
            # Update header row (1:1)
            sheet.update('1:1', [existing_headers])
            print("✅ Google Sheet headers updated.")
        return existing_headers
    except Exception as e:
        print(f"❌ Failed to ensure sheet headers: {e}")
        traceback.print_exc()
        return get_sheet_headers()

def append_row_to_sheet(row_data):
    """
    row_data: dict mapping header -> value
    Appends a row to the sheet in correct header order, expanding headers if needed.
    """
    if not sheet:
        print("⚠️ Skipping Google Sheets logging due to missing credentials.")
        return

    # Ensure headers include keys from row_data
    headers = ensure_sheet_headers(list(row_data.keys()))
    # Build row in header order
    row = [row_data.get(h, "") for h in headers]
    try:
        sheet.append_row(row)
        print("✅ Row appended to Google Sheet.")
    except Exception as e:
        # fallback: try using insert_row (older API) or print failure
        print(f"❌ Failed to append to Google Sheets: {e}")
        traceback.print_exc()

# --- Routes ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    try:
        if request.method == 'POST':
            email = request.form.get('email', '').strip().lower()
            if email in ALLOWED_USERS:
                session['authenticated'] = True
                session['user_email'] = email
                return redirect(url_for('form'))
            else:
                return "Access denied. You are not authorized to use this portal.", 403
        return render_template('login.html')
    except Exception as e:
        traceback.print_exc()
        return f"Login error: {e}", 500

@app.route('/', methods=['GET'])
def form():
    if not session.get('authenticated'):
        return redirect(url_for('login'))
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        if not session.get('authenticated'):
            return redirect(url_for('login'))

        # Basic server-side required checks (don't rely solely on client JS)
        advisor_name = request.form.get('advisor_name', '').strip()
        wfm_request = request.form.get('wfm_request', '').strip()
        if not advisor_name or not wfm_request:
            flash("Advisor name and Request Type are required.")
            return redirect(url_for('form'))

        # Convert form fields into a dict
        # We use to_dict(flat=True) to keep single values; files handled separately
        form_data = request.form.to_dict(flat=True)

        # Save uploaded files; for each file field store two values in form_data:
        # <field>_saved_path  -> path we saved (uploads/...)
        # <field>_orig_name   -> original uploaded filename
        for file_field in request.files:
            file_obj = request.files.get(file_field)
            if file_obj and file_obj.filename:
                saved_path, orig_name = save_upload_and_return_paths(file_obj)
                # store both so you have original name and saved path
                form_data[f"{file_field}_saved_path"] = saved_path
                form_data[f"{file_field}_orig_name"] = orig_name

        # Add timestamp and metadata
        timestamp = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S")
        form_data["Submitted By"] = session.get('user_email', '')
        form_data["Submitted At"] = timestamp
        form_data["Closed At"] = ""  # empty until closed

        # Append to Google Sheet (auto-creates new columns for new keys)
        append_row_to_sheet(form_data)

        # Build email body to send back to submitter (if env provided)
        advisor_display = form_data.get('advisor_name') or form_data.get('Advisor Name') or 'Advisor'
        html_body = f"<html><body><p>Hi {advisor_display},</p><p>Your WFM ticket has been received:</p><ul>"
        plain_body = f"Hi {advisor_display},\n\nYour WFM ticket has been received.\n\n"

        # iterate sorted form data to keep email deterministic
        for key, value in sorted(form_data.items()):
            # avoid listing internal blanks for nicer email; include uploaded paths too
            html_body += f"<li><strong>{key}:</strong> {value}</li>"
            plain_body += f"{key}: {value}\n"

        html_body += "</ul><p>We'll notify you once it's resolved.</p><p>Thanks,<br>Workforce Management</p></body></html>"
        plain_body += "\nWe'll notify you once it's resolved.\n\nThanks,\nWorkforce Management"

        # Send confirmation email (safely; function checks env vars)
        send_ticket_email(form_data.get("Submitted By", session.get('user_email', '')), "WFM Ticket Received", html_body, plain_body)

        # Render confirmation page
        return render_template('confirmation.html')

    except Exception as e:
        print(f"❌ Submit error: {e}")
        traceback.print_exc()
        return f"Form submission error: {e}", 500

@app.route('/close_ticket', methods=['POST'])
def close_ticket():
    try:
        if not session.get('authenticated'):
            return redirect(url_for('login'))

        ticket_id = request.form.get('ticket_id')
        if not ticket_id:
            return "Missing ticket ID", 400

        if not sheet:
            return "Google Sheets not configured", 500

        all_rows = sheet.get_all_values()
        if not all_rows or len(all_rows) < 2:
            return "No tickets found", 404

        headers = all_rows[0]
        ticket_index = None

        # Find first row containing ticket_id in any column (case-sensitive)
        for i, row in enumerate(all_rows[1:], start=2):
            if ticket_id in row:
                ticket_index = i
                break

        if ticket_index:
            closed_at = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S")
            try:
                col_index = headers.index("Closed At") + 1
            except ValueError:
                # ensure "Closed At" exists
                ensure_sheet_headers(["Closed At"])
                headers = get_sheet_headers()
                col_index = headers.index("Closed At") + 1

            sheet.update_cell(ticket_index, col_index, closed_at)

            # find submitted_by and advisor name columns if present
            submitted_by = ""
            advisor_name = ""
            if "Submitted By" in headers:
                sb_col = headers.index("Submitted By")
                submitted_by = all_rows[ticket_index - 1][sb_col]
            if "advisor_name" in headers:
                ad_col = headers.index("advisor_name")
                advisor_name = all_rows[ticket_index - 1][ad_col]
            elif "Advisor Name" in headers:
                ad_col = headers.index("Advisor Name")
                advisor_name = all_rows[ticket_index - 1][ad_col]

            html_body = f"""
            <html>
            <body>
                <p>Hi {advisor_name},</p>
                <p>Your WFM ticket <strong>{ticket_id}</strong> has been marked as complete.</p>
                <p>If you have any questions, feel free to reach out.</p>
                <p>Thanks,<br>Workforce Management</p>
            </body>
            </html>
            """

            plain_body = f"""Hi {advisor_name},

Your WFM ticket {ticket_id} has been marked as complete.

If you have any questions, feel free to reach out.

Thanks,
Workforce Management
"""
            if submitted_by:
                send_ticket_email(submitted_by, "WFM Ticket Closed", html_body, plain_body)
            return "Ticket closed and email sent."
        else:
            return "Ticket not found.", 404
    except Exception as e:
        traceback.print_exc()
        return f"Closure error: {e}", 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# Run
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
