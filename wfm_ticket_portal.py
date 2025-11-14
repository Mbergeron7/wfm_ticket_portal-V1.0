#!/usr/bin/env python3
"""
wfm_ticket_portal.py - Flask app for WFM request portal.

Fully integrated with updated form:
- Captures all fields (including newly added hidden/submit button fields)
- Supports file uploads
- Dynamically expands Google Sheets headers
- Sends confirmation emails (SendGrid)
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

ALLOWED_UPLOAD_EXT = {".png", ".jpg", ".jpeg", ".pdf", ".doc", ".docx", ".xls", ".xlsx"}

ALLOWED_USERS = [
    'mbergeron@storagevaultcanada.com',
    'jsauve@storagevaultcanada.com',
    'ddevenny@storagevaultcanada.com',
    'mikebergeron36@gmail.com',
    'szenasni@storagevaultcanada.com'
]

SHEET_KEY = os.environ.get("GOOGLE_SHEET_KEY", "1gzJ30wmAAcwEJ8H_nte7ZgH6suYZjGX_w86BhPIRndU")
sheet = None

def init_google_sheet():
    global sheet
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        local_path = os.path.join(BASE_DIR, "service_account.json")
        if not os.path.exists(local_path):
            raise FileNotFoundError("service_account.json not found.")
        creds = ServiceAccountCredentials.from_json_keyfile_name(local_path, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SHEET_KEY).sheet1
        print("✅ Google Sheets connected.")
    except Exception as e:
        print(f"❌ Google Sheets init failed: {e}")
        traceback.print_exc()
        sheet = None

init_google_sheet()

# --- Helpers ---
def send_ticket_email(to_email, subject, html_body, plain_body):
    try:
        api_key = os.environ.get("SENDGRID_API_KEY")
        if not api_key: return
        from_addr = os.environ.get("EMAIL_USER")
        if not from_addr: return
        mail = Mail(from_email=Email(from_addr), to_emails=To(to_email), subject=subject)
        mail.add_content(Content("text/html", html_body))
        mail.add_content(PlainTextContent(plain_body))
        sg = SendGridAPIClient(api_key)
        sg.send(mail)
        print(f"✅ Email sent to {to_email}")
    except Exception as e:
        print(f"❌ SendGrid error: {e}")
        traceback.print_exc()

def allowed_file(filename):
    if not filename: return False
    _, ext = os.path.splitext(filename)
    return ext.lower() in ALLOWED_UPLOAD_EXT

def save_upload_and_return_paths(file_storage):
    if not file_storage or not file_storage.filename: return ("", "")
    filename = secure_filename(file_storage.filename)
    if not allowed_file(filename): return ("", file_storage.filename)
    ts = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y%m%d_%H%M%S")
    saved_filename = f"{ts}_{filename}"
    dest = os.path.join(UPLOAD_FOLDER, saved_filename)
    try:
        file_storage.save(dest)
        return (os.path.join("uploads", saved_filename), file_storage.filename)
    except Exception as e:
        print(f"❌ Failed saving upload: {e}")
        traceback.print_exc()
        return ("", file_storage.filename)

def get_sheet_headers():
    if not sheet: return []
    try:
        vals = sheet.get_all_values()
        if not vals: return []
        return vals[0]
    except Exception as e:
        print(f"❌ Error reading headers: {e}")
        traceback.print_exc()
        return []

def ensure_sheet_headers(include_keys):
    if not sheet: return []
    try:
        headers = get_sheet_headers() or []
        changed = False
        for key in include_keys:
            if key not in headers:
                headers.append(key)
                changed = True
        for admin in ["Submitted By", "Submitted At", "Closed At"]:
            if admin not in headers:
                headers.append(admin)
                changed = True
        if changed:
            sheet.update('1:1', [headers])
        return headers
    except Exception as e:
        print(f"❌ Failed ensuring headers: {e}")
        traceback.print_exc()
        return get_sheet_headers()

def append_row_to_sheet(row_data):
    if not sheet: return
    headers = ensure_sheet_headers(list(row_data.keys()))
    row = [row_data.get(h, "") for h in headers]
    try:
        sheet.append_row(row)
    except Exception as e:
        print(f"❌ Append row failed: {e}")
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
                return "Access denied.", 403
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

        advisor_name = request.form.get('advisor_name', '').strip()
        wfm_request = request.form.get('wfm_request', '').strip()
        if not advisor_name or not wfm_request:
            flash("Advisor name and Request Type are required.")
            return redirect(url_for('form'))

        form_data = request.form.to_dict(flat=True)

        # Save uploaded files
        for file_field in request.files:
            file_obj = request.files.get(file_field)
            if file_obj and file_obj.filename:
                saved_path, orig_name = save_upload_and_return_paths(file_obj)
                form_data[f"{file_field}_saved_path"] = saved_path
                form_data[f"{file_field}_orig_name"] = orig_name

        # Add metadata
        timestamp = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S")
        form_data["Submitted By"] = session.get('user_email', '')
        form_data["Submitted At"] = timestamp
        form_data["Closed At"] = ""

        # Append to sheet
        append_row_to_sheet(form_data)

        # Email confirmation
        advisor_display = form_data.get('advisor_name') or 'Advisor'
        html_body = f"<html><body><p>Hi {advisor_display},</p><p>Your WFM ticket has been received:</p><ul>"
        plain_body = f"Hi {advisor_display},\n\nYour WFM ticket has been received.\n\n"
        for key, value in sorted(form_data.items()):
            html_body += f"<li><strong>{key}:</strong> {value}</li>"
            plain_body += f"{key}: {value}\n"
        html_body += "</ul><p>We'll notify you once resolved.</p></body></html>"
        plain_body += "\nWe'll notify you once resolved.\n"

        send_ticket_email(form_data.get("Submitted By", session.get('user_email', '')), "WFM Ticket Received", html_body, plain_body)

        return render_template('confirmation.html')
    except Exception as e:
        traceback.print_exc()
        return f"Submit error: {e}", 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
