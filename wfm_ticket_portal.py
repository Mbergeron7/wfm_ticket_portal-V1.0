from flask import Flask, render_template, request, redirect, session, url_for, flash
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from zoneinfo import ZoneInfo
import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Email, To, Content, PlainTextContent
from werkzeug.utils import secure_filename

# --- Config ---
app = Flask(__name__)
app.secret_key = '9234b8aa0a7c5f289c4fee35b3153713d22a910f'  # keep or move to env var

# Uploads: ensure this folder exists and is writable
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
ALLOWED_UPLOAD_EXT = {".png", ".jpg", ".jpeg", ".pdf", ".doc", ".docx", ".xls", ".xlsx"}

# Whitelisted email addresses
ALLOWED_USERS = [
    'mbergeron@storagevaultcanada.com',
    'jsauve@storagevaultcanada.com',
    'ddevenny@storagevaultcanada.com',
    'mikebergeron36@gmail.com',
    'szenasni@storagevaultcanada.com'
]

# Google Sheets setup
print("📂 Current working directory:", os.getcwd())
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
sheet = None

try:
    local_path = r"C:\Users\mikeb\OneDrive - StorageVault Canada Inc\3.  Workforce Management\Mike Files\Power BI Files\Power Automate Schedule Files\Ticketing Tool Flask\wfm_ticket_portal\service_account.json"
    fallback_path = os.path.join(os.path.dirname(__file__), "service_account.json")
    key_path = local_path if os.path.exists(local_path) else fallback_path
    print(f"🔑 Using service account file: {key_path}")

    creds = ServiceAccountCredentials.from_json_keyfile_name(key_path, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key("1gzJ30wmAAcwEJ8H_nte7ZgH6suYZjGX_w86BhPIRndU").sheet1
    print("✅ Google Sheets connection established.")
except Exception as e:
    print(f"❌ Google Sheets setup failed: {e}")

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

def allowed_file(filename):
    if not filename:
        return False
    _, ext = os.path.splitext(filename)
    return ext.lower() in ALLOWED_UPLOAD_EXT

def safe_save_file(file_storage):
    """
    Save uploaded file and return saved path or empty string on fail.
    """
    if not file_storage:
        return ""
    filename = secure_filename(file_storage.filename)
    if filename == "":
        return ""
    if not allowed_file(filename):
        print(f"⚠️ Upload blocked (disallowed extension): {filename}")
        return ""
    dest = os.path.join(UPLOAD_FOLDER, filename)
    try:
        file_storage.save(dest)
        print(f"✅ Saved upload to {dest}")
        return dest
    except Exception as e:
        print(f"❌ Failed to save upload {filename}: {e}")
        return ""

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
        return f"Login error: {e}", 500

@app.route('/', methods=['GET'])
def form():
    # GET only: show the form (user must be logged in)
    if not session.get('authenticated'):
        return redirect(url_for('login'))
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        if not session.get('authenticated'):
            return redirect(url_for('login'))

        # Validate JS-level validation is not bypassed: basic checks
        # (You can expand validation here as needed)
        wfm_request = request.form.get('wfm_request', '').strip()
        advisor_name = request.form.get('advisor_name', '').strip()
        if not advisor_name or not wfm_request:
            flash("Advisor name and Request Type are required.")
            return redirect(url_for('form'))

        # Build data dict from form
        form_data = request.form.to_dict(flat=True)  # single-valued fields
        # Save uploaded files (if any) and add paths to form_data
        for file_field in request.files:
            file_obj = request.files.get(file_field)
            if file_obj and file_obj.filename:
                saved = safe_save_file(file_obj)
                # store the saved path or original filename into the form data for logging
                form_data[file_field] = saved or file_obj.filename

        timestamp = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S")
        closed_timestamp = ""

        # Capture metadata
        row_data = form_data.copy()
        row_data["Submitted By"] = session.get('user_email', '')
        row_data["Submitted At"] = timestamp
        row_data["Closed At"] = closed_timestamp

        # Write to Google Sheet (if configured)
        if sheet:
            try:
                existing_values = sheet.get_all_values()
                existing_headers = existing_values[0] if existing_values else []

                # Add headers if missing
                for key in row_data.keys():
                    if key not in existing_headers:
                        existing_headers.append(key)
                # Update header row once
                sheet.update('1:1', [existing_headers])

                # Append row in correct header order
                row = [row_data.get(header, '') for header in existing_headers]
                sheet.append_row(row)
                print("✅ Row appended to Google Sheet.")
            except Exception as e:
                print(f"❌ Failed to append to Google Sheets: {e}")
        else:
            print("⚠️ Skipping Google Sheets logging due to missing credentials.")

        # Build email body
        advisor_display = row_data.get('advisor_name') or row_data.get('Advisor Name') or 'Advisor'
        html_body = f"<html><body><p>Hi {advisor_display},</p><p>Your WFM ticket has been received:</p><ul>"
        plain_body = f"Hi {advisor_display},\n\nYour WFM ticket has been received.\n\n"

        for key, value in row_data.items():
            if key not in ["Submitted By", "Closed At"]:
                html_body += f"<li><strong>{key}:</strong> {value}</li>"
                plain_body += f"{key}: {value}\n"

        html_body += "</ul><p>We'll notify you once it's resolved.</p><p>Thanks,<br>Workforce Management</p></body></html>"
        plain_body += "\nWe'll notify you once it's resolved.\n\nThanks,\nWorkforce Management"

        # Send confirmation email to the submitter (if env set)
        send_ticket_email(row_data.get("Submitted By", session.get('user_email', '')), "WFM Ticket Received", html_body, plain_body)

        # Successful submit: render confirmation page
        return render_template('confirmation.html')
    except Exception as e:
        print(f"❌ Submit error: {e}")
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
        headers = all_rows[0]
        ticket_index = None

        for i, row in enumerate(all_rows[1:], start=2):
            if ticket_id in row:
                ticket_index = i
                break

        if ticket_index:
            closed_at = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S")
            col_index = headers.index("Closed At") + 1
            sheet.update_cell(ticket_index, col_index, closed_at)

            email_col = headers.index("Submitted By")
            advisor_col = headers.index("Advisor Name")
            submitted_by = all_rows[ticket_index - 1][email_col]
            advisor_name = all_rows[ticket_index - 1][advisor_col]

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

            send_ticket_email(submitted_by, "WFM Ticket Closed", html_body, plain_body)
            return "Ticket closed and email sent."
        else:
            return "Ticket not found.", 404
    except Exception as e:
        return f"Closure error: {e}", 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# If you run directly
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
