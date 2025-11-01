from flask import Flask, render_template, request, redirect, session
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Email, To, Cc, Content, PlainTextContent

app = Flask(__name__)
app.secret_key = '9234b8aa0a7c5f289c4fee35b3153713d22a910f'

# Whitelisted email addresses
ALLOWED_USERS = [
    'mbergeron@storagevaultcanada.com',
    'jsauve@storagevaultcanada.com',
    'ddevenny@storagevaultcanada.com',
    'mikebergeron36@gmail.com'
]

# Static CC list
CC_EMAILS = [
    "rc.wfm@storagevaultcanada.com",
    "wfm@yourcompany.com"
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

def send_ticket_email(to_email, subject, html_body, plain_body):
    try:
        sg = SendGridAPIClient(api_key=os.environ.get("SENDGRID_API_KEY"))
        from_email = Email(os.environ.get("EMAIL_USER"))
        to = To(to_email)
        cc = [Cc(email) for email in CC_EMAILS]

        mail = Mail(from_email=from_email, to_emails=to, subject=subject)
        mail.cc = cc
        mail.add_content(Content("text/html", html_body))
        mail.add_content(PlainTextContent(plain_body))

        response = sg.send(mail)
        print(f"✅ Email sent: {response.status_code}")
    except Exception as e:
        print(f"❌ SendGrid error: {e}")

@app.route('/login', methods=['GET', 'POST'])
def login():
    try:
        if request.method == 'POST':
            email = request.form['email'].strip().lower()
            if email in ALLOWED_USERS:
                session['authenticated'] = True
                session['user_email'] = email
                return redirect('/')
            else:
                return "Access denied. You are not authorized to use this portal."
        return render_template('login.html')
    except Exception as e:
        return f"Login error: {e}", 500

@app.route('/', methods=['GET', 'POST'])
def home():
    try:
        if not session.get('authenticated'):
            return redirect('/login')
        if request.method == 'POST':
            data = request.form.to_dict()
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            closed_timestamp = ""

            row_data = {
                "Advisor Name": data.get('advisor_name', ''),
                "Date of Request": data.get('request_date', ''),
                "LoB": data.get('lob', ''),
                "Team Lead": data.get('team_lead', ''),
                "WFM Request Type": data.get('wfm_request', ''),
                "Details": data.get('details', ''),
                "Submitted By": session.get('user_email', ''),
                "Submitted At": timestamp,
                "Closed At": closed_timestamp
            }

            if sheet:
                existing_values = sheet.get_all_values()
                existing_headers = existing_values[0] if existing_values else []

                if not existing_headers:
                    sheet.insert_row(list(row_data.keys()), 1)
                    existing_headers = list(row_data.keys())

                for key in row_data.keys():
                    if key not in existing_headers:
                        existing_headers.append(key)
                        sheet.update('1:1', [existing_headers])

                row = [row_data.get(header, '') for header in existing_headers]
                sheet.append_row(row)
                print("✅ Row appended to Google Sheet.")

                html_body = f"""
                <html>
                <body>
                    <p>Hi {row_data['Advisor Name']},</p>
                    <p>Your WFM ticket has been received:</p>
                    <ul>
                        <li><strong>Date:</strong> {row_data['Date of Request']}</li>
                        <li><strong>LoB:</strong> {row_data['LoB']}</li>
                        <li><strong>Team Lead:</strong> {row_data['Team Lead']}</li>
                        <li><strong>Request Type:</strong> {row_data['WFM Request Type']}</li>
                        <li><strong>Details:</strong> {row_data['Details']}</li>
                        <li><strong>Submitted At:</strong> {row_data['Submitted At']}</li>
                    </ul>
                    <p>We'll notify you once it's resolved.</p>
                    <p>Thanks,<br>Workforce Management</p>
                </body>
                </html>
                """

                plain_body = f"""Hi {row_data['Advisor Name']},

Your WFM ticket has been received.

Date: {row_data['Date of Request']}
LoB: {row_data['LoB']}
Team Lead: {row_data['Team Lead']}
Request Type: {row_data['WFM Request Type']}
Details: {row_data['Details']}
Submitted At: {row_data['Submitted At']}

We'll notify you once it's resolved.

Thanks,
Workforce Management
"""

                send_ticket_email(row_data["Submitted By"], "WFM Ticket Received", html_body, plain_body)
            else:
                print("⚠️ Skipping Google Sheets logging due to missing credentials.")

            return render_template('confirmation.html')
        return render_template('form.html')
    except Exception as e:
        return f"Form error: {e}", 500

@app.route('/close_ticket', methods=['POST'])
def close_ticket():
    try:
        if not session.get('authenticated'):
            return redirect('/login')
        
        ticket_id = request.form.get('ticket_id')
        if not ticket_id:
            return "Missing ticket ID", 400

        all_rows = sheet.get_all_values()
        headers = all_rows[0]
        ticket_index = None

        for i, row in enumerate(all_rows[1:], start=2):
            if ticket_id in row:
                ticket_index = i
                break

        if ticket_index:
            closed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
    return redirect('/login')
