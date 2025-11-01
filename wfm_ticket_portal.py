from flask import Flask, render_template, request, redirect, session
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os

app = Flask(__name__)
app.secret_key = '9234b8aa0a7c5f289c4fee35b3153713d22a910f'  # Replace with a secure random key

# Whitelisted email addresses
ALLOWED_USERS = [
    'mbergeron@storagevaultcanada.com',
    'jsauve@storagevaultcanada.com',
    'ddevenny@storagevaultcanada.com',
    'teresa@example.com'
]

# Google Sheets setup
print("📂 Current working directory:", os.getcwd())
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
sheet = None

try:
    # Absolute path for local dev
    local_path = r"C:\Users\mikeb\OneDrive - StorageVault Canada Inc\3.  Workforce Management\Mike Files\Power BI Files\Power Automate Schedule Files\Ticketing Tool Flask\wfm_ticket_portal\service_account.json"
    
    # Fallback to relative path for deployment
    fallback_path = os.path.join(os.path.dirname(__file__), "service_account.json")
    
    key_path = local_path if os.path.exists(local_path) else fallback_path
    print(f"🔑 Using service account file: {key_path}")

    creds = ServiceAccountCredentials.from_json_keyfile_name(key_path, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key("1gzJ30wmAAcwEJ8H_nte7ZgH6suYZjGX_w86BhPIRndU").sheet1
    print("✅ Google Sheets connection established.")
except Exception as e:
    print(f"❌ Google Sheets setup failed: {e}")

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
            closed_timestamp = ""  # Placeholder for future ticket closure

            # Build row and headers dynamically
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

                # Insert headers if sheet is empty
                if not existing_headers:
                    sheet.insert_row(list(row_data.keys()), 1)
                    existing_headers = list(row_data.keys())

                # Add any new headers dynamically
                for key in row_data.keys():
                    if key not in existing_headers:
                        existing_headers.append(key)
                        sheet.update('1:1', [existing_headers])

                # Align row with header order
                row = [row_data.get(header, '') for header in existing_headers]
                sheet.append_row(row)
                print("✅ Row appended to Google Sheet.")
            else:
                print("⚠️ Skipping Google Sheets logging due to missing credentials.")

            return render_template('confirmation.html')
        return render_template('form.html')
    except Exception as e:
        return f"Form error: {e}", 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')
