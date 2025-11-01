from flask import Flask, render_template, request, redirect, session
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'

# Whitelisted email addresses
ALLOWED_USERS = [
    'mbergeron@storagevaultcanada.com',
    'jsauve@storagevaultcanada.com',
    'ddevenny@storagevaultcanada.com',
    'teresa@example.com'
]

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key("1gzJ30wmAAcwEJ8H_nte7ZgH6suYZjGX_w86BhPIRndU").sheet1

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

            row = [
                data.get('advisor_name', ''),
                data.get('request_date', ''),
                data.get('lob', ''),
                data.get('team_lead', ''),
                data.get('wfm_request', ''),
                data.get('details', ''),
                session.get('user_email', ''),
                timestamp,
                closed_timestamp
            ]
            sheet.append_row(row)
            return render_template('confirmation.html')
        return render_template('form.html')
    except Exception as e:
        return f"Form error: {e}", 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')




