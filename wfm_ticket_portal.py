from flask import Flask, render_template, request, redirect, session

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Replace with a secure random key

# Whitelisted email addresses
ALLOWED_USERS = [
    'mbergeron@storagevaultcanada.com',
    'jsauve@storagevaultcanada.com',
    'ddevenny@storagevaultcanada.com',
    'teresa@example.com'
]

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
            print(request.form)  # Logs submitted form data
            return "Form submitted"
        return render_template('form.html')
    except Exception as e:
        return f"Form error: {e}", 500

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')
