from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        advisor_name = request.form.get('advisor_name')
        team_lead = request.form.get('team_lead')
        request_type = request.form.get('request_type')
        details = request.form.get('details')
        # You can log or store this data here
        return render_template('confirmation.html', advisor_name=advisor_name)
    return render_template('form.html')
