geron@storagevaultcanada.comfrom flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def home():
    try:
        if request.method == 'POST':
            print(request.form)  # This will log submitted form data
            return "Form submitted"
        return render_template('form.html')  # This loads your form page
    except Exception as e:
        return f"Error: {e}", 500  # This shows the error in the browser






