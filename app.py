import os
from flask import Flask, render_template, request, redirect, url_for, abort, session, current_app
from werkzeug.utils import secure_filename

from extractor import Extractor
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024
app.config['UPLOAD_EXTENSIONS'] = ['.pdf', '.PDF']
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = 'SomeVerySecretKey'
app.config['SCOPES'] = ["https://www.googleapis.com/auth/calendar"]


def get_extractor() -> Extractor:
    if not hasattr(current_app, 'extractor'):
        current_app.extractor = Extractor(app.config['SCOPES'])
    return current_app.extractor


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/login')
def login():
    extractor = get_extractor()
    logged_in = extractor.check_login()

    if logged_in:
        res = extractor.get_calendars()
        if res == ValueError or res == HttpError:
            abort(400, res)
        else:
            session['calendars'] = res
            return redirect(url_for('upload'))


@app.route('/upload')
def upload_index():
    calendars = session['calendars']
    return render_template('upload.html', calendars=calendars)


@app.route('/upload', methods=['POST'])
def upload():
    selected_calendar = request.form.get('calendar')
    uploaded_file = request.files['work_plan']
    filename = secure_filename(uploaded_file.filename)
    if filename != '':
        file_ext = os.path.splitext(filename)[1]
        if file_ext not in app.config['UPLOAD_EXTENSIONS']:
            abort(400)
        uploaded_file.save(os.path.join(app.config['UPLOAD_FOLDER'], 'work_plan.pdf'))

    return redirect(url_for('upload'))


if __name__ == '__main__':
    app.run(debug=True)
