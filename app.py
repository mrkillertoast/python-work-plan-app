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

    if type(logged_in).isException:
        return render_template('login.html', error=logged_in)

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
    uploaded_file = request.files['work_plan']
    filename = secure_filename(uploaded_file.filename)
    if filename != '':
        file_ext = os.path.splitext(filename)[1]
        if file_ext not in app.config['UPLOAD_EXTENSIONS']:
            abort(400)
        uploaded_file.save(os.path.join(app.config['UPLOAD_FOLDER'], 'work_plan.pdf'))

    extractor = get_extractor()
    extractor.extract_data_and_names()

    return redirect(url_for('selection'))


@app.route('/selection')
def selection():
    calendars = session['calendars']

    extractor = get_extractor()
    names = extractor.names

    return render_template('selection.html', calendars=calendars, names=names)


@app.route('/selection', methods=['POST'])
def handle_selection():
    selected_name = request.form['person_name']
    selected_calendar = request.form['calendar']

    if selected_name and selected_calendar:
        extractor = get_extractor()
        extractor.process_data(selected_name, selected_calendar)

    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
