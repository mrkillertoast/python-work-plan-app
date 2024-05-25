import os
from flask import Flask, render_template, request, redirect, url_for, abort, session
from werkzeug.utils import secure_filename

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024
app.config['UPLOAD_EXTENSIONS'] = ['.pdf']
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = 'SomeVerySecretKey'


@app.route('/')
def index():
    print('tst')
    return render_template('index.html')


@app.route('/login')
def login():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES
            )
            creds = flow.run_local_server(port=0)

    with open('token.json', 'w') as token:
        token.write(creds.to_json())

    try:
        service = build("calendar", "v3", credentials=creds)

        calendars_result = service.calendarList().list().execute()
        calendars = calendars_result.get('items', [])

        if not calendars:
            print('No calendars found.')
        else:
            session['calendars'] = calendars
            return redirect(url_for('upload'))

    except HttpError as e:
        print(e)


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

    return redirect(url_for('upload'))


if __name__ == '__main__':
    app.run(debug=True)
