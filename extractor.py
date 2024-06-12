import tabula as tb
import pandas as pd
from datetime import datetime

import os.path

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class Extractor:
    def __init__(self, scopes):
        self.scopes = scopes
        self.creds = None

    @staticmethod
    def extract_shifts(name):
        """
            Extracts shifts, the concert month, and the working dates from the Work Plan PDF.
            First extract data from pdf and store them in a csv file.
            After loading it with Pandas to manipulate the data.
            :param name: The name of the Person to extract the shifts for.
            :return: the extracted shifts, the concert month, and the month dates.
            """
        extracted_shifts = []

        with open("uploads/work_plan.pdf", 'rb') as file:
            pdf_data = tb.read_pdf(file, pages=1, stream=True)
            pdf_data[0].to_csv("uploads/output.csv", header=False, index=False)

            df = pd.read_csv("uploads/output.csv", header=None)

            month = df[0][0]
            dates_array = df.iloc[1]

            for index, row in df.iterrows():
                if row[0] == name:
                    extracted_shifts.append(row)
        return extracted_shifts, month, dates_array

    @staticmethod
    def calender_event_start(shift, date):
        """
        Function to crate the start date of the calendar event.
        :param shift: Shift information e.g. "K00"
        :param date: the date for the calendar event
        :return: datetime format for the start of the calendar event
        """
        shift_time_dict = {
            "K00": datetime.strptime(date + "03:15", "%d.%m.%y%H:%M"),
            "K01": datetime.strptime(date + "03:15", "%d.%m.%y%H:%M"),
            "K02": datetime.strptime(date + "04:00", "%d.%m.%y%H:%M"),
            "K03": datetime.strptime(date + "06:00", "%d.%m.%y%H:%M"),
            "K04": datetime.strptime(date + "07:00", "%d.%m.%y%H:%M"),
            "SK1": datetime.strptime(date + "02:15", "%d.%m.%y%H:%M"),
            "SK2": datetime.strptime(date + "03:00", "%d.%m.%y%H:%M"),
            "DK1": datetime.strptime(date + "01:30", "%d.%m.%y%H:%M"),
            "DK2": datetime.strptime(date + "02:00", "%d.%m.%y%H:%M"),
            "DK3": datetime.strptime(date + "03:00", "%d.%m.%y%H:%M"),
        }

        return shift_time_dict.get(shift)

    @staticmethod
    def calender_event_end(shift, date):
        """
        Function to crate the end date of the calendar event.
        :param shift: Shift information e.g. "K00"
        :param date: the date for the calendar event
        :return: datetime format for the end of the calendar event
        """
        shift_time_dict = {
            "K00": datetime.strptime(date + "11:00", "%d.%m.%y%H:%M"),
            "K01": datetime.strptime(date + "12:15", "%d.%m.%y%H:%M"),
            "K02": datetime.strptime(date + "13:15", "%d.%m.%y%H:%M"),
            "K03": datetime.strptime(date + "15:15", "%d.%m.%y%H:%M"),
            "K04": datetime.strptime(date + "16:15", "%d.%m.%y%H:%M"),
            "SK1": datetime.strptime(date + "11:15", "%d.%m.%y%H:%M"),
            "SK2": datetime.strptime(date + "12:00", "%d.%m.%y%H:%M"),
            "DK1": datetime.strptime(date + "09:00", "%d.%m.%y%H:%M"),
            "DK2": datetime.strptime(date + "09:00", "%d.%m.%y%H:%M"),
            "DK3": datetime.strptime(date + "09:00", "%d.%m.%y%H:%M"),
        }

        return shift_time_dict.get(shift)

    def generate_shifts_date_object(self, shift_data, dates_data, year_data):
        """
        Function to create the shifts date object.
        :param shift_data: date for the shift
        :param dates_data: all dates for the month
        :param year_data:  and year information
        :return:
        """
        month_dict = {
            "Jan": "1", "Feb": "2", "Mär": "3", "Apr": "4", "Mai": "5", "Jun": "6",
            "Jul": "7", "Aug": "8", "Sep": "9", "Okt": "10", "Nov": "11", "Dez": "12"
        }

        shift_dates = []
        month, year = year_data.split()
        month_number = month_dict[month]

        # Loading, shifts_data as dataframe & transpose it from columns to rows
        shift_df = pd.DataFrame(shift_data, index=None)
        shift_df_transposed = shift_df.transpose()

        # loading dates_data as dataframe & transpose it
        dates_df = pd.DataFrame(dates_data, index=None)

        # get lengths of dataframes
        shifts_max_count = len(shift_df_transposed)

        loop_count = 0

        for _ in range(shifts_max_count):
            if loop_count > 1:

                shift_to_append = shift_df_transposed.iat[loop_count, 0]
                if pd.isna(shift_to_append):
                    shift_to_append = "X"

                # shift_to_append = shift_df_transposed.iat[loop_count, 0]
                date_string = dates_df.iat[loop_count, 0] + month_number + "." + year
                start_date = self.calender_event_start(shift_to_append, date_string)
                end_date = self.calender_event_end(shift_to_append, date_string)

                if start_date is not None and end_date is not None:
                    shift_dates.append({"shift": shift_to_append, "start_date": start_date, "end_date": end_date})
            loop_count += 1

        return shift_dates

    def create_google_events(self, events, calendar_id):
        """
        Creates all the events for every shift in the Work Plan.
        Uses the Google Calendar API to create the events.
        :param calendar_id:
        :param events: list of shift information
        """

        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists("credentials/token.json"):
            creds = Credentials.from_authorized_user_file("credentials/token.json", self.scopes)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials/credentials.json", self.scopes
                )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("credentials/token.json", "w") as token:
                token.write(creds.to_json())
        """

        try:
            service = build("calendar", "v3", credentials=self.creds)

            for entry in events:
                event = {
                    'summary': entry['shift'],
                    'start': {
                        'dateTime': entry['start_date'].isoformat(),
                        'timeZone': 'Europe/Zurich',
                    },
                    'end': {
                        'dateTime': entry['end_date'].isoformat(),
                        'timeZone': 'Europe/Zurich',
                    }
                }
                event = service.events().insert(
                    calendarId=calendar_id,
                    # Insert CalendarID
                    body=event).execute()
                print('Event created: %s' % (event.get('htmlLink')))

        except HttpError as error:
            print(f"An error occurred: {error}")

    def check_login(self):
        if os.path.exists('credentials/token.json'):
            self.creds = Credentials.from_authorized_user_file('credentials/token.json', self.scopes)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                try:
                    self.creds.refresh(Request())
                except RefreshError:
                    os.remove('credentials/token.json')
                    flow = InstalledAppFlow.from_client_secrets_file(
                        "credentials/credentials.json", self.scopes
                    )
                    self.creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials/credentials.json', self.scopes
                )
                self.creds = flow.run_local_server(port=0)

        with open('credentials/token.json', 'w') as token:
            token.write(self.creds.to_json())

        return True

    def get_calendars(self):
        try:
            service = build("calendar", "v3", credentials=self.creds)
            calendars_result = service.calendarList().list().execute()
            all_calendars = calendars_result.get('items', [])

            if not all_calendars:
                e = ValueError('no Values')
                return e
            else:
                valid_calendars = []
                for calendar in all_calendars:
                    if calendar.get('accessRole') == 'owner':
                        valid_calendars.append(calendar)

                return valid_calendars

        except HttpError as e:
            return e

    def process_data(self, name, calendar_id):
        shifts, month_year, dates = self.extract_shifts(name)
        events_data = self.generate_shifts_date_object(shifts, dates, month_year)
        self.create_google_events(events_data, calendar_id)
