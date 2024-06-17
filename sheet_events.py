import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import time
import datetime
import pytz

class GoogleCalendarScheduler:
    def __init__(self, creds_file, spreadsheet_id):
        self.scopes = ["https://www.googleapis.com/auth/spreadsheets",
                       "https://www.googleapis.com/auth/drive"]
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, self.scopes)
        self.client = gspread.authorize(self.creds)
        self.service = build('sheets', 'v4', credentials=self.creds)
        self.spreadsheet_id = spreadsheet_id

    def get_sheets(self):
        start_cell = "Add name of the person to find. Ex: Nathan"
        end_cell = "Add reference to the time of the last event. Ex: 3:00 PM"
        sheets = []
        columns = []
        start_rows = []
        end_rows = []

        spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()

        for sheet in spreadsheet.get('sheets', []):
            sheet_title = sheet.get("properties", {}).get("title", "")
            try:
                sheet_data = self.service.spreadsheets().values().get(
                    spreadsheetId=self.spreadsheet_id,
                    range=sheet_title
                ).execute()

                values = sheet_data.get('values', [])
                start = None

                for i, row in enumerate(values):
                    if start_cell in row:
                        start = (i + 1, row.index(start_cell) + 1)
                        break

                if start:
                    sheets.append(sheet_title)
                    start_rows.append(start[0])
                    columns.append(start[1])

                    for i in range(start[0], start[0] + 1000):
                        if end_cell in values[i - 1]:
                            end_rows.append(i)
                            break
                time.sleep(1)
            except HttpError as e:
                if e.resp.status == 429:
                    print("Rate limit exceeded. Retrying...")
                    time.sleep(60)
                else:
                    raise

        if sheets:
            return [sheets, columns, start_rows, end_rows]
        else:
            return None

    def get_elements_in_columns(self, sheet_title, start_row, end_row, column1='A', column2=None):
        range1 = f"{sheet_title}!{column1}{start_row+1}:{column1}{end_row}"
        ranges = [range1]

        if column2:
            range2 = f"{sheet_title}!{column2}{start_row+1}:{column2}{end_row}"
            ranges.append(range2)

        try:
            result = self.service.spreadsheets().values().batchGet(
                spreadsheetId=self.spreadsheet_id,
                ranges=ranges
            ).execute()

            values = [res.get('values', []) for res in result.get('valueRanges', [])]

            column1_values = [item[0] if item else "" for item in values[0]]
            column2_values = [item[0] if item else "" for item in values[1]] if column2 else []

            max_len = max(len(column1_values), len(column2_values))
            column1_values += [""] * (max_len - len(column1_values))
            column2_values += [""] * (max_len - len(column2_values))

            for idx, value in enumerate(column1_values):
                if value == "6:00 AM":
                    column2_values[idx] = "Report Time"

            return column1_values, column2_values

        except HttpError as e:
            if e.resp.status == 429:
                print("Rate limit exceeded. Retrying...")
                time.sleep(60)
                return self.get_elements_in_columns(sheet_title, start_row, end_row, column1, column2)
            else:
                raise

    def calendar_formatted(self):
        date_dict = {}
        date_set = set()
        events = self.get_sheets()
        if events:
            sheets_list, columns, start_rows, end_rows = events
            for i in range(len(sheets_list)):
                sheet_title = sheets_list[i]
                column_letter = chr(64 + columns[i])
                column1_values, column2_values = self.get_elements_in_columns(sheet_title, start_rows[i], end_rows[i], 'A', column_letter)

                for row_idx in range(len(column1_values)):
                    year = str(datetime.date.today().year)
                    month, day = map(int, sheet_title.split('/'))
                    date_str = f"{year}-{month:02}-{day:02}"
                    time_obj = datetime.datetime.strptime(column1_values[row_idx], "%I:%M %p")
                    time_str = time_obj.strftime("%H:%M:%S")
                    datetime_str = f"{date_str}T{time_str}"

                    date_set.add(date_str)

                    if column2_values[row_idx]:
                        if date_str in date_dict:
                            date_dict[date_str].append(f"{datetime_str} | {column2_values[row_idx]}")
                        else:
                            date_dict[date_str] = [f"{datetime_str} | {column2_values[row_idx]}"]

            date_list = list(date_set)
            return [date_list, date_dict]

        return None

    def google_events(self):
        data = self.calendar_formatted()
        if data:
            dates = data[1]
            event_list = []
            for key in dates:
                for i in range(len(dates[key]) - 1):
                    start1 = dates[key][i].split("|")[0].strip()
                    description = "|".join(dates[key][i].split("|")[1:]).strip()
                    end1 = start1.split("T")[0] + "T17:00:00"  # End time set to 5 PM
                    event = {
                        'summary': description,
                        "start": {
                            'dateTime': start1,
                            'timeZone': "America/Chicago",
                        },
                        "end": {
                            'dateTime': end1,
                            'timeZone': "America/Chicago",
                        }
                    }
                    event_list.append(event)
            return event_list
        return None