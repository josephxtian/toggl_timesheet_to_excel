import os
import dotenv
import datetime as dt
import requests
from requests.auth import HTTPBasicAuth
from openpyxl import load_workbook
from collections import defaultdict
from dateutil import parser

COL_DATE = "A"
COL_MORNING_IN = "B"
COL_MORNING_OUT = "C"
COL_AFTERNOON_IN = "D"
COL_AFTERNOON_OUT = "E"
START_ROW = 11

dotenv.load_dotenv()
workspace_id = os.getenv("WORKSPACE_ID")
user_agent = os.getenv("TOGGL_EMAIL")
api_token = os.getenv("TOGGL_API_TOKEN")
excel_path = os.getenv("EXCEL_PATH")

def find_last_filled_row(ws):
    row = START_ROW
    print(ws[f"{COL_MORNING_IN}{row}"].value)
    while ws[f"{COL_MORNING_IN}{row}"].value:
        row += 1
    return row - 1

def get_fetch_range(ws, last_row):
    if last_row < START_ROW:
        start_date = ws[f"{COL_DATE}{START_ROW}"].value
    else:
        print(ws[f"{COL_DATE}{last_row}"].value)
        start_date = ws[f"{COL_DATE}{last_row}"].value + dt.timedelta(days=1)

    end_date = dt.datetime.today()
    return start_date, end_date

def fetch_toggl_entries(start_date:dt.date,end_date:dt.date):
    url = r"https://api.track.toggl.com/reports/api/v2/details"
    params = {
        "workspace_id": workspace_id,
        "since": start_date.date().isoformat(),
        "until": end_date.date().isoformat(),
        "user_agent": user_agent,
        "page": 1
    }

    entries = []

    while True:
        r = requests.get(
            url,
            params=params,
            auth=HTTPBasicAuth(api_token,"api_token")
        )
        r.raise_for_status()
        data = r.json()
        print(data)

        entries.extend(data["data"])

        if params["page"] * data["per_page"] >= data["total_count"]:
            break
        params["page"] +=1
    
    return entries

def group_entries_by_date(entries) -> dt.datetime:
    days = defaultdict(list)

    for e in entries:
        start = parser.isoparse(e["start"])
        end = parser.isoparse(e["end"])
        days[start.date()].append((start,end))

    return days

def write_times(ws, start_row, grouped_entires:dt.datetime):
    row = start_row

    for day in sorted(grouped_entires.keys()):
        blocks = sorted(grouped_entires[day], key=lambda x: x[0])

        if len(blocks) != 2:
            raise ValueError(f"{day} does not have exactly two time blocks")
    
        (m_in, m_out), (a_in, a_out) = blocks

        ws[f"{COL_DATE}{row}"] = day
        ws[f"{COL_MORNING_IN}{row}"] = m_in.strftime("%H:%M")
        ws[f"{COL_MORNING_OUT}{row}"] = m_out.strftime("%H:%M")
        ws[f"{COL_AFTERNOON_IN}{row}"] = a_in.strftime("%H:%M")
        ws[f"{COL_AFTERNOON_OUT}{row}"] = a_out.strftime("%H:%M")

        row += 1

def main():
    wb = load_workbook(excel_path)
    print("success")
    ws = wb.active
    print(ws)

    last_row = find_last_filled_row(ws)
    start_date, end_date = get_fetch_range(ws, last_row)

    if start_date > end_date:
        print("Nothing to fill.")
        return
    
    entries = fetch_toggl_entries(start_date,end_date)
    grouped = group_entries_by_date(entries)

    write_times(ws, last_row + 1, grouped)

    wb.save(excel_path)

    print("Timesheet updated successfully.")

if __name__ == "__main__":
    main()