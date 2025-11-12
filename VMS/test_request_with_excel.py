from datetime import datetime
import os
import pytz
import requests
import pandas as pd
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from html.parser import HTMLParser
import dotenv
import json
import gspread
from google.oauth2.service_account import Credentials
dotenv.load_dotenv()

url = os.getenv("GOOGLE_DOCS_URL")
excel_file_path = os.getenv("MAIN_EXCEL_FILE_PATH")
nameOfOperator = os.getenv("NAME_OF_OPERATOR")
operatorContactNumber = os.getenv("OPERATOR_CONTACT_NUMBER")
operatorEmailAddress = os.getenv("OPERATOR_EMAIL_ADDRESS")
operatorVmsPassword = os.getenv("OPERATOR_VMS_PASSWORD")
completedStatus = os.getenv("COMPLETED_STATUS")
googleFormStatusColumn = os.getenv("GOOGLE_FORM_STATUS_COLUMN")
googleFormDateColumn = os.getenv("GOOGLE_FORM_DATE_COLUMN")
completed_counter = 0
failed_counter = 0
parkingMapPath = os.getenv("PARKING_MAP_PATH")
with open(parkingMapPath, "r", encoding="utf-8") as f:
    parkingMap = json.load(f)

## VMS URLs
VMS_BASE_URL = os.getenv("VMS_BASE_URL")
LOGIN_API = f"{VMS_BASE_URL}/login"
CREATE_VISITOR = f"{VMS_BASE_URL}/admin/visitors"
GET_VISITOR = f"{VMS_BASE_URL}/admin/get-visitors"
SHOW_VISITOR_INFO = f"{VMS_BASE_URL}/admin/visitors"

# Minimal HTML parser that finds hidden input fields by name
class InputFinder(HTMLParser):
    def __init__(self, names=("csrfmiddlewaretoken", "_token", "authenticity_token", "csrf_token")):
        super().__init__()
        self.names = set(names)
        self.found = {}  # {name: value}

    def handle_starttag(self, tag, attrs):
        if tag != "input":
            return
        attrs = dict(attrs)
        name = attrs.get("name")
        typ = attrs.get("type", "").lower()
        if name in self.names and (typ in ("hidden", "", None)):
            val = attrs.get("value")
            if val:
                self.found[name] = val

# Read the Excel file
def read_excel(file_path):
    df = pd.read_csv(file_path) 
    df.columns = df.columns.str.lower()
    if googleFormStatusColumn not in df.columns:
        df[googleFormStatusColumn] = "Pending"  # Ensure 'google form status' column exists and initialized with "Pending"
    if googleFormDateColumn not in df.columns:
        df[googleFormDateColumn] = None  # Ensure 'google form status' column exists and initialized with "Pending"
    
    # Explicitly set the data types to object to handle mixed types (None and strings)
    df[googleFormStatusColumn] = df[googleFormStatusColumn].astype('object')
    df[googleFormDateColumn] = df[googleFormDateColumn].astype('object')
    
    return df

# Write to the Excel file with a lock
lock = threading.Lock()
def update_excel(file_path, df):
    df.to_csv(file_path, index=False)

# Function to send POST request
def send_request(row, file_path, df, session, csrf_field, csrf_value, worksheet):
    global completed_counter
    global failed_counter
    if "2+1" in row["room types"].lower():
        numberOfAdults = min(row["number of adults"],7)
    elif "3" in row["room types"].lower():
        numberOfAdults = min(row["number of adults"],9)
    else:
        numberOfAdults = min(row["number of adults"],4)

    
    visitorEmailAddress = ""
    if pd.notna(row["guest email"]):
        visitorEmailAddress = row["guest email"]
    if visitorEmailAddress == "":
        visitorEmailAddress = "test@gmail.com"

    fullName = f'{row["guest first name"]} {row["guest last name"]}'
   
    carParkLot = parkingMap[f'{row["rooms"]}']
    if carParkLot == "":
        carParkLot = "-"

    ## Create Visitor
    data = {
        "full_name": f'{fullName}',
        "phone": f'{row["guest phone number"]}',
        "email": f'{visitorEmailAddress}',
        "national_identification_no": "-",
        "unit_no": f'{row["rooms"]}',
        "car_park_lot": f'{carParkLot}',
        "booking_source": f'{row["channel name"]}',
    }
    data[csrf_field] = csrf_value
    try:
        createResp = session.post(CREATE_VISITOR, data=data, timeout=30)
        # print("Status:", createResp.status_code)
        # print("Response headers:", createResp.headers)
        # print("Body:", createResp)
    except requests.RequestException as e:
        print("Request failed:", e)

    ## Get Visitor List

    # columns spec from your console, condensed:
    cols = [
        {"data": "reg_no",  "name": "reg_no",      "searchable": "true",  "orderable": "true"},
        {"data": "name",    "name": "name",        "searchable": "true",  "orderable": "true"},
        {"data": "phone",   "name": "phone",       "searchable": "true",  "orderable": "true"},
        {"data": "checkin", "name": "checkin_at",  "searchable": "true",  "orderable": "true"},
        {"data": "checkout","name": "checkout_at", "searchable": "true",  "orderable": "true"},
        {"data": "status",  "name": "status",      "searchable": "true",  "orderable": "true"},
        {"data": "action",  "name": "action",      "searchable": "false", "orderable": "false"},
    ]

    params = {
        "start": 0,
        "length": 10,
        "order[0][column]": 0,
        "order[0][dir]": "desc",
        "order[0][name]": "reg_no",
        "search[value]": f'{fullName}',   # <- your search term
        "search[regex]": "false",
    }
    for i, c in enumerate(cols):
        for k, v in c.items():
            params[f"columns[{i}][{k}]"] = v

    try:
        resp = session.get(GET_VISITOR, params=params, timeout=10)
        jsonResp = resp.json()
        # print("Status:", resp.status_code)
        # print("Response headers:", resp.headers)
        # print("Body:", jsonResp)
        chosen_visitor_id = jsonResp["data"][0]["id"]
    except requests.RequestException as e:
        print("Request failed:", e)


    
    ## Get Specific Visitor Info
    session.headers["X-Requested-With"] = "XMLHttpRequest"
    try:
        resp2 = session.get(f"{SHOW_VISITOR_INFO}/{chosen_visitor_id}/show", timeout=10)
        respData = resp2.json()
        # print("Status:", resp2.status_code)
        # print("Response headers:", resp2.headers)
        # print("Body:", respData)
        qrcode_url = respData["qrcode_url"]
        # print("QR Code URL:", qrcode_url)
        status = completedStatus
    except requests.RequestException as e:
        print("Request failed:", e)

    if status != completedStatus:
        status = "Something went wrong"
    
    print(f'\nName:{row["guest first name"]} {row["guest last name"]} \nPhone number:{row["guest phone number"]} \nRoom type: {row["room types"]} \nNumber of Adults: {numberOfAdults}\nQR Code URL:{qrcode_url}\nStatus: {status}')

    # Update the dataframe with the status
    with lock:
        if status == completedStatus:
            completed_counter += 1
        else:
            failed_counter += 1

        # Save completed date (example: 27/03/2025)
        now = datetime.now(pytz.timezone('Asia/Kuala_Lumpur')).strftime("%d/%m/%Y")
        df.loc[row.name, googleFormStatusColumn] = status
        df.loc[row.name, googleFormDateColumn] = now
        update_excel(file_path, df)

        # Write a header row (A1:H1) in one call
        worksheet.update(range_name="A1:H1", values=[["Full Name", "Phone Number", "Room Number", "Channel Name", "Check In Date", "Check Out Date", "QR Code URL", "Created At"]])

        # Append a row to google sheet
        worksheet.append_row(
            [
                f'{fullName}', 
                f'{row["guest phone number"]}',
                f'{row["rooms"]}',
                f'{row["channel name"]}',
                f'{row["check in date"]}',
                f'{row["check out date"]}',
                f'{qrcode_url}',
                f'{now}',
            ],
            value_input_option="RAW")

def cleanup_old_google_sheet_rows(worksheet):
    """
    Delete rows from the worksheet where 'Created At' column is more than 7 days old
    """
    try:
        print("\nStarting cleanup of old rows in Google Sheet...")
        # Get all records from the worksheet
        all_records = worksheet.get_all_records()
        
        if not all_records:
            print("\nNo records found in worksheet to clean up")
            return
        
        # Get current date in KL timezone
        current_date = datetime.now(pytz.timezone('Asia/Kuala_Lumpur'))
        rows_to_delete = []
        
        # Check each row (starting from row 2, since row 1 is header)
        for i, record in enumerate(all_records, start=2):
            checkout_date_str = record.get('Check Out Date', '')
            
            if checkout_date_str:
                try:
                    # Parse the date (assuming format is DD/MM/YYYY)
                    created_at = datetime.strptime(checkout_date_str, "%d/%m/%Y")
                    created_at = pytz.timezone('Asia/Kuala_Lumpur').localize(created_at)
                    
                    # Calculate the difference
                    days_diff = (current_date - created_at).days
                    
                    # If more than 7 days old, mark for deletion
                    if days_diff > 1:
                        rows_to_delete.append(i)
                        print(f"\nMarking row {i} for deletion - Created: {checkout_date_str}, Days old: {days_diff} \n{record}")
                        
                except ValueError as e:
                    print(f"Error parsing date '{checkout_date_str}' in row {i}: {e}")
        
        # Delete rows in reverse order to maintain correct row indices
        for row_num in reversed(rows_to_delete):
            try:
                worksheet.delete_rows(row_num)
                print(f"Deleted row number {row_num}")
            except Exception as e:
                print(f"Error deleting row {row_num}: {e}")
        
        if rows_to_delete:
            print(f"Successfully deleted {len(rows_to_delete)} old rows")
        else:
            print("No old rows found to delete")
            
    except Exception as e:
        print(f"Error during cleanup: {e}")

def login_google_and_get_worksheet():
    print("\nLogging into Google Sheets...")
    
    # Scopes: spreadsheets read/write
    SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]

    creds_path = os.environ["GOOGLE_APPLICATION_CREDENTIALS"]
    spreadsheet_id = os.environ["SHEET_ID"]

    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPE)
    client = gspread.authorize(creds)

    # Open the spreadsheet and select the first worksheet
    sh = client.open_by_key(spreadsheet_id)
    ws = sh.sheet1  # or sh.worksheet("Sheet1")

    return ws

def login_vms_and_get_token(session):
    print("\nLogging into Opus VMS to get tokens...")
    creds = {"email": f"{operatorEmailAddress}", "password": f"{operatorVmsPassword}"}
    # 1) Authenticate
    r = session.get(LOGIN_API)
    r.raise_for_status()
    
    parser = InputFinder()
    parser.feed(r.text)

    # Try hidden input first
    csrf_field, csrf_value = None, None
    if parser.found:
        csrf_field, csrf_value = next(iter(parser.found.items()))
        creds[csrf_field] = csrf_value
        # print("Found CSRF token in login page:", csrf_field, csrf_value)

    # If the server sets auth cookies (e.g., Set-Cookie: session=...; HttpOnly),
    # requests.Session will store them automatically in s.cookies.
    # Some APIs also return a CSRF token in headers/body; add it if needed:
    r = session.post(LOGIN_API, json=creds, timeout=15)
    r.raise_for_status()

    
    parser = InputFinder()
    parser.feed(r.text)

    # Refresh CSRF token after login
    csrf_field, csrf_value= None, None
    if parser.found:
        field, value = next(iter(parser.found.items()))
        if field == "_token":
            csrf_field, csrf_value = field, value
            # print("Found CSRF token after logged in:", csrf_field, csrf_value)
    return csrf_field, csrf_value


# Main function to read, process, and update Excel file
def main():
    with requests.Session() as s:
        print(f'Excel File Path: {excel_file_path}')
        df = read_excel(excel_file_path)
        if df is None:
            print("Failed to read Excel file. Exiting.")
            return

        csrf_field, csrf_value = login_vms_and_get_token(session=s)
        if not csrf_field or not csrf_value:
            print("Failed to retrieve CSRF token from VMS. Exiting.")
            return
        
        ws = login_google_and_get_worksheet()
        if not ws:
            print("Failed to access Google Sheet worksheet. Exiting.")
            return

        with ThreadPoolExecutor(max_workers=5) as executor:  # Adjust max_workers as needed
            futures = {}
            for _, row in df.iterrows():
                if row[googleFormStatusColumn] != completedStatus:
                    future = executor.submit(send_request, row, excel_file_path, df, worksheet=ws,
                                             session=s, csrf_field=csrf_field, csrf_value=csrf_value)
                    futures[future] = row
            for future in as_completed(futures):
                future.result()  # Ensures all tasks complete
        
        print(f'\nCompleted: {completed_counter} \nFailed: {failed_counter}')
        
        # Clean up google sheet rows older than 7 days
        cleanup_old_google_sheet_rows(ws)
        return


if __name__ == "__main__":
    main()

