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
    return df

# Write to the Excel file with a lock
lock = threading.Lock()
def update_excel(file_path, df):
    df.to_csv(file_path, index=False)

# Function to send POST request
def send_request(row, file_path, df, session, csrf_field, csrf_value):
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
        "car_park_lot": "p1",
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

# Main function to read, process, and update Excel file
def main():
    creds = {"email": f"{operatorEmailAddress}", "password": f"{operatorVmsPassword}"}
    with requests.Session() as s:
        s.headers.update({
            "User-Agent": "Mozilla/5.0 (requests)",
            "Accept": "application/json",
        })

        # 1) Authenticate
        r = s.get(LOGIN_API)
        r.raise_for_status()
        
        parser = InputFinder()
        parser.feed(r.text)

        # Try hidden input first
        csrf_field, csrf_value = None, None
        if parser.found:
            csrf_field, csrf_value = next(iter(parser.found.items()))
            creds[csrf_field] = csrf_value
            print("Found CSRF token in login page:", csrf_field, csrf_value)

        # If the server sets auth cookies (e.g., Set-Cookie: session=...; HttpOnly),
        # requests.Session will store them automatically in s.cookies.
        # Some APIs also return a CSRF token in headers/body; add it if needed:
        r = s.post(LOGIN_API, json=creds, timeout=15)
        r.raise_for_status()

        
        parser = InputFinder()
        parser.feed(r.text)

        # Refresh CSRF token after login
        csrf_field, csrf_value= None, None
        if parser.found:
            field, value = next(iter(parser.found.items()))
            if field == "_token":
                csrf_field, csrf_value = field, value
                print("Found CSRF token in login page:", csrf_field, csrf_value)

        print(f'Excel File Path: {excel_file_path}')
        df = read_excel(excel_file_path)
        with ThreadPoolExecutor(max_workers=1) as executor:  # Adjust max_workers as needed
            futures = {}
            for _, row in df.iterrows():
                if row[googleFormStatusColumn] != completedStatus:
                    future = executor.submit(send_request, row, excel_file_path, df, 
                                             session=s ,csrf_field=csrf_field, csrf_value=csrf_value)
                    futures[future] = row
            for future in as_completed(futures):
                future.result()  # Ensures all tasks complete
        
        print(f'Completed: {completed_counter} \nFailed: {failed_counter}')
        return

if __name__ == "__main__":
    main()

