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
import io
dotenv.load_dotenv()

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
BULK_IMPORT_URL = f"{VMS_BASE_URL}/app/guests/import/upload"

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

# Transform dataframe to VMS format with required fields
def transform_dataframe(df, parkingMap):
    """
    Transform the original dataframe into VMS-specific format with mapped fields.
    
    Args:
        df: Original dataframe from Excel/CSV
        parkingMap: Parking map dictionary for car park lot lookup
    
    Returns:
        New dataframe with VMS-specific columns
    """
    transformed_rows = []
    
    for idx, row in df.iterrows():
        # Get visitor email
        visitorEmailAddress = ""
        if pd.notna(row["guest email"]):
            visitorEmailAddress = row["guest email"]
        if visitorEmailAddress == "":
            visitorEmailAddress = "test@gmail.com"
        
        # Get car park lot from parking map
        carParkLot = parkingMap.get(str(row["rooms"]), "")
        if carParkLot == "":
            carParkLot = "-"

        # Determine number of adults based on room types
        if "2+1" in row["room types"].lower():
            numberOfAdults = min(row["number of adults"],7)
        elif "3" in row["room types"].lower():
            numberOfAdults = min(row["number of adults"],9)
        else:
            numberOfAdults = min(row["number of adults"],4)
        
        # Helper function to parse and format dates with specific time
        def parse_date_with_time(date_value, hour, date_name):
            try:
                if pd.notna(date_value):
                    if isinstance(date_value, str):
                        parsed_date = pd.to_datetime(date_value, dayfirst=True)
                    else:
                        parsed_date = date_value
                    datetime_with_time = parsed_date.replace(hour=hour, minute=0, second=0)
                    return datetime_with_time.strftime("%Y-%m-%d %H:%M:%S")
                return ""
            except Exception as e:
                print(f"Error parsing {date_name} for row {idx}: {e}")
                return str(date_value)
        
        # Parse check-in (3pm) and check-out (11am) dates
        planned_checkin = parse_date_with_time(row["check in date"], 15, "check-in date")
        planned_checkout = parse_date_with_time(row["check out date"], 11, "check-out date")
        
        # Create transformed row with VMS fields
        transformed_row = {
            "name": f'{row["guest first name"]} {row["guest last name"]}',
            "phone": str(row["guest phone number"]),
            "email": visitorEmailAddress,
            "IC/Passport No": "-",  # Required field, but setting it as empty
            "unit_no": str(row["rooms"]),
            "car_park_lot": carParkLot,
            "booking_source": str(row["channel name"]),
            "note": "",  # Optional field, keeping as empty
            "number_of_pax": f'{numberOfAdults}',  # Optional field, keeping as empty
            "planned_checkin_at": planned_checkin,
            "planned_checkout_at": planned_checkout,
            "deposit_collected": "",  # Optional field, keeping as empty
            "deposit_amount": "",  # Optional field, keeping as empty
            "deposit_currency": "",  # Optional field, keeping as empty
        }
        transformed_rows.append(transformed_row)
    
    # Create new dataframe
    transformed_df = pd.DataFrame(transformed_rows)
    
    # Set original index to maintain row.name references
    transformed_df.index = df.index
    
    return transformed_df

# Write to the Excel file with a lock
lock = threading.Lock()
def update_excel(file_path, df):
    df.to_csv(file_path, index=False)

# Function to send bulk CSV upload request
def send_request(transformed_df, file_path, original_df, session, csrf_value):
    """
    Bulk upload guests to VMS via CSV file upload.
    
    Args:
        transformed_df: Dataframe with VMS-formatted guest data
        file_path: Path to the original Excel file for status updates
        original_df: Original dataframe to update with status
        session: requests.Session with authenticated cookies
        csrf_value: CSRF token value for request
    """
    global completed_counter
    global failed_counter
    
    print("\nPreparing bulk CSV upload...")
    
    # Create CSV in memory from transformed dataframe
    csv_buffer = io.StringIO()
    transformed_df.to_csv(csv_buffer, index=False)
    csv_content = csv_buffer.getvalue()
    csv_bytes = io.BytesIO(csv_content.encode('utf-8'))
    
    # Generate filename with timestamp
    todayDate = datetime.now().strftime("%Y-%m-%d")
    timestamp = datetime.now().strftime("%H%M%S")
    filename = f"{todayDate}_merged_file_{timestamp}.csv"
    
    # Prepare multipart/form-data payload
    files = {
        'file': (filename, csv_bytes, 'application/vnd.ms-excel')
    }
    
    data = {
        '_token': csrf_value,
        'send_notifications': '0'
    }
    
    # Set required headers
    headers = {
        'X-CSRF-TOKEN': csrf_value,
        'X-Requested-With': 'XMLHttpRequest'
    }
    
    try:
        print(f"\nUploading {len(transformed_df)} guests to VMS...")
        response = session.post(
            BULK_IMPORT_URL,
            files=files,
            data=data,
            headers=headers,
            timeout=60
        )
        
        # print(f"Response Status: {response.status_code}")
        
        # Check if upload was successful
        if response.status_code == 200:
            try:
                response_data = response.json()
                # print(f"Response: {response_data}")
                
                # Validate import results
                if response_data.get('success'):
                    results = response_data.get('results', {})
                    total_records = results.get('total_records', 0)
                    successful_imports = results.get('successful_imports', 0)
                    failed_imports = results.get('failed_imports', 0)
                    
                    print(f"\n--- Import Summary ---")
                    print(f"Expected records: {len(transformed_df)}")
                    print(f"Total records processed: {total_records}")
                    print(f"Successful imports: {successful_imports}")
                    print(f"Failed imports: {failed_imports}")
                    
                    # Check if all records were processed
                    if total_records != len(transformed_df):
                        missing_count = len(transformed_df) - total_records
                        print(f"\n⚠ WARNING: {missing_count} record(s) missing from import!")
                        
                        # Get the row numbers from success_rows and error_rows
                        success_rows = results.get('success_rows', [])
                        error_rows = results.get('error_rows', [])
                        skipped_rows = results.get('skipped_rows', [])
                        
                        processed_row_numbers = set()
                        for row_data in success_rows + error_rows + skipped_rows:
                            if 'row' in row_data:
                                processed_row_numbers.add(row_data['row'])
                        
                        # CSV rows start at 2 (row 1 is header)
                        expected_row_numbers = set(range(2, len(transformed_df) + 2))
                        missing_row_numbers = expected_row_numbers - processed_row_numbers
                        
                        if missing_row_numbers:
                            print(f"\nMissing CSV row numbers: {sorted(missing_row_numbers)}")
                            print("\n--- Missing Records ---")
                            for csv_row_num in sorted(missing_row_numbers):
                                # Convert CSV row number to dataframe index (CSV row 2 = df index 0)
                                df_index = csv_row_num - 2
                                if df_index < len(transformed_df):
                                    missing_record = transformed_df.iloc[df_index]
                                    print(f"\nCSV Row {csv_row_num} (DataFrame index {df_index}):")
                                    print(f"  Name: {missing_record['name']}")
                                    print(f"  Phone: {missing_record['phone']}")
                                    print(f"  Email: {missing_record['email']}")
                                    print(f"  Unit: {missing_record['unit_no']}")
                                    print(f"  Check-in: {missing_record['planned_checkin_at']}")
                    
                    # Show failed imports if any
                    if failed_imports > 0:
                        print(f"\n--- Failed Imports ({failed_imports}) ---")
                        for error_row in results.get('error_rows', []):
                            print(f"\nRow {error_row.get('row')}: {error_row.get('name', 'N/A')}")
                            print(f"  Error: {error_row.get('error', 'Unknown error')}")
                    
                    # Determine overall status
                    if successful_imports == len(transformed_df):
                        status = completedStatus
                        print(f"\n✓ All {len(transformed_df)} guests successfully uploaded")
                    elif successful_imports > 0:
                        status = f"Partial success - {successful_imports}/{len(transformed_df)} imported"
                        print(f"\n⚠ Partial success: {successful_imports}/{len(transformed_df)} guests imported")
                    else:
                        status = "Upload failed - No successful imports"
                        print(f"\n✗ Upload failed: No guests were imported")
                else:
                    status = "Upload failed - Server returned error"
                    print(f"\n✗ Upload failed: {response_data.get('message', 'Unknown error')}")
                
            except json.JSONDecodeError:
                # If response is not JSON, check text for success indicators
                if 'success' in response.text.lower() or response.status_code == 200:
                    status = completedStatus
                    print(f"\n✓ Successfully uploaded {len(transformed_df)} guests")
                else:
                    status = "Upload failed - Invalid response"
                    print(f"\n✗ Upload failed: Invalid response format")
        else:
            status = f"Upload failed - Status {response.status_code}"
            print(f"\n✗ Upload failed with status {response.status_code}")
            print(f"Response: {response.text}")
            
    except requests.RequestException as e:
        status = f"Upload failed - {str(e)}"
        print(f"\n✗ Request failed: {e}")
    
    # Update all rows in original dataframe with status
    with lock:
        now = datetime.now(pytz.timezone('Asia/Kuala_Lumpur')).strftime("%d/%m/%Y")
        
        if status == completedStatus:
            completed_counter = len(transformed_df)
            print(f"\nUpdating {len(transformed_df)} rows with completed status...")
            
            # Update original dataframe
            for idx in transformed_df.index:
                original_df.loc[idx, googleFormStatusColumn] = status
                original_df.loc[idx, googleFormDateColumn] = now
            
            # Save updated dataframe
            update_excel(file_path, original_df)
        else:
            failed_counter = len(transformed_df)
            print(f"\n✗ All {len(transformed_df)} guests marked as failed")
            
            # Update original dataframe with failed status
            for idx in transformed_df.index:
                original_df.loc[idx, googleFormStatusColumn] = status
                original_df.loc[idx, googleFormDateColumn] = now
            
            update_excel(file_path, original_df)
    
    print(f"\nBulk upload completed: {completed_counter} succeeded, {failed_counter} failed")

def cleanup_old_excel_rows():
    """
    Delete rows from the Excel file where 'check out date' is more than 30 days old
    """
    try:
        print("\nStarting cleanup of old rows in Excel file...")
        
        # Read the Excel file
        df = pd.read_csv(excel_file_path)
        df.columns = df.columns.str.lower()
        
        if 'check out date' not in df.columns:
            print("\n'check out date' column not found in Excel file")
            return
        
        initial_row_count = len(df)
        
        # Get current date in KL timezone
        current_date = datetime.now(pytz.timezone('Asia/Kuala_Lumpur'))
        
        # Track rows to keep
        rows_to_keep = []
        deleted_count = 0
        
        # Check each row
        for idx, row in df.iterrows():
            checkout_date_str = row.get('check out date', '')
            
            if pd.notna(checkout_date_str) and checkout_date_str != '':
                try:
                    # Parse the checkout date
                    if isinstance(checkout_date_str, str):
                        # Try multiple date formats
                        checkout_date = None
                        for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"]:
                            try:
                                checkout_date = datetime.strptime(checkout_date_str, fmt)
                                break
                            except ValueError:
                                continue
                        
                        if checkout_date is None:
                            # If all formats fail, try pandas
                            checkout_date = pd.to_datetime(checkout_date_str, dayfirst=True)
                    else:
                        checkout_date = pd.to_datetime(checkout_date_str)
                    
                    # Localize to KL timezone
                    checkout_date = pytz.timezone('Asia/Kuala_Lumpur').localize(checkout_date)
                    
                    # Calculate the difference in days
                    days_diff = (current_date - checkout_date).days
                    
                    # If more than 30 days old, skip this row (delete it)
                    if days_diff > 30:
                        deleted_count += 1
                        print(f"Deleting row {idx}: Checkout date: {checkout_date_str}, Days old: {days_diff}")
                        continue
                    
                except Exception as e:
                    print(f"Error parsing date '{checkout_date_str}' in row {idx}: {e}")
            
            # Keep this row
            rows_to_keep.append(idx)
        
        # Filter dataframe to keep only valid rows
        df_filtered = df.loc[rows_to_keep]
        
        # Save the filtered dataframe back to the Excel file
        if deleted_count > 0:
            df_filtered.to_csv(excel_file_path, index=False)
            print(f"\nSuccessfully deleted {deleted_count} old rows from Excel file")
            print(f"Rows before cleanup: {initial_row_count}, Rows after cleanup: {len(df_filtered)}")
        else:
            print("\nNo old rows found to delete (all checkout dates are within 30 days)")
            
    except Exception as e:
        print(f"Error during cleanup: {e}")

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

        # Transform dataframe to VMS format
        print("Transforming dataframe to VMS format...")
        transformed_df = transform_dataframe(df, parkingMap)

        
        # Save the result back to the old file or a new one
        todayDate = datetime.now().strftime("%Y-%m-%d")
        timestamp = datetime.now().strftime("%H%M%S")

        logsFolder = f"./logs/{todayDate}"
        os.makedirs(logsFolder, exist_ok=True)

        transformed_df.to_csv(f"{logsFolder}/{todayDate}_opusvms_bulkimport_{timestamp}.csv", index=False)
        
        # Login to VMS and get CSRF token
        csrf_field, csrf_value = login_vms_and_get_token(session=s)
        if not csrf_field or not csrf_value:
            print("Failed to retrieve CSRF token from VMS. Exiting.")
            return
        
        # Perform bulk CSV upload (single request for all guests)
        send_request(
            transformed_df=transformed_df,
            file_path=excel_file_path,
            original_df=df,
            session=s,
            csrf_value=csrf_value,
        )
        
        print(f'\nCompleted: {completed_counter} \nFailed: {failed_counter}')
        
        # Clean up google sheet rows older than 7 days
        cleanup_old_excel_rows()
        return


if __name__ == "__main__":
    main()

