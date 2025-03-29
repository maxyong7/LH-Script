from datetime import datetime
import os
import pytz
import requests
import pandas as pd
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import dotenv
dotenv.load_dotenv()

url = os.getenv("GOOGLE_DOCS_URL")
excel_file_path = os.getenv("MAIN_EXCEL_FILE_PATH")
nameOfOperator = os.getenv("NAME_OF_OPERATOR")
operatorContactNumber = os.getenv("OPERATOR_CONTACT_NUMBER")
operatorEmailAddress = os.getenv("OPERATOR_EMAIL_ADDRESS")
completedStatus = os.getenv("COMPLETED_STATUS")
googleFormStatusColumn = os.getenv("GOOGLE_FORM_STATUS_COLUMN")
googleFormDateColumn = os.getenv("GOOGLE_FORM_DATE_COLUMN")
completed_counter = 0
failed_counter = 0

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
def send_request(row, file_path, df):
    global completed_counter
    global failed_counter
    numberOfAdults = min(row["number of adults"],8)
    data = f'entry.1478325381={nameOfOperator}&entry.887535986={operatorContactNumber}&entry.1577804325={row["channel name"]}&entry.1080752479={row["guest first name"]+ " " + row["guest last name"]}&entry.1314980540={row["guest phone number"]}&entry.473851324={row["rooms"]}&entry.2006614104={numberOfAdults}&entry.1937352817={row["check in date"]}&entry.929840929={row["check out date"]}&entry.1273587710=None&emailAddress={operatorEmailAddress}'
    encoded_data = data.encode('utf-8')

    # Create a session object
    session = requests.Session()
    session.headers.update({'Content-Type': 'application/x-www-form-urlencoded'})
    try:
        response = session.post(url, data=encoded_data)
        if response.status_code != 200:
            status = f'Failed with status code {response.status_code}'
            print(status)
        else:
            status = completedStatus
    except Exception as e:
        status = f"Error: {str(e)}"
        print(status)

    
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
    print(f'Excel File Path: {excel_file_path}')
    df = read_excel(excel_file_path)
    with ThreadPoolExecutor(max_workers=5) as executor:  # Adjust max_workers as needed
        futures = {}
        for _, row in df.iterrows():
            if row[googleFormStatusColumn] != completedStatus:
                future = executor.submit(send_request, row, excel_file_path, df)
                futures[future] = row
        for future in as_completed(futures):
            future.result()  # Ensures all tasks complete
    
    print(f'Completed: {completed_counter} \nFailed: {failed_counter}')
    return

if __name__ == "__main__":
    main()

