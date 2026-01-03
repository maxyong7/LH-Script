import os
import sys
import pandas as pd
from datetime import datetime, time
import dotenv
import pytz
dotenv.load_dotenv()

completedStatus = os.getenv("COMPLETED_STATUS")
contactExportStatusColumn = os.getenv("CONTACT_EXPORT_STATUS_COLUMN")
contactExportDateColumn = os.getenv("CONTACT_EXPORT_DATE_COLUMN")

# Load the reservations file
reservations_file_path = os.getenv("MAIN_EXCEL_FILE_PATH")
reservations_df = pd.read_csv(reservations_file_path)

if "contact export status" not in reservations_df.columns:
    reservations_df["contact export status"] = None
if "contact export date" not in reservations_df.columns:
    reservations_df["contact export date"] = None

# Filter out rows where the status column equals the completed status
exclude_completed_status = reservations_df[reservations_df[contactExportStatusColumn] != completedStatus]

# Define the mapping for channel names
channel_mapping = {
    "Airbnb": "Ab",
    "Booking.com": "Bk",
    "Agoda": "Ag",
    "Extranet": "Dr",
    "Mobile App": "Dr",
    "Trip.com(New)": "Tp"
}

# Clean and filter necessary columns from reservations data
reservations_filtered = exclude_completed_status[
    ["guest phone number", "guest first name", "guest last name", "rooms", "check in date", "check out date", "channel name"]
]

# Function to format check-in and check-out dates
def format_date(date_str):
    try:
        date_obj = datetime.strptime(date_str, "%d/%m/%Y")
        return date_obj.strftime("%d%b").lower()
    except:
        pass
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.strftime("%d%b").lower()
    except ValueError:
        return date_str  # Return as is if there's an issue

# Function to format the first name column
def format_first_name(row):
    channel_short = channel_mapping.get(row["channel name"], row["channel name"])
    check_in = format_date(str(row["check in date"]))
    check_out = format_date(str(row["check out date"]))

    return f"Op {row['rooms']} {channel_short} {check_in}-{check_out}"

def format_phone_number(phone_number):
    if pd.isna(phone_number):
        return "-"
    phone_number_str = str(phone_number).strip()
    if phone_number_str == "":
        return "-"
    return f"+{phone_number_str}"

def print_formatted_contacts(row):
    print(f"{row['First Name']} {row['Middle Name']} {row['Last Name']}")

# Split into rows with and without phone numbers
with_phone = reservations_filtered[reservations_filtered["guest phone number"].notna()]
missing_phone = reservations_filtered[reservations_filtered["guest phone number"].isna()]


# Create new formatted dataframe
contacts_with_phone  = pd.DataFrame({
    "First Name": with_phone.apply(format_first_name, axis=1),
    "Middle Name": with_phone["guest first name"],
    "Last Name": with_phone["guest last name"],
    "Phonetic First Name": None,
    "Phonetic Middle Name": None,
    "Phonetic Last Name": None,
    "Name Prefix": None,
    "Name Suffix": None,
    "Nickname": None,
    "File As": None,
    "Organization Name": None,
    "Organization Title": None,
    "Organization Department": None,
    "Birthday": None,
    "Notes": None,
    "Photo": None,
    "Labels": "* myContacts",
    "Phone 1 - Label": "Mobile",
    "Phone 1 - Value": with_phone["guest phone number"].apply(format_phone_number)
})

# Format contacts missing phone numbers (just for tracking)
contacts_missing_phone = pd.DataFrame({
    "First Name": missing_phone.apply(format_first_name, axis=1),
    "Middle Name": missing_phone["guest first name"],
    "Last Name": missing_phone["guest last name"]
})

# Print the exported contacts
print(f"\n=== Contacts WITH Phone Numbers (total {len(contacts_with_phone)}) ===")
contacts_with_phone.apply(print_formatted_contacts, axis=1)

print(f"\n=== Contacts MISSING Phone Numbers (total {len(contacts_missing_phone)}) === ===")
contacts_missing_phone.apply(print_formatted_contacts, axis=1)

processed_contacts = pd.DataFrame({
    "First Name": reservations_filtered.apply(format_first_name, axis=1),
    "Middle Name": reservations_filtered["guest first name"],
    "Last Name": reservations_filtered["guest last name"],
    "Phonetic First Name": None,
    "Phonetic Middle Name": None,
    "Phonetic Last Name": None,
    "Name Prefix": None,
    "Name Suffix": None,
    "Nickname": None,
    "File As": None,
    "Organization Name": None,
    "Organization Title": None,
    "Organization Department": None,
    "Birthday": None,
    "Notes": None,
    "Photo": None,
    "Labels": "* myContacts",
    "Phone 1 - Label": "Mobile",
    "Phone 1 - Value": reservations_filtered["guest phone number"].apply(format_phone_number)
})

# Print how many contacts were processed. Including contacts without phone number
total_processed_indices = processed_contacts.index
print(f"\nNumber of processed rows for contacts: {len(total_processed_indices)}")

# Update rows that were exported as contacts to "completed"
reservations_df.loc[total_processed_indices, contactExportStatusColumn] = completedStatus

if len(total_processed_indices) == 0:
    sys.exit(0)

# Save completed date (example: 27/03/2025)
now = datetime.now(pytz.timezone('Asia/Kuala_Lumpur')).strftime("%d/%m/%Y")
reservations_df.loc[total_processed_indices, contactExportDateColumn] = now
reservations_df.to_csv(reservations_file_path, index=False)


# Save to a new CSV file
todayDate = datetime.now().strftime("%Y-%m-%d")

logsFolder = f"./logs/{todayDate}"
timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")

output_file_path = f"{logsFolder}/formatted_contacts_{timestamp}.csv"
processed_contacts.to_csv(output_file_path, index=False)

