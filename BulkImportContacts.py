import pandas as pd
from datetime import datetime, time

# Load the reservations file
reservations_file_path = "./env/reservations(4).csv"
reservations_df = pd.read_csv(reservations_file_path)

# Display the first few rows of both files
reservations_df.head()

# Define the mapping for channel names
channel_mapping = {
    "Airbnb": "Ab",
    "Booking.com": "Bk",
    "Agoda": "Ag"
}

# Clean and filter necessary columns from reservations data
reservations_filtered = reservations_df[
    ["Guest phone number", "Guest first name", "Guest last name", "Rooms", "Check in date", "Check out date", "Channel name"]
].dropna()

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
    channel_short = channel_mapping.get(row["Channel name"], row["Channel name"])
    check_in = format_date(str(row["Check in date"]))
    check_out = format_date(str(row["Check out date"]))

    return f"Op {row['Rooms']} {channel_short} {check_in}-{check_out}"


# Create new formatted dataframe
contacts_formatted = pd.DataFrame({
    "First Name": reservations_filtered.apply(format_first_name, axis=1),
    "Middle Name": reservations_filtered["Guest first name"],
    "Last Name": reservations_filtered["Guest last name"],
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
    "Phone 1 - Value": reservations_filtered["Guest phone number"].astype(str).apply(lambda x: f"+{x}")
})



# Save to a new CSV file
curr_dt = datetime.now()
timestamp = int(round(curr_dt.timestamp()))
output_file_path = f"./env/formatted_contacts_{timestamp}.csv"
contacts_formatted.to_csv(output_file_path, index=False)

