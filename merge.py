from datetime import datetime
import glob
import os
import sys
import pandas as pd
import pytz
import shutil
import dotenv
dotenv.load_dotenv()


def getImportedCSV(newExcelFilePath:str) -> str:
    # Get all matching files
    matching_files = glob.glob(os.path.join(newExcelFilePath, "*.csv"))

    # Check how many there are
    if len(matching_files) == 1:
        print(f"Found the file: {matching_files[0]}")
        return matching_files[0]
    else:
        print("There should only be 1 csv file in this folder {newExcelFilePath}")
        sys.exit(1)


# Load the old and new Excel files
mainExcelFilePath = os.getenv("MAIN_EXCEL_FILE_PATH")
newExcelFilePath = os.getenv("NEW_EXCEL_FILE_PATH")

# old_reservations_file_path = "./env/march28_2025/merged_file_27032025_195551.csv"
# new_reservations_file_path = "./env/march28_2025/reservations(9).csv"
old_df = pd.read_csv(mainExcelFilePath)
old_df.columns = old_df.columns.str.lower()


importedCSVFilePath = getImportedCSV(newExcelFilePath)
new_df = pd.read_csv(importedCSVFilePath)
new_df.columns = new_df.columns.str.lower()

now = datetime.now(pytz.timezone('Asia/Kuala_Lumpur')).replace(tzinfo=None).date()
temp_checkin = pd.to_datetime(new_df['check in date'], errors='coerce').dt.date
new_df = new_df[temp_checkin >= now] # Only include check-in dates matches Malaysia time now

# Combine the two
combined_df = pd.concat([old_df, new_df], ignore_index=True)

# Remove duplicates (based on all columns)
# If you want to deduplicate by specific columns, list them in the subset parameter
deduped_df = combined_df.drop_duplicates(subset=["booking reference", "guest first name","guest last name"])

deduped_df = deduped_df.sort_values(by=['check in date','guest first name'], ascending=True)

# Save the result back to the old file or a new one
todayDate = datetime.now().strftime("%Y-%m-%d")
timestamp = datetime.now().strftime("%H%M%S")

logsFolder = f"./logs/{todayDate}"
os.makedirs(logsFolder, exist_ok=True)

deduped_df.to_csv(f"{logsFolder}/{todayDate}_merged_file_{timestamp}.csv", index=False)
deduped_df.to_csv(mainExcelFilePath, index=False)

shutil.move(importedCSVFilePath, logsFolder)
print(f"Moved {importedCSVFilePath} -> {logsFolder}")