## Steps to use
1. Clone this repo
2. Install the requirements (if first time using this script)
    ```
    pip install -r requirements.txt
    ```
3. Go to Little Hotelier > Front Desk > Reservations > Select "Status = Confirmed" > Select Dates > Click "Export"
4. Create a new folder within the repo called "env" and import the excel file from Step 3
5. Rename `reservations_file_path` variable to the file name in `BulkImportContacts.py`
    ```
    reservations_file_path = "./env/reservations(4).csv"
    ```
6. Run 
    ```
    python .\BulkImportContacts.py
    ```
7. Import into https://contacts.google.com/


## Sync account in Samsung
1. Go to "Settings" > "Accounts and backup" > "Manage accounts" > "Whatsapp" > "Sync account" > Click the ":" button > "Sync now"


# Steps to isolate python environment within this repository (Optional)
## 1. Create Virtual Environment
```
python3 -m venv <name:which can be venv>
```
## 2. Active Virtual Environment
```
venv\Scripts\activate
```
## 3. Once activated, run
```
python .\BulkImportContacts.py
```

