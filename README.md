## Steps to use
.. Clone this repo
    ```
    git clone https://github.com/maxyong7/LH-Script.git
    ```
.. Install the requirements (if first time using this script)
    ```
    pip install -r requirements.txt
    ```
.. Go to Little Hotelier > Front Desk > Reservations > Select "Status = Confirmed" > Select Dates > Click "Export"
.. Create a new folder within the repo called "import" and import the excel file from Step 3
.. Create another new folder within the repo called "main" and import another excel file from Google Drive.
.. Create an `.env` file with these variables
    ```
    GOOGLE_DOCS_URL=
    NAME_OF_OPERATOR=
    OPERATOR_CONTACT_NUMBER=
    OPERATOR_EMAIL_ADDRESS=
    COMPLETED_STATUS=
    GOOGLE_FORM_STATUS_COLUMN=
    GOOGLE_FORM_DATE_COLUMN=
    CONTACT_EXPORT_STATUS_COLUMN=
    CONTACT_EXPORT_DATE_COLUMN=
    MAIN_EXCEL_FILE_PATH=
    NEW_EXCEL_FILE_PATH=
    ```
.. Run
    To merge the new excel file into the existing excel file
    ```
    python .\merge.py
    ```
    To generate an excel file of contacts to be imported into "google contacts"
    ```
    python .\BulkImportContacts.py
    ```
    To submit google form
    ```
    python .\BulkFillGoogleFormProd.py
    ```
.. Import into https://contacts.google.com/


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

