## Steps to use
1. Clone this repo
    ```
    https://github.com/maxyong7/LH-Script.git
    ```

2. Change directory into this file 
    ```
    cd .\BulkFillGoogleForm
    ```
3. Install the requirements (if first time using this script)
    ```
    pip install -r requirements.txt
    ```
4. Create an `.env` file with these variables
    ```
    GOOGLE_DOCS_URL=
    EXCEL_FILE_PATH=
    NAME_OF_OPERATOR=
    OPERATOR_CONTACT_NUMBER=
    OPERATOR_EMAIL_ADDRESS=
    COMPLETED_STATUS=
    GOOGLE_FORM_STATUS_COLUMN=
    ```
5. Go to Little Hotelier > Front Desk > Reservations > Select "Status = Confirmed" > Select Dates > Click "Export"
6. Import the excel file from Step 4 into this "BulkFillGoogleForm" folder
7. Rename `EXCEL_FILE_PATH` variable to the excel name in `.env` file. Example:
    ```
    EXCEL_FILE_PATH = "./reservations(4).csv"
    ```
8. Run 
    ```
    python .\BulkFillGoogleForm.py
    ```

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

