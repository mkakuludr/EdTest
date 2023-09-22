import gspread
import sys
import boto3
import json
from gspread.exceptions import CellNotFound
from botocore.exceptions import ClientError

SPREADSHEET_KEY = "1ixJTLq4KiwafYuwxh7LObcX6vsDHoQTRfCNUcJTgePQ"
SHEET_NAME = "Test"
SEARCH_COLUMN = "B"
SEARCH_VALUE = sys.argv[2]
WRITE_COLUMN = "D"
WRITE_VALUE = sys.argv[1]
AWS_PROFILE = "shared-DevOps"
KEY = "gspread-cred"

def initialize_secrets_client():
    boto3.setup_default_session(profile_name=AWS_PROFILE)
    return boto3.client(service_name='secretsmanager', region_name="us-east-1")

def get_secret(key_name):
    client = initialize_secrets_client()
    try:
        get_secret_value_response = client.get_secret_value(SecretId="platform/moxi-roads")
        if 'SecretString' in get_secret_value_response:
            secret = get_secret_value_response['SecretString']
            secret_json = json.loads(secret)
            if key_name in secret_json:
                return secret_json[key_name]
            else:
                print(f"Key '{key_name}' not found in the secret JSON.")
                return None
        else:
            print("SecretString not found in the secret response.")
            return None
    except Exception as e:
        print(f"Error retrieving secret: {e}")
        return None

credentials = get_secret(KEY)
credentials_dict = json.loads(credentials)

gc = gspread.service_account_from_dict(credentials_dict)

sh = gc.open_by_key(SPREADSHEET_KEY)

sheet = sh.worksheet(SHEET_NAME)

try:
    # Find the row with the search value in the search column
    print(sheet.findall("16",in_column=2))
    cell_list = sheet.findall(SEARCH_VALUE, in_column=2)
    if cell_list:
        for cell in cell_list:
            row = cell.row
            try:
                col = 2
                write_cell = sheet.cell(row, col + 1)  # Next column
                write_cell.value = WRITE_VALUE
                sheet.update_cells([write_cell])
                print(f"Value '{WRITE_VALUE}' written to row {row}, column {col + 1}")
            except CellNotFound:
                print(f"Write column '{WRITE_COLUMN}' not found.")
    else:
        print("Search value not found in the specified column. Adding a new row...")
        
        # Append a new row with the search value and write value
        new_row = ["",SEARCH_VALUE, WRITE_VALUE]
        sheet.append_rows([new_row])
        print(f"New row added with Search value '{SEARCH_VALUE}' and Write value '{WRITE_VALUE}'")
except CellNotFound:
    print(f"Search column '{SEARCH_COLUMN}' not found.")