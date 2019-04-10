import gspread
from oauth2client.service_account import ServiceAccountCredentials

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(creds)

# Open a worksheet from spreadsheet with one shot
wks = gc.open("CS_PG_Curriculum").sheet1

list_of_hashes = sheet.get_all_records()
print(list_of_hashes)
