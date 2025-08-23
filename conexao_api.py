import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import json

google_json = os.environ.get("GOOGLE_JSON")
creds_dict = json.loads(google_json)
scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)
