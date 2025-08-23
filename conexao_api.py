import os, json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Lê o JSON do Secret do GitHub
google_json = os.getenv("GOOGLE_JSON")
creds_dict = json.loads(google_json)

scopes = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scopes)
client = gspread.authorize(creds)

