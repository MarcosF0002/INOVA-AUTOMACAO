import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Lê o JSON do secret
google_json = os.environ.get("GOOGLE_JSON")
if not google_json:
    raise ValueError("O secret GSHEETS_CREDENTIALS_JSON não está definido!")

creds_dict = json.loads(google_json)

# Define o escopo do Google Sheets
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]

# Cria credenciais a partir do dicionário
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

# Autoriza e cria o cliente do gspread
client = gspread.authorize(creds)
