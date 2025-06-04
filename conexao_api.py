import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

filename = "C:\\Users\\marco\\OneDrive\\Área de Trabalho\\Economia\\INOVA\\credenciais\\inova-461311-6f78bcbbcfc8.json"

scopes = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

creds = ServiceAccountCredentials.from_json_keyfile_name(
    filename = filename,
    scopes = scopes
)

client = gspread.authorize(creds)
print(client)

planilha_Startups = client.open(
    title = "PORTAL DA INOVAÇÃO E STARTUPS", 
    folder_id="1JktHXPDmaY3xQGlXAkRWx2z4ZmLOXCCY",
    )

