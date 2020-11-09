import pickle
import os.path
import os
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from dotenv import load_dotenv

#Carregando variáveis de ambiente
load_dotenv()
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")

SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

def main():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    #Chamar a Sheets API para pegar as informações de pedidos e estoque
    sheet = sheets_service.spreadsheets()
    orders_result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='Pedidos',
        majorDimension='COLUMNS').execute()
    orders = orders_result.get('values', [])
    d = {col[0]: col[1:] for col in orders}    
    orders_df = pd.DataFrame(data=d)

if __name__ == "__main__":
    main()