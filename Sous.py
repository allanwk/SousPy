import pickle
import os.path
import os
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from dotenv import load_dotenv
from docs_util import read_structural_elements

#Carregando variáveis de ambiente
load_dotenv()
ORDERS_SPREADSHEET_ID = os.environ.get("ORDERS_SPREADSHEET_ID")
STOCK_SPREADSHEET_ID = os.environ.get("STOCK_SPREADSHEET_ID")
RECIPES_DIR_ID = os.environ.get("RECIPES_DIR_ID")

SCOPES = ['https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/documents.readonly']

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
    docs_service = build('docs', 'v1', credentials=creds)

    #Chamar a Sheets API para pegar as informações de pedidos
    sheet = sheets_service.spreadsheets()
    orders_result = sheet.values().get(
        spreadsheetId=ORDERS_SPREADSHEET_ID,
        range='Pedidos',
        majorDimension='COLUMNS').execute()
    orders = orders_result.get('values', [])
    d = {col[0]: col[1:] for col in orders}    
    orders_df = pd.DataFrame(data=d)

    #Obtendo informações do estoque
    stock_result = sheet.values().get(
        spreadsheetId=STOCK_SPREADSHEET_ID,
        range='Estoque',
        majorDimension='COLUMNS').execute()
    stock = stock_result.get('values', [])
    d = {col[0]: col[1:] for col in stock}    
    stock_df = pd.DataFrame(data=d)
    stock_df["Quantidade necessaria"] = [float(0)] * len(stock_df.index)
    stock_df = stock_df.set_index("Ingrediente")

    #Buscando receitas no diretorio 
    recipes_response = drive_service.files().list(q="'{}' in parents and trashed = False".format(RECIPES_DIR_ID),
                                          spaces='drive',
                                          fields='files(id, name)').execute()
    recipes = {}
    for recipe_file in recipes_response['files']:
        document = docs_service.documents().get(documentId=recipe_file['id']).execute()
        title = document.get('title')
        content = document.get('body').get('content')
        recipe_lines = read_structural_elements(content).splitlines()
        recipe_dict = {}
        for line in recipe_lines:
            if line != "":
                try:
                    recipe_dict[line[:line.index(':')]] = float(line[line.index(':')+1:])
                except Exception as e:
                    print("{}: line: <{}>".format(e, line))
        recipes[title] = recipe_dict
    
    #Calculando quantidade necessaria de cada ingrediente
    for index, row in orders_df.iterrows():
        for col in orders_df.columns:
            if row[col] != "0" and col != "Cliente":
                for ingredient, qty in recipes[col].items():
                    print(qty)
                    print(row[col])
                    stock_df.at[ingredient, "Quantidade necessaria"] += float(qty) * float(row[col])

if __name__ == "__main__":
    main()