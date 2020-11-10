import pickle
import os.path
import os
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from dotenv import load_dotenv
from google_apis_util import read_structural_elements, create_drive_file
from math import ceil

#Carregando variáveis de ambiente
load_dotenv()
ORDERS_SPREADSHEET_ID = os.environ.get("ORDERS_SPREADSHEET_ID")
STOCK_SPREADSHEET_ID = os.environ.get("STOCK_SPREADSHEET_ID")
RECIPES_DIR_ID = os.environ.get("RECIPES_DIR_ID")
TEMPLATE_DOC_ID = os.environ.get("TEMPLATE_DOC_ID")
MAIN_FOLDER_ID = os.environ.get("MAIN_FOLDER_ID")

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

    #Obtendo informações de preço
    sheet = sheets_service.spreadsheets()
    prices_result = sheet.values().get(
        spreadsheetId=ORDERS_SPREADSHEET_ID,
        range='Menu').execute()
    prices = dict(prices_result.get('values', [])[1:])

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

    #Obtendo template, contendo nome do produto e política de descontos
    document = docs_service.documents().get(documentId=TEMPLATE_DOC_ID).execute()
    content = document.get('body').get('content')
    template = read_structural_elements(content)
    template_lines = template.splitlines(True)
    product_info = template_lines[0].split(',')
    discount_treshold = int(product_info[2][product_info[2].index('>')+1:-1])
    discount_mult = float(product_info[2][:product_info[2].index('>')])

    #Analisando pedidos
    bills = open("bills.txt", "w+", encoding="utf-8")
    for index, row in orders_df.iterrows():
        template = read_structural_elements(content)
        template_lines = template.splitlines(True)
        current_orders = []
        total_price = 0
        orders_qty = 0
        for col in orders_df.columns:
            if row[col] != "0" and col != "Cliente":
                #Calculo da quantidade necessaria de cada ingrediente
                for ingredient, qty in recipes[col].items():
                    stock_df.at[ingredient, "Quantidade necessaria"] += float(qty) * float(row[col])

                #Construindo mensagem de confirmacao de pedido
                orders_qty += int(row[col])
                total_price += float(row[col])*float(prices[col].replace(',','.'))
                name = product_info[0]
                if int(row[col]) > 1:
                    name = product_info[1]
                current_orders.append("{} {} de {}".format(int(row[col]), name, col))
        
        #Inserindo pedidos no template de mensagem
        insertion_index = template_lines.index('Vamos conferir seu pedido:\n') + 1
        template_lines[insertion_index:insertion_index] = current_orders

        #Calculo de preço
        if orders_qty > discount_treshold:
            total_price *= discount_mult
        total_price = format(total_price, '.2f')
    
        #Adicionando informações de cliente e preço na mensagem
        for index, line in enumerate(template_lines):
            if "{total}" in line:
                template_lines[index] = line.format(total=total_price)
            elif "{client}" in line:
                template_lines[index] = line.format(client=row["Cliente"])
        
        bills.writelines(template_lines[1:])
    
    bills.close()

    with open("shopping_list.txt", "w+") as shopping_list:
        for index, row in stock_df.iterrows():
            """Quantidade por embalagem 0 = ingrediente medido em unidades"""
            buy_qty = 0
            if int(row["Quantidade por Embalagem"]) == 0:
                if float(row["Quantidade necessaria"]) >= float(row["Gramas"]):
                    buy_qty = ceil((float(row["Quantidade necessaria"]) - float(row["Gramas"])))
            elif (float(row["Quantidade necessaria"]) >= float(row["Gramas"]) or float(row["Gramas"]) <= 11):
                buy_qty = ceil((float(row["Quantidade necessaria"]) - float(row["Gramas"])) / float(row["Quantidade por Embalagem"]))
            if buy_qty != 0:
                shopping_list.write("{} x{}\n".format(index, buy_qty))

    #Chamando a Drive API para atualizar as informações
    needed_files = {
        'shopping_list.txt': '', 
        'bills.txt': ''
        }

    for path in needed_files.keys():
        response = drive_service.files().list(q="name='{}' and '{}' in parents and trashed = False".format(path, MAIN_FOLDER_ID),
                                              spaces='drive',
                                              fields='files(id)').execute()
        if len(response['files']) == 0:
            print("Criando o arquivo {} no drive.".format(path))
            response = create_drive_file(path, "./" + path, MAIN_FOLDER_ID, drive_service)
            needed_files[path] = response['id']
        else:
            needed_files[path] = response['files'][0]['id']

        media = MediaFileUpload("./" + path)
        file = drive_service.files().update(
                                        media_body=media,
                                        fileId=needed_files[path],
                                        fields='id').execute()

    print("Informações salvas no drive com sucesso.")
if __name__ == "__main__":
    main()