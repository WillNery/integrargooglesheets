## passo 1: import das libraries importantes da API pra conexão
from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import requests
from bs4 import BeautifulSoup
import pandas as pd

from datetime import datetime

############## WEB SCRAPPING ###################################

# acesso ao site
url1 = 'https://statusinvest.com.br/acoes/b3sa3'
headers = {'User-Agent' :'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36'}
page = requests.get( url1, headers = headers)
soup1 = BeautifulSoup(page.text, 'html.parser')
indicadores1 = soup1.find_all('strong', class_ =  'value d-block lh-4 fs-4 fw-700')


url2 = 'https://statusinvest.com.br/acoes/abev3'
headers = {'User-Agent' :'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36'}
page = requests.get( url2, headers = headers)
soup2 = BeautifulSoup(page.text, 'html.parser')
indicadores2 = soup2.find_all('strong', class_ =  'value d-block lh-4 fs-4 fw-700')


url3 = 'https://statusinvest.com.br/acoes/azul4'
headers = {'User-Agent' :'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36'}
page = requests.get( url3, headers = headers)
soup3 = BeautifulSoup(page.text, 'html.parser')
indicadores3 = soup3.find_all('strong', class_ =  'value d-block lh-4 fs-4 fw-700')





url4 = 'https://statusinvest.com.br/acoes/bbas3'
headers = {'User-Agent' :'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36'}
page = requests.get( url4, headers = headers)
soup4 = BeautifulSoup(page.text, 'html.parser')
indicadores4 = soup4.find_all('strong', class_ =  'value d-block lh-4 fs-4 fw-700')



url5 = 'https://statusinvest.com.br/acoes/cash3'
headers = {'User-Agent' :'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36'}
page = requests.get( url5, headers = headers)
soup5 = BeautifulSoup(page.text, 'html.parser')
indicadores5 = soup5.find_all('strong', class_ =  'value d-block lh-4 fs-4 fw-700')




url1 = 'https://statusinvest.com.br/acoes/b3sa3'
headers = {'User-Agent' :'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36'}
page = requests.get( url1, headers = headers)
soup1 = BeautifulSoup(page.text, 'html.parser')
indicadores1 = soup1.find_all('strong', class_ =  'value d-block lh-4 fs-4 fw-700')






############## WEB SCRAPPING ###################################

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '131WafSDoKKUz7SMQ9zRAxQmIqnJAZuJkD3rbH5lkh2o'
SAMPLE_RANGE_NAME = 'Summary!AF11:AP38'

# Permitindo a conexão da do Pycharm com a API do google
def main():

    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        #ler informações do google sheets
        # sheet = service.spreadsheets()
        # result = sheet.values().get(spreadsheetId='131WafSDoKKUz7SMQ9zRAxQmIqnJAZuJkD3rbH5lkh2o',
        #                             range='Summary!AF11:AF38').execute()
        # values = result.get('values', [])
        #
        # if not values:
        #     print('No data found.')
        #     return
        #
        # print(values)

        # -- INDICADORES
        ## indicadores[2]: Peg Ratio
        ## indicadores[3]: P/ VP
        ## indicadores[4]: EV/EBITDA
        ## indicadores[8]: VPA
        ## indicadores[16]: DIV LIQ/ EBIT
        ## indicadores[22]: M.EBIT
        ## indicadores[23]: Margem Liquida
        ## indicadores[24]: ROE
        ## indicadores[28]: CAGR RECEITA
        ## indicadores[29]: CAGR LUCRO
        ## indicadores[0]: YIELD ( 12 meses)


        # #adicionar ou editar valores no google sheets
        valores_adicionar1 = [
                            [indicadores1[2].get_text(),
                            indicadores1[3].get_text(),
                            indicadores1[4].get_text(),
                            indicadores1[8].get_text(),
                            indicadores1[16].get_text(),
                            indicadores1[22].get_text(),
                            indicadores1[23].get_text(),
                            indicadores1[24].get_text(),
                            indicadores1[28].get_text(),
                            indicadores1[29].get_text(),
                            indicadores1[0].get_text()]

                            ]

        valores_adicionar2 = [
                            [indicadores2[2].get_text(),
                            indicadores2[3].get_text(),
                            indicadores2[4].get_text(),
                            indicadores2[8].get_text(),
                            indicadores2[16].get_text(),
                            indicadores2[22].get_text(),
                            indicadores2[23].get_text(),
                            indicadores2[24].get_text(),
                            indicadores2[28].get_text(),
                            indicadores2[29].get_text(),
                            indicadores2[0].get_text()]

                            ]

        valores_adicionar3 = [
                            [indicadores3[2].get_text(),
                            indicadores3[3].get_text(),
                            indicadores3[4].get_text(),
                            indicadores3[8].get_text(),
                            indicadores3[16].get_text(),
                            indicadores3[22].get_text(),
                            indicadores3[23].get_text(),
                            indicadores3[24].get_text(),
                            indicadores3[28].get_text(),
                            indicadores3[29].get_text(),
                            indicadores3[0].get_text()]
                            ]


        valores_adicionar4 = [
                            [indicadores4[2].get_text(),
                            indicadores4[3].get_text(),
                            indicadores4[4].get_text(),
                            indicadores4[8].get_text(),
                            indicadores4[16].get_text(),
                            indicadores4[22].get_text(),
                            indicadores4[23].get_text(),
                            indicadores4[24].get_text(),
                            indicadores4[28].get_text(),
                            indicadores4[29].get_text(),
                            indicadores4[0].get_text()]

                            ]

        valores_adicionar5 = [
                            [indicadores5[2].get_text(),
                            indicadores5[3].get_text(),
                            indicadores5[4].get_text(),
                            indicadores5[8].get_text(),
                            indicadores5[16].get_text(),
                            indicadores5[22].get_text(),
                            indicadores5[23].get_text(),
                            indicadores5[24].get_text(),
                            indicadores5[28].get_text(),
                            indicadores5[29].get_text(),
                            indicadores5[0].get_text()]

                            ]

        #Ação 1
        sheet = service.spreadsheets()
        result = sheet.values().update(spreadsheetId='131WafSDoKKUz7SMQ9zRAxQmIqnJAZuJkD3rbH5lkh2o',
                                       range='Summary!AF14', valueInputOption = 'USER_ENTERED',
                                       body = {'values': valores_adicionar1}).execute()
        values1 = result.get('values', [])

        #Ação 2
        sheet = service.spreadsheets()
        result = sheet.values().update(spreadsheetId='131WafSDoKKUz7SMQ9zRAxQmIqnJAZuJkD3rbH5lkh2o',
                                       range='Summary!AF15', valueInputOption = 'USER_ENTERED',
                                       body = {'values': valores_adicionar2}).execute()
        values2 = result.get('values', [])

        #Ação 3
        sheet = service.spreadsheets()
        result = sheet.values().update(spreadsheetId='131WafSDoKKUz7SMQ9zRAxQmIqnJAZuJkD3rbH5lkh2o',
                                       range='Summary!AF16', valueInputOption = 'USER_ENTERED',
                                       body = {'values': valores_adicionar3}).execute()
        values3 = result.get('values', [])

        #Ação 4
        sheet = service.spreadsheets()
        result = sheet.values().update(spreadsheetId='131WafSDoKKUz7SMQ9zRAxQmIqnJAZuJkD3rbH5lkh2o',
                                       range='Summary!AF17', valueInputOption = 'USER_ENTERED',
                                       body = {'values': valores_adicionar4}).execute()
        values4 = result.get('values', [])

        #Ação 5
        sheet = service.spreadsheets()
        result = sheet.values().update(spreadsheetId='131WafSDoKKUz7SMQ9zRAxQmIqnJAZuJkD3rbH5lkh2o',
                                       range='Summary!AF18', valueInputOption = 'USER_ENTERED',
                                       body = {'values': valores_adicionar5}).execute()
        values5 = result.get('values', [])

        print(values1)
        print(values2)
        print(values3)
        print(values4)
        print(values5)






    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()