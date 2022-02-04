# Для загрузок
from yaml import load

# Стартовые библиотеки для выгрузки/загрузки отчета
import httplib2
import argparse
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# Работа с папками
import os


class Googlesheets:

    #example: gs = Googlesheets('start!a1:z1000', [[21, 12323], [23, 345]], 'imvo')

    def __init__(self, rangeName, dataSheet, clientName):

        self.rangeName = rangeName
        self.dataSheet = dataSheet
        self.clientName = clientName

        self.credentials = None
        self.spreadsheetId = None
        self.service = None

        self.authorization()
        self.requestMethod()

        self.body = {'values': self.dataSheet}

        self.listCreate()
        self.dataUpload()

    def authorization(self):

        # Разрешения на просмотр и редактирование
        SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
        # Название файла с идентификаторами
        CLIENT_SECRET_FILE = 'client_secret_grants.json'
        # Название приложения для онлайн-отчета (можно использовать любое)
        APPLICATION_NAME = 'Google Sheets API Report'
        # Токены хранятся в 'analytics.dat'
        credential_path = 'sheets.googleapis.com-report.json'

        # Процесс авторизации
        store = Storage(credential_path)
        credentials = store.get()
        parser = \
            argparse.ArgumentParser(
                formatter_class=argparse.RawDescriptionHelpFormatter, parents=[tools.argparser])
        flags = parser.parse_args([])
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            print('Storing credentials to ' + credential_path)
        self.credentials = credentials

    def requestMethod(self):

        # Формирование запроса
        http = self.credentials.authorize(httplib2.Http())
        discoveryUrl = (
            'https://sheets.googleapis.com/$discovery/rest?version=v4')
        service = discovery.build(
            'sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)
        f = open(os.path.join('clients/'+self.clientName, 'config.yaml'), 'r')
        sheet_config = load(f)

        # Вызов идентификатора файла
        self.spreadsheetId = str(sheet_config['sheet_id'])
        self.service = service

    def dataUpload(self):

        # Обращаемся к айдишнику документа
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId, range=self.rangeName).execute()

        # Отправляем запрос
        self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheetId,
                                                    range=self.rangeName,
                                                    valueInputOption='USER_ENTERED',
                                                    body=self.body).execute()
        print('Выгрузка завершена успешно')

    def listCreate(self):

        try:
            # Добавление листа
            results = self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheetId,
                body={"requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": self.rangeName.split('!')[0],
                            }}}]}).execute()
        except:
            pass
