import gspread
from oauth2client.service_account import ServiceAccountCredentials


class GoogleSheet:
    credencial = 'cad-ba-48568d667654.json'
    def __init__(self):
        # Escopo utilizado
        scope = ['https://spreadsheets.google.com/feeds']

        # Dados de autenticação
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(
            self.credencial, scope)

    def AbrirPlanilha(self, planilha, pagina=0):
        gc = gspread.authorize(self.credentials)
        wks = gc.open_by_key(planilha)
        return wks.get_worksheet(pagina)
