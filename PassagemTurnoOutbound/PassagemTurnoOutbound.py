# -------------------------------------------------------------------------------
# Name:        PassagemTurnoOutbound.py
# Purpose:     Relatório de Passagem de Turno
#
# Author:      Antônio Lázaro
#
# Created:     14/01/2021
# Copyright:   (c) Antônio Lázaro
# -------------------------------------------------------------------------------

import configparser
import os
import sqlite3 as sqlite
import sys

import pandas as pd
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import qApp, QMessageBox, QFileDialog

import PandasModelo as pdm
from GoogleSheets import GoogleSheet


def atualizarPlanilhaGoogle(pDataframe):
    gs = GoogleSheet()
    wk = gs.AbrirPlanilha('1f7JeT1mx04yfkOHilJuF9cIEJrA6YbHkL2qEh91pJX0')
    wk.clear()
    linhas = pDataframe.values.tolist()
    for x in linhas:
        wk.insert_row(x)
    wk.insert_rows([pDataframe.columns.values.tolist()])


def lerArquivoConfiguracao():
    cfg = configparser.ConfigParser()
    cfg.read('etd.ini')
    caminho_completo = cfg.get('Database', 'Arquivo')
    return caminho_completo


def lerETD(pData, pHora):
    arquivo = lerArquivoConfiguracao()
    conn = sqlite.connect(arquivo)
    cursor = conn.cursor()
    sql = "select Data, Hora,Planning, Picking, Packing, Shipping FROM etd_resultado e where e.data = '{0}' and " \
          "e.hora = (select " \
          "Max(etd.hora) from etd_resultado etd where etd.data = '{0}' and etd.hora <= '{1}')".format(pData, pHora)
    cursor.execute(sql)
    lista = []
    for linha in cursor.fetchall():
        lista = linha
    conn.close()
    if len(lista) == 0:
        lista = ['', '', 0, 0, 0, 0]
    return lista


def carregarBuffer(area, data):
    listaBuffer = lerETD(data[0:10], data[11:-3])
    if area == 'Picking':
        return listaBuffer[3]
    elif area == 'Packing':
        return listaBuffer[4]
    elif area == 'Shipping':
        return listaBuffer[5]
    elif area == 'Put to wall':
        return 0


def processarMovimento(listaArquivo):
    lista_processos = [
        {"nome": "Picking", "arquivo": listaArquivo[0], "delimitador": ",", "campo_data": "Fecha",
         "campo_quantidade": "Cantidad"},
        {"nome": "Packing", "arquivo": listaArquivo[0], "delimitador": ",", "campo_data": "Fecha",
         "campo_quantidade": "Cantidad"},
        {"nome": "Put to wall", "arquivo": listaArquivo[0], "delimitador": ",", "campo_data": "Fecha",
         "campo_quantidade": "Cantidad"},
        {"nome": "Withdrawal", "arquivo": listaArquivo[0], "delimitador": ",", "campo_data": "Fecha",
         "campo_quantidade": "Cantidad"}]
    listaResultado = []
    for processo in lista_processos:
        if not os.path.isfile(processo["arquivo"]):
            continue

        dfOriginal = pd.read_csv(processo["arquivo"], delimiter="{0}".format(processo["delimitador"]), low_memory=False)

        if processo["nome"] == "Shipping":
            dfOriginal.rename({"Outbound Included Date": "Outbound_Included_Date"}, axis=1, inplace=True)

        campo_data = processo["campo_data"]
        campo_qtd = processo["campo_quantidade"]

        dataMinima = dfOriginal[campo_data].min(axis=0)
        dataMaxima = dfOriginal[campo_data].max(axis=0)
        dataMinima = str(dataMinima)[0:10]
        dataMaxima = str(dataMaxima)[0:10]

        listaTurno = [
            {"turno": "Anterior", "inicio": '{0} 00:00:00'.format(dataMinima), "fim": '{0} 05:19:59'.format(
                dataMinima)},
            {"turno": "Matutino", "inicio": '{0} 05:20:00'.format(dataMinima), "fim": '{0} 13:39:59'.format(
                dataMinima)},
            {"turno": "Vespertino", "inicio": '{0} 13:40:00'.format(dataMinima), "fim": '{0} 21:59:59'.format(
                dataMinima)},
            {"turno": "Noturno", "inicio": '{0} 22:00:00'.format(dataMinima), "fim": '{0} 05:19:59'.format(
                dataMaxima)}]

        total = 0;
        for periodo in listaTurno:
            if processo["nome"] == "Shipping":
                str_query = "(({0} >= '{1}') & ({0} <= '{2}'))".format(campo_data,
                                                                       periodo["inicio"],
                                                                       periodo["fim"])

            else:
                str_query = f"(({'Proceso'} == '{processo['nome']}') & ({campo_data} >= '{periodo['inicio']}') & ({campo_data} <= '{periodo['fim']}'))"

            # qtdBuffer = np.int64(carregarBuffer(processo["nome"], periodo["fim"]))

            dfSomaQtd = dfOriginal.copy()
            dfSomaQtd = dfSomaQtd.query(str_query)
            qtdSomaQtd = 0
            if not dfSomaQtd.empty:
                if processo["nome"] == "Shipping":
                    qtdSomaQtd = dfSomaQtd.shape[0]
                else:
                    qtdSomaQtd = dfSomaQtd[campo_qtd].sum(axis=0)
            listaResultado.append([processo["nome"], periodo["turno"], periodo["inicio"], periodo["fim"],
                                   qtdSomaQtd])
            total += qtdSomaQtd
    return pd.DataFrame(listaResultado,
                        columns=['Processo', 'Turno', 'Início', 'Término', 'Realizado'])


def actionSair():
    resposta = QMessageBox.question(qApp.activeWindow(), 'Pergunta',
                                    "Deseja encerrar a aplicação?",
                                    QMessageBox.Yes | QMessageBox.No,
                                    QMessageBox.No)
    if resposta == QMessageBox.Yes:
        qApp.quit()


def actionCarregarArquivo():
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    qApp.activeWindow().arquivo, _ = QFileDialog.getOpenFileName(qApp.activeWindow(), "Abrir arquivo", "",
                                                                 "Arquivo csv (*.csv)", options=options)
    if qApp.activeWindow().arquivo:
        atualizarInformacao()


def atualizarInformacao():
    qApp.activeWindow().statusBar().showMessage("Aguarde atualizando informações...")
    df = processarMovimento([qApp.activeWindow().arquivo])
    modelo = pdm.pandasModel(df)
    qApp.activeWindow().tableView.setModel(modelo)
    rowCount = modelo.rowCount() - 1
    qApp.activeWindow().tableView.selectRow(rowCount)
    df.to_excel("RelatorioPassagemdeTurno.xlsx")
    atualizarPlanilhaGoogle(df)
    qApp.activeWindow().statusBar().clearMessage()


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('frmPassagemTurnoOutbound.ui', self)
        self.arquivo = ""
        tableview_x = 950
        tableview_y = 950
        # Redimensionamento da tela e do tableview
        self.tableView.resize(tableview_x, tableview_y)
        self.tableView.setMaximumHeight(tableview_x - 100)
        self.tableView.setMinimumHeight(tableview_x - 100)
        self.tableView.setMaximumWidth(tableview_x)
        self.tableView.setMinimumWidth(tableview_x)
        self.resize(tableview_x + 30, tableview_y - 250)
        self.setMaximumHeight(tableview_x - 40)
        self.setMinimumHeight(tableview_x - 40)
        self.setMaximumWidth(tableview_x)
        self.setMinimumWidth(tableview_x)
        # Associação de Eventos
        self.actionSair.triggered.connect(actionSair)
        self.actCarregarArquivo.triggered.connect(actionCarregarArquivo)


def main():
    app = QtWidgets.QApplication(sys.argv)
    FWindow = Ui()
    qApp.setActiveWindow(FWindow)
    FWindow.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
