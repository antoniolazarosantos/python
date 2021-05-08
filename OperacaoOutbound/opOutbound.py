import os
import sys

from PyQt5.QtWidgets import QMainWindow, QApplication, QAction, QDesktopWidget, QFileDialog, \
    QTableView, QMessageBox, qApp, QComboBox, QLabel
from PyQt5.QtGui import QIcon
import PandasModelo as pdm
import pandas as pd
from datetime import timedelta
from dateutil.parser import parse


class App(QMainWindow):

    def ProcessarDados(self, pArquivo, pProcesso):
        df = pd.read_csv(pArquivo, delimiter=",")
        listaProcesso = []
        if pProcesso != "Todos":
            listaProcesso.append(pProcesso)
        else:
            listaProcesso.append("Picking")
            listaProcesso.append("Packing")
            listaProcesso.append("Put to wall")
            listaProcesso.append("Withdrawal")

        dataInicial = df["Fecha"].min(axis=0)
        dataFinal = df["Fecha"].max(axis=0)
        d2 = pd.to_datetime(dataFinal).strftime("%m/%d/%Y")
        d1 = pd.to_datetime(dataInicial).strftime("%m/%d/%Y")
        quantidade_dias = abs((parse(d2, dayfirst=True) - parse(d1, dayfirst=True)).days) + 1
        data = parse(pd.to_datetime(dataInicial).strftime("%m/%d/%Y"), dayfirst=True)
        lista_data_mov = []
        for x in range(quantidade_dias):
            lista_data_mov.append(pd.to_datetime(data).strftime("%d/%m/%Y"))
            dt = parse(pd.to_datetime(data).strftime("%d/%m/%Y"), dayfirst=True)
            data = dt + timedelta(days=1)

        listaHorarios = []
        for x in range(24):
            r = {"hora": "{0}h".format(x),
                 "inicio": "{:0>2}:00:00".format(x),
                 "termino": "{:0>2}:59:59".format(x)}
            listaHorarios.append(r)
        lista_data = []
        lista_hora = []
        lista_processo = []
        lista_realizado = []
        lista_total_dia = []
        lista_realizado_total_processo = []
        lista_unidade_hora = []
        lista_hc = []
        lista_produtividade = []
        lista_total_processo = []
        for data in lista_data_mov:
            str_query_total_dia = "((Fecha >= '{0} {1}') & " \
                                  "(Fecha <= '{0} {2}'))".format(data,
                                                                 '00:00:00',
                                                                 '23:59:59')
            dfTotalDia = df.query(str_query_total_dia).copy()
            qtdTotalDia = dfTotalDia["Cantidad"].sum(axis=0)

            for nomeProcesso in listaProcesso:
                str_query_total_processo = "((Proceso == '{0}') & (Fecha >= '{1} {2}') & " \
                                           "(Fecha <= '{1} {3}'))".format(nomeProcesso,
                                                                          data,
                                                                          '00:00:00',
                                                                          '23:59:59')
                dfTotalProcesso = df.query(str_query_total_processo).copy()
                qtdTotalProcesso = dfTotalProcesso["Cantidad"].sum(axis=0)
                for hora in listaHorarios:
                    str_query = "((Proceso == '{0}') & (Fecha >= '{1} {2}') & (Fecha <= '{1} {3}'))".format(
                        nomeProcesso,
                        data,
                        hora["inicio"],
                        hora["termino"])
                    dfResultado = df.query(str_query).copy()
                    qtdTotalRealizado = dfResultado["Cantidad"].sum(axis=0)
                    qtdHC = dfResultado["Responsable"].nunique()
                    lista_hc.append(qtdHC)
                    lista_data.append(data)
                    lista_hora.append(hora["hora"])
                    lista_processo.append(nomeProcesso)
                    lista_realizado.append(qtdTotalRealizado)
                    lista_total_processo.append(qtdTotalProcesso)
                    lista_total_dia.append(qtdTotalDia)
                    percentual = 0
                    if qtdTotalProcesso > 0:
                        percentual = round((qtdTotalRealizado / qtdTotalProcesso) * 100, 2)
                    lista_realizado_total_processo.append(percentual)
                    unidade_hora = round(qtdTotalProcesso / 24, 2)
                    lista_unidade_hora.append(unidade_hora)
                    produtividade = 0
                    if qtdHC > 0:
                        produtividade = round(qtdTotalRealizado / qtdHC, 2)
                    lista_produtividade.append(produtividade)

        dados = {'Processo': lista_processo,
                 'Data': lista_data,
                 'Hora': lista_hora,
                 'Realizado( A )': lista_realizado,
                 'Total do Processo ( B )': lista_total_processo,
                 'Total do Dia( C )': lista_total_dia,
                 '( A x B )%': lista_realizado_total_processo,
                 'Média Und./Hora ( B / 24h )': lista_unidade_hora,
                 'HC Utilizado ( D )': lista_hc,
                 'HC Prod./Hora ( A / D )': lista_produtividade}

        return pd.DataFrame(dados)

    def criarTabela(self, pArquivo, pProcesso):
        tv = QTableView(self)
        self.df = self.ProcessarDados(pArquivo, pProcesso)
        modelo = pdm.pandasModel(self.df)
        tv.setModel(modelo)
        tv.left = 5
        tv.top = 105  # 35
        tv.width = 1550
        tv.height = 850
        tv.setGeometry(tv.left, tv.top, tv.width, tv.height)
        tv.setMaximumHeight(tv.height)
        tv.setMinimumHeight(tv.height)
        tv.setMaximumWidth(tv.width)
        tv.setMinimumWidth(tv.width)
        tv.setShowGrid(True)
        vh = tv.verticalHeader()
        vh.setVisible(True)
        hh = tv.horizontalHeader()
        hh.setStretchLastSection(True)
        tv.resizeColumnsToContents()
        tv.resizeRowsToContents()
        tv.setSortingEnabled(True)
        return tv

    def atualizarTabela(self):
        self.tabela = self.criarTabela(self.fileName, self.cboProcesso.currentData())
        self.tabela.setVisible(True)

    def abrirArquivo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self.fileName, _ = QFileDialog.getOpenFileName(self, "Abrir arquivo", "",
                                                       "Arquivo csv (*.csv)", options=options)
        if self.fileName:
            self.habilitarComboProcesso(True)
            self.tabela = self.criarTabela(self.fileName, "Todos")
            self.tabela.setVisible(True)

    def salvarResultado(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "Salvar arquivo", "",
                                                  "Excel (*.xlsx)", options=options)
        if fileName:
            arquivo = fileName + ".xlsx"
            self.df.to_excel(arquivo)
            if os.path.isfile(arquivo):
                QMessageBox.information(self, "Informação", "Arquivo gravado com sucesso.")

    def fecharAplicacao(self):
        resposta = QMessageBox.question(qApp.activeWindow(), 'Pergunta',
                                        "Deseja encerrar a aplicação?",
                                        QMessageBox.Yes | QMessageBox.No,
                                        QMessageBox.No)
        if resposta == QMessageBox.Yes:
            self.close()

    def verGrafico(self):
        pass

    def __init__(self):
        super().__init__()
        self.title = 'Operação Outbound - 1.0'
        self.left = 10
        self.top = 10
        self.width = 1550
        self.height = 850
        self.setMaximumHeight(self.height)
        self.setMinimumHeight(self.height)
        self.setMaximumWidth(self.width)
        self.setMinimumWidth(self.width)

        self.initUI()

    def centralizarJanela(self):
        qtRectangle = self.frameGeometry()
        centerPoint = QDesktopWidget().availableGeometry().center()
        qtRectangle.moveCenter(centerPoint)
        return qtRectangle.topLeft()

    def initMenu(self):
        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('Arquivo')

        loadFileButton = QAction(QIcon('images/excel.png'), 'Carregar Arquivo', self)
        loadFileButton.setShortcut('Ctrl+A')
        loadFileButton.setStatusTip('Carregar arquivo')
        loadFileButton.triggered.connect(self.abrirArquivo)
        fileMenu.addAction(loadFileButton)

        loadFileButton = QAction(QIcon('images/grafico.png'), 'Visualizar gráfico', self)
        loadFileButton.setShortcut('Ctrl+G')
        loadFileButton.setStatusTip('Visualizar gráfico')
        loadFileButton.triggered.connect(self.verGrafico)
        fileMenu.addAction(loadFileButton)

        saveFileButton = QAction(QIcon('images/save.png'), 'Gravar Resultado', self)
        saveFileButton.setShortcut('Ctrl+S')
        saveFileButton.setStatusTip('Gravar resultado')
        saveFileButton.triggered.connect(self.salvarResultado)
        fileMenu.addAction(saveFileButton)

        exitButton = QAction(QIcon('images/exit.png'), 'Sair', self)
        exitButton.setShortcut('Ctrl+Q')
        exitButton.setStatusTip('Encerrar aplicação')
        exitButton.triggered.connect(self.fecharAplicacao)
        fileMenu.addAction(exitButton)

    def selecionarProcesso(self):
        self.atualizarTabela()

    def habilitarComboProcesso(self, pValor):
        self.lblProcesso.setEnabled(pValor)
        self.cboProcesso.setEnabled(pValor)

    def criarComboProcesso(self):
        self.lblProcesso = QLabel(self)
        self.lblProcesso.setText("Processo")
        self.lblProcesso.setGeometry(10, 35, 120, 30)

        self.cboProcesso = QComboBox(self)
        self.cboProcesso.setGeometry(10, 65, 120, 30)
        lista_opcoes = ["Todos", "Picking", "Packing", "Put to wall", "Withdrawal"]

        self.cboProcesso.addItems(lista_opcoes)

        self.cboProcesso.setItemData(0, "Todos")
        self.cboProcesso.setItemData(1, "Picking")
        self.cboProcesso.setItemData(2, "Packing")
        self.cboProcesso.setItemData(3, "Put to wall")
        self.cboProcesso.setItemData(4, "Withdrawal")

        self.cboProcesso.activated.connect(self.selecionarProcesso)

        self.habilitarComboProcesso(False)

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.move(self.centralizarJanela())

        self.initMenu()
        self.criarComboProcesso()
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = App()
    qApp.setActiveWindow(window)
    sys.exit(app.exec_())
