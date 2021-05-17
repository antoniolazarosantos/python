import os
import sys

from PyQt5.QtWidgets import QMainWindow, QApplication, QAction, QDesktopWidget, QFileDialog, \
    QTableView, QMessageBox, qApp, QComboBox, QLabel, QDialog
from PyQt5.QtGui import QIcon
import PandasModelo as pdm
import pandas as pd
from datetime import timedelta
from dateutil.parser import parse

class AppGrafico(QDialog):
    def __init__(self, parent=None):
        super(AppGrafico, self).__init__(parent)

class App(QMainWindow):

    def processarDados(self, pArquivo, pRepositor):
        df = pd.read_csv(pArquivo, delimiter=",")

        self.cboRepositor.clear()
        self.cboRepositor.addItem("Todos")
        self.cboRepositor.addItems(df['Responsable'].unique())

        listaRepositor = []

        if pRepositor != "Todos":
            self.cboRepositor.setCurrentText(pRepositor)
            listaRepositor.append(pRepositor)
        else:
            listaRepositor = df['Responsable'].unique()

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
        lista_repositor = []
        lista_realizado = []
        lista_realizado_total_repositor = []
        lista_unidade_minuto = []
        lista_total_repositor = []
        for data in lista_data_mov:
            for nomeRepositor in listaRepositor:
                str_query_total_repositor = "((Responsable == '{0}') & (Fecha >= '{1} {2}') & " \
                                           "(Fecha <= '{1} {3}'))".format(nomeRepositor,
                                                                          data,
                                                                          '00:00:00',
                                                                          '23:59:59')
                dfTotalRepositor = df.query(str_query_total_repositor).copy()
                qtdTotalRepositor = dfTotalRepositor["Cantidad"].sum(axis=0)
                for hora in listaHorarios:
                    str_query = "((Responsable == '{0}') & (Fecha >= '{1} {2}') & (Fecha <= '{1} {3}'))".format(
                        nomeRepositor,
                        data,
                        hora["inicio"],
                        hora["termino"])
                    dfResultado = df.query(str_query).copy()
                    qtdTotalRealizado = dfResultado["Cantidad"].sum(axis=0)
                    lista_data.append(data)
                    lista_hora.append(hora["hora"])
                    lista_repositor.append(nomeRepositor)
                    lista_realizado.append(qtdTotalRealizado)
                    lista_total_repositor.append(qtdTotalRepositor)
                    percentual = 0
                    if qtdTotalRepositor > 0:
                        percentual = round((qtdTotalRealizado / qtdTotalRepositor) * 100, 2)
                    lista_realizado_total_repositor.append(percentual)
                    unidade_minuto = round(qtdTotalRealizado / 60, 2)
                    lista_unidade_minuto.append(unidade_minuto)

        dados = {'Repositor': lista_repositor,
                 'Data': lista_data,
                 'Hora': lista_hora,
                 'Realizado( A )': lista_realizado,
                 'Total do Realizado no dia ( B )': lista_total_repositor,
                 '( A x B )%': lista_realizado_total_repositor,
                 'Média Und./Minuto ( A / 60 min. )': lista_unidade_minuto}

        return pd.DataFrame(dados)

    def criarTabela(self, pArquivo, pRepositor):
        tv = QTableView(self)
        self.df = self.processarDados(pArquivo, pRepositor)
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
        self.tabela = self.criarTabela(self.fileName, self.cboRepositor.currentText())
        self.tabela.setVisible(True)

    def abrirArquivo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self.fileName, _ = QFileDialog.getOpenFileName(self, "Abrir arquivo", "",
                                                       "Arquivo csv (*.csv)", options=options)
        if self.fileName:
            self.habilitarComboRepositores(True)
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
        t = AppGrafico(self)
        t.show()
        t.exec_()


    def __init__(self):
        super().__init__()
        self.title = 'Movimentação Realizada - Repositores - 1.0'
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

    def selecionarRepositor(self):
        self.atualizarTabela()

    def habilitarComboRepositores(self, pValor):
        self.lblRepositor.setEnabled(pValor)
        self.cboRepositor.setEnabled(pValor)

    def criarComboRepositores(self):
        self.lblRepositor = QLabel(self)
        self.lblRepositor.setText("Repositor")
        self.lblRepositor.setGeometry(10, 35, 120, 30)

        self.cboRepositor = QComboBox(self)
        self.cboRepositor.setGeometry(10, 65, 450, 30)
        lista_opcoes = ["Todos"]

        self.cboRepositor.addItems(lista_opcoes)

        self.cboRepositor.setItemData(0, "Todos")

        self.cboRepositor.activated.connect(self.selecionarRepositor)

        self.habilitarComboRepositores(False)

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.move(self.centralizarJanela())

        self.initMenu()
        self.criarComboRepositores()
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = App()
    qApp.setActiveWindow(window)
    sys.exit(app.exec_())
