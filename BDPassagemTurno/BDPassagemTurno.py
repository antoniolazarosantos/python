import sys
from datetime import date
import pyodbc
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QDesktopWidget, QApplication, qApp, QLabel, QCalendarWidget, QLineEdit, \
    QPushButton, QAction, QMessageBox, QFileDialog


class ConectarDB:
    def __init__(self):
        # Criando conexão..
        self.con = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                                  r'DBQ=C:\Mercado Livre\Áreas\Waves\Passagem de Turno\BDPassagemTurnoOutbound.accdb;')
        # Criando cursor.
        self.cur = self.con.cursor()

    def exibir_tabelas(self):
        for tabela in self.cur.tables(tableType='TABLE'):
            print(tabela.table_name)

    def consultar_registros(self):
        return self.cur.execute('SELECT * FROM Calendario order by 1').fetchall()

    def inserir_script_diario(self, pLista):
        # Calendário
        # self.cur.execute("select max(código) from calendario")
        # row = self.cur.fetchone()
        # id = 0
        # while row:
        #    id = row[0]
        #    row = self.cur.fetchone()

        print("Limpando dados da data...")
        lista = ["calendario", "sla", "produtividade", "[Unidades em Buffer]",
                 "[Unidades Processadas]", "waves"]
        lData = pLista[0].split("/")
        xData = lData[1] + '/' + lData[0] + '/' + lData[2]
        for t in lista:
            self.cur.execute("delete from {0} where data = #{1}#".format(t, xData))
        self.cur.commit()

        print("Iniciando calendário...")
        self.cur.execute("insert into calendario (código,data,target) "
                         "select max(código)+1,'{0}',{1} from calendario".format(pLista[0], pLista[1]))
        # SLA
        print("Iniciando SLA...")
        self.cur.execute("insert into sla (id,data) "
                         "select max(id)+1,'{0}' from sla".format(pLista[0]))
        # Produtividade
        print("Iniciando produtividade...")
        self.cur.execute("insert into produtividade (id,data) "
                         "select max(id)+1,'{0}' from produtividade".format(pLista[0]))
        # Unidades em Buffer
        print("Iniciando Unidades em Buffer...")
        lista = ["Planning", "Picking", "Packing", "Put to Wall", "Shipping"]
        for processo in lista:
            self.cur.execute("insert into [Unidades em Buffer] (processo,data,turno,buffer) "
                             "values ('{0}','{1}','Noturno',0)".format(processo, pLista[0]))
        # Unidades Processadas
        print("Iniciando Unidades Processadas...")
        lista = ["Picking", "Wall", "Packing", "Shipping"]
        for processo in lista:
            self.cur.execute("insert into [Unidades Processadas] (processo,data,turno,realizado) "
                             "values ('{0}','{1}','Noturno',0)".format(processo, pLista[0]))

        # Waves
        print("Iniciando Waves...")
        listaHora = ["22:00", "23:00", "00:00", "01:00", "02:00", "03:00", "04:00"]
        for hora in listaHora:
            self.cur.execute("insert into waves (processo,data,turno,intervalo,"
                             "mono,multi,emergência,retiro) "
                             "values ('Wave','{0}','Noturno','{1}',1,1,0,0)".format(pLista[0], hora))

        self.cur.commit()
        print("Processo finalizado com sucesso.")


class App(QMainWindow):

    def __init__(self):
        super().__init__()
        self.title = 'Script Inicial - 1.0'
        self.left = 10
        self.top = 10
        self.width = 450
        self.height = 450
        self.setMaximumHeight(self.height)
        self.setMinimumHeight(self.height)
        self.setMaximumWidth(self.width)
        self.setMinimumWidth(self.width)

        self.initUI()

    def abrirArquivo(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self.fileName, _ = QFileDialog.getOpenFileName(self, "Carregar arquivo", "",
                                                       "Access (*.mdb, *.accdb)", options=options)
        if self.fileName:
            pass

    def fecharAplicacao(self):
        resposta = QMessageBox.question(qApp.activeWindow(), 'Pergunta',
                                        "Deseja encerrar a aplicação?",
                                        QMessageBox.Yes | QMessageBox.No,
                                        QMessageBox.No)
        if resposta == QMessageBox.Yes:
            self.close()

    def initMenu(self):
        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('Arquivo')

        loadFileButton = QAction(QIcon('images/excel.png'), 'Carregar Arquivo', self)
        loadFileButton.setShortcut('Ctrl+A')
        loadFileButton.setStatusTip('Carregar arquivo')
        loadFileButton.triggered.connect(self.abrirArquivo)
        fileMenu.addAction(loadFileButton)

        exitButton = QAction(QIcon('images/exit.png'), 'Sair', self)
        exitButton.setShortcut('Ctrl+Q')
        exitButton.setStatusTip('Encerrar aplicação')
        exitButton.triggered.connect(self.fecharAplicacao)
        fileMenu.addAction(exitButton)

    def centralizarJanela(self):
        qtRectangle = self.frameGeometry()
        centerPoint = QDesktopWidget().availableGeometry().center()
        qtRectangle.moveCenter(centerPoint)
        return qtRectangle.topLeft()

    def criarCalendario(self):
        self.lblCalendario = QLabel(self)
        self.lblCalendario.setText("Selecione a data de processamento:")
        self.lblCalendario.setGeometry(5, 10, 350, 30)
        self.calendario = QCalendarWidget(self)
        self.calendario.setGeometry(5, 40, 400, 250)

    def criarInputTarget(self):
        self.lblTarget = QLabel(self)
        self.lblTarget.setText("Target:")
        self.lblTarget.setGeometry(5, 310, 500, 30)
        self.target = QLineEdit(self)
        self.target.setGeometry(65, 310, 150, 30)

    def on_click_processar(self):
        self.statusBar().showMessage("Aguarde processando script...")
        data = self.calendario.selectedDate().toPyDate().strftime('%d/%m/%Y')
        lista = [data,
                 self.target.text()]
        print(lista)
        self.banco.inserir_script_diario(lista)
        self.statusBar().showMessage("Script processado com sucesso...")

    def criarBtnProcessar(self):
        self.btnProcessar = QPushButton('Processar', self)
        self.btnProcessar.setToolTip('Processar script de inicialização')
        self.btnProcessar.move(310, 310)
        self.btnProcessar.clicked.connect(self.on_click_processar)

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.move(self.centralizarJanela())
        #self.initMenu()
        self.criarCalendario()
        self.criarInputTarget()
        self.criarBtnProcessar()
        self.banco = ConectarDB()
        self.show()


if __name__ == '__main__':
    # banco.exibir_tabelas()
    # data = str(input("Data de Processamento: "))
    # target = int(input("Target: "))
    # banco.inserir_script_diario([data, target])
    # print(banco.consultar_registros())
    app = QApplication(sys.argv)
    window = App()
    qApp.setActiveWindow(window)
    sys.exit(app.exec_())
