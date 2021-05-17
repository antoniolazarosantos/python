import sys

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QDesktopWidget, QApplication, qApp, QLabel, QCalendarWidget, QLineEdit, \
    QPushButton, QAction, QMessageBox, QFileDialog
from BancoAccess import ConectarDB


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
        # self.initMenu()
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
