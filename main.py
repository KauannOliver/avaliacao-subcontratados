import sys
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QLabel,
                             QLineEdit, QPushButton, QListWidget, QTabWidget,
                             QDateEdit, QMessageBox, QComboBox, QFileDialog)
import functions  ### importa as funções do arquivo functions.py ###
import pandas as pd
from openpyxl import load_workbook

### classe principal da janela do aplicativo ###
class JanelaPrincipal(QMainWindow):
    def __init__(self):
        super(JanelaPrincipal, self).__init__()

        ### define o título e o tamanho da janela principal ###
        self.setWindowTitle("Avaliação Subcontratados")
        self.setGeometry(100, 100, 800, 600)

        ### seleciona o arquivo do Excel a ser utilizado ###
        self.caminhoArquivo = self.selecionarArquivo()
        if not self.caminhoArquivo:
            QMessageBox.warning(self, "Erro", "Nenhum arquivo selecionado. O aplicativo será fechado.")
            sys.exit()

        ### carrega os dados necessários das planilhas ###
        self.motoristas = functions.carregarNomes('Motorista x Fornecedor', self.caminhoArquivo)
        self.fornecedores = functions.carregarNomes('Fornecedor', self.caminhoArquivo)
        self.criteriosMotorista = functions.carregarCriterios('Critério Motorista', self.caminhoArquivo)
        self.criteriosFornecedor = functions.carregarCriterios('Critério Fornecedor', self.caminhoArquivo)

        ### inicializa a interface do usuário ###
        self.inicializarInterface()

    ### função para selecionar o arquivo do Excel ###
    def selecionarArquivo(self):
        opcoes = QFileDialog.Options()
        opcoes |= QFileDialog.ReadOnly
        caminhoArquivo, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha", "",
                                                        "Excel Files (*.xlsx);;All Files (*)", options=opcoes)
        return caminhoArquivo

    ### função para inicializar a interface do usuário ###
    def inicializarInterface(self):
        self.abaMotorista = QtWidgets.QWidget()
        self.abaWidget = QTabWidget()
        self.setCentralWidget(self.abaWidget)
        self.abaWidget.addTab(self.abaMotorista, "Lançamento Motorista")

        self.layoutMotorista = QVBoxLayout()
        self.layoutFormularioMotorista = QVBoxLayout()

        ### define o rótulo e o campo de texto para o nome do motorista ###
        self.labelMotorista = QLabel("Motorista:")
        self.layoutFormularioMotorista.addWidget(self.labelMotorista)
        self.campoMotorista = QLineEdit()
        self.layoutFormularioMotorista.addWidget(self.campoMotorista)
        self.listaMotorista = QListWidget()
        self.layoutFormularioMotorista.addWidget(self.listaMotorista)

        ### conecta a entrada de texto à função de atualização da lista ###
        self.campoMotorista.textChanged.connect(
            lambda: self.atualizarLista(self.motoristas, self.campoMotorista, self.listaMotorista))
        self.listaMotorista.itemClicked.connect(
            lambda item: self.selecionarItemLista(item, self.campoMotorista, self.listaMotorista))

        ### define o rótulo e o dropdown para o critério do motorista ###
        self.labelCriterioMotorista = QLabel("Critério:")
        self.layoutFormularioMotorista.addWidget(self.labelCriterioMotorista)
        self.dropdownCriterioMotorista = QComboBox()
        self.dropdownCriterioMotorista.addItems(self.criteriosMotorista.keys())
        self.layoutFormularioMotorista.addWidget(self.dropdownCriterioMotorista)

        ### conecta a mudança de critério à função de atualização dos pontos ###
        self.dropdownCriterioMotorista.currentIndexChanged.connect(self.atualizarPontosMotorista)

        ### define o rótulo e o campo de texto para os pontos do motorista ###
        self.labelPontosMotorista = QLabel("Pontos:")
        self.layoutFormularioMotorista.addWidget(self.labelPontosMotorista)
        self.campoPontosMotorista = QLineEdit()
        self.campoPontosMotorista.setReadOnly(True)
        self.layoutFormularioMotorista.addWidget(self.campoPontosMotorista)

        ### define o rótulo e o campo de data para a data de lançamento ###
        self.labelDataMotorista = QLabel("Data:")
        self.layoutFormularioMotorista.addWidget(self.labelDataMotorista)
        self.dataMotorista = QDateEdit()
        self.dataMotorista.setDisplayFormat('dd/MM/yyyy')
        self.layoutFormularioMotorista.addWidget(self.dataMotorista)

        ### botão para submeter os dados do motorista ###
        self.botaoSubmeterMotorista = QPushButton("Lançar Motorista")
        self.botaoSubmeterMotorista.setFixedHeight(40)
        self.botaoSubmeterMotorista.clicked.connect(self.submeterMotorista)
        self.layoutFormularioMotorista.addWidget(self.botaoSubmeterMotorista)

        self.layoutMotorista.addLayout(self.layoutFormularioMotorista)
        self.abaMotorista.setLayout(self.layoutMotorista)

        ### configuração da aba para fornecedores ###
        self.abaFornecedor = QtWidgets.QWidget()
        self.abaWidget.addTab(self.abaFornecedor, "Lançamento Fornecedor")

        self.layoutFornecedor = QVBoxLayout()
        self.layoutFormularioFornecedor = QVBoxLayout()

        ### define o rótulo e o campo de texto para o nome do fornecedor ###
        self.labelFornecedor = QLabel("Fornecedor:")
        self.layoutFormularioFornecedor.addWidget(self.labelFornecedor)
        self.campoFornecedor = QLineEdit()
        self.layoutFormularioFornecedor.addWidget(self.campoFornecedor)
        self.listaFornecedor = QListWidget()
        self.layoutFormularioFornecedor.addWidget(self.listaFornecedor)

        ### conecta a entrada de texto à função de atualização da lista ###
        self.campoFornecedor.textChanged.connect(
            lambda: self.atualizarLista(self.fornecedores, self.campoFornecedor, self.listaFornecedor))
        self.listaFornecedor.itemClicked.connect(
            lambda item: self.selecionarItemLista(item, self.campoFornecedor, self.listaFornecedor))

        ### define o rótulo e o dropdown para o critério do fornecedor ###
        self.labelCriterioFornecedor = QLabel("Critério:")
        self.layoutFormularioFornecedor.addWidget(self.labelCriterioFornecedor)
        self.dropdownCriterioFornecedor = QComboBox()
        self.dropdownCriterioFornecedor.addItems(self.criteriosFornecedor.keys())
        self.layoutFormularioFornecedor.addWidget(self.dropdownCriterioFornecedor)

        ### conecta a mudança de critério à função de atualização dos pontos ###
        self.dropdownCriterioFornecedor.currentIndexChanged.connect(self.atualizarPontosFornecedor)

        ### define o rótulo e o campo de texto para os pontos do fornecedor ###
        self.labelPontosFornecedor = QLabel("Pontos:")
        self.layoutFormularioFornecedor.addWidget(self.labelPontosFornecedor)
        self.campoPontosFornecedor = QLineEdit()
        self.campoPontosFornecedor.setReadOnly(True)
        self.layoutFormularioFornecedor.addWidget(self.campoPontosFornecedor)

        ### define o rótulo e o campo de data para a data de lançamento ###
        self.labelDataFornecedor = QLabel("Data:")
        self.layoutFormularioFornecedor.addWidget(self.labelDataFornecedor)
        self.dataFornecedor = QDateEdit()
        self.dataFornecedor.setDisplayFormat('dd/MM/yyyy')
        self.layoutFormularioFornecedor.addWidget(self.dataFornecedor)

        ### botão para submeter os dados do fornecedor ###
        self.botaoSubmeterFornecedor = QPushButton("Lançar Fornecedor")
        self.botaoSubmeterFornecedor.setFixedHeight(40)
        self.botaoSubmeterFornecedor.clicked.connect(self.submeterFornecedor)
        self.layoutFormularioFornecedor.addWidget(self.botaoSubmeterFornecedor)

        self.layoutFornecedor.addLayout(self.layoutFormularioFornecedor)
        self.abaFornecedor.setLayout(self.layoutFornecedor)

    ### função para atualizar a lista com base na entrada do usuário ###
    def atualizarLista(self, listaDados, campoEntrada, listbox):
        termoBusca = campoEntrada.text().lower()
        listbox.clear()
        for item in listaDados:
            if termoBusca in item.lower():
                listbox.addItem(item)

    ### função para selecionar um item da lista e atualizar a entrada ###
    def selecionarItemLista(self, item, campoEntrada, listbox):
        campoEntrada.setText(item.text())

    ### função para atualizar os pontos com base no critério do motorista selecionado ###
    def atualizarPontosMotorista(self):
        criterio = self.dropdownCriterioMotorista.currentText()
        pontos = self.criteriosMotorista.get(criterio, "")
        self.campoPontosMotorista.setText(str(pontos))

    ### função para atualizar os pontos com base no critério do fornecedor selecionado ###
    def atualizarPontosFornecedor(self):
        criterio = self.dropdownCriterioFornecedor.currentText()
        pontos = self.criteriosFornecedor.get(criterio, "")
        self.campoPontosFornecedor.setText(str(pontos))

    ### função para submeter os dados do motorista e salvar na planilha ###
    def submeterMotorista(self):
        motorista = self.campoMotorista.text()
        criterio = self.dropdownCriterioMotorista.currentText()
        pontos = self.campoPontosMotorista.text()
        data = self.dataMotorista.date().toString('dd/MM/yyyy')

        if motorista and criterio and pontos and data:
            functions.adicionarDados(self.caminhoArquivo, 'Lançamento Motorista', [motorista, criterio, pontos, data])
            QMessageBox.information(self, "Sucesso", "Dados do motorista lançados com sucesso!")
        else:
            QMessageBox.warning(self, "Erro", "Por favor, preencha todos os campos.")

    ### função para submeter os dados do fornecedor e salvar na planilha ###
    def submeterFornecedor(self):
        fornecedor = self.campoFornecedor.text()
        criterio = self.dropdownCriterioFornecedor.currentText()
        pontos = self.campoPontosFornecedor.text()
        data = self.dataFornecedor.date().toString('dd/MM/yyyy')

        if fornecedor and criterio and pontos and data:
            functions.adicionarDados(self.caminhoArquivo, 'Lançamento Fornecedor', [fornecedor, criterio, pontos, data])
            QMessageBox.information(self, "Sucesso", "Dados do fornecedor lançados com sucesso!")
        else:
            QMessageBox.warning(self, "Erro", "Por favor, preencha todos os campos.")

### ponto de entrada da aplicação ###
if __name__ == "__main__":
    app = QApplication(sys.argv)

    ### aplica o tema escuro à aplicação ###
    app.setStyle('Fusion')
    paletaEscura = QtGui.QPalette()
    paletaEscura.setColor(QtGui.QPalette.Window, QtGui.QColor(53, 53, 53))
    paletaEscura.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
    paletaEscura.setColor(QtGui.QPalette.Base, QtGui.QColor(25, 25, 25))
    paletaEscura.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(53, 53, 53))
    paletaEscura.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
    paletaEscura.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
    paletaEscura.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
    paletaEscura.setColor(QtGui.QPalette.Button, QtGui.QColor(53, 53, 53))
    paletaEscura.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
    paletaEscura.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
    paletaEscura.setColor(QtGui.QPalette.Link, QtGui.QColor(42, 130, 218))
    paletaEscura.setColor(QtGui.QPalette.Highlight, QtGui.QColor(42, 130, 218))
    paletaEscura.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.black)
    app.setPalette(paletaEscura)

    janela = JanelaPrincipal()
    janela.show()
    sys.exit(app.exec_())
