import os
import shutil
from PyQt5 import QtCore, QtGui, QtWidgets
from datetime import datetime
import sys

"""
Manipulador de Demonstrativos de Caixa

Este aplicativo permite mover, renomear e remover arquivos de demonstrativos de caixa
que são baixados para a pasta Downloads do usuário. 

Especificidades:
- A data deve ser digitada no formato "DDMMAA" (por exemplo, 291024 para 29 de outubro de 2024).
- O aplicativo procura arquivos que contenham as palavras 'DEMONSTRATIVO' ou 'DEMONSTRATIVODECAIXA' na pasta Downloads.
- Ao mover os arquivos, eles serão organizados em subpastas baseadas no estado (BA, PE, RN, NEOS).
- O aplicativo utiliza a biblioteca PyQt5 para a interface gráfica e requer que ela esteja instalada.

Requisitos:
- Python 3.x
- PyQt5
- Acesso à pasta Downloads e às pastas de destino para movimentação dos arquivos.

Uso:
1. Execute o script.
2. Insira a data desejada no campo de entrada.
3. Clique nos botões para realizar as ações de remover, mover ou renomear os arquivos.
"""

class Ui_Demonstrativos(object):
    def setupUi(self, Demonstrativos):
        Demonstrativos.setObjectName("Demonstrativos")
        Demonstrativos.resize(390, 400)

                # Defina o caminho do ícone
        if getattr(sys, 'frozen', False):
            # Caminho para o ícone ao rodar como executável
            icon_path = os.path.join(sys._MEIPASS, "icone.ico")
        else:
            # Caminho para o ícone ao rodar como script Python
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icone.ico")

        Demonstrativos.setWindowIcon(QtGui.QIcon(icon_path))  # Configure o ícone da janela

        self.tela = QtWidgets.QWidget(Demonstrativos)
        self.tela.setObjectName("tela")

        # Criando um layout vertical principal
        main_layout = QtWidgets.QVBoxLayout(self.tela)

        self.explicaco = QtWidgets.QLabel(self.tela)
        self.explicaco.setObjectName("explicaco")
        self.explicaco.setAlignment(QtCore.Qt.AlignCenter)
        main_layout.addWidget(self.explicaco)

        self.label_data = QtWidgets.QLabel(self.tela)
        self.label_data.setText("Data (DDMMAA):")
        main_layout.addWidget(self.label_data)

        self.input_data = QtWidgets.QLineEdit(self.tela)
        main_layout.addWidget(self.input_data)

        # Criando um layout horizontal para os botões
        button_layout = QtWidgets.QHBoxLayout()

        self.botao_remover = QtWidgets.QPushButton(self.tela)
        self.botao_remover.setObjectName("botao_remover")
        button_layout.addWidget(self.botao_remover)

        self.botao_renomear = QtWidgets.QPushButton(self.tela)
        self.botao_renomear.setObjectName("botao_renomear")
        button_layout.addWidget(self.botao_renomear)

        self.botao_mover = QtWidgets.QPushButton(self.tela)
        self.botao_mover.setObjectName("botao_mover")
        button_layout.addWidget(self.botao_mover)

        self.botao_abrir = QtWidgets.QPushButton(self.tela)
        self.botao_abrir.setObjectName("botao_abrir")
        button_layout.addWidget(self.botao_abrir)

        self.botao_sair = QtWidgets.QPushButton(self.tela)
        self.botao_sair.setObjectName("botao_sair")
        button_layout.addWidget(self.botao_sair)

        main_layout.addLayout(button_layout)

        # Área de mensagem abaixo dos botões
        self.mensagem_area = QtWidgets.QTextEdit(self.tela)
        self.mensagem_area.setReadOnly(True)
        self.mensagem_area.setObjectName("mensagem_area")
        main_layout.addWidget(self.mensagem_area)

        Demonstrativos.setCentralWidget(self.tela)

        self.retranslateUi(Demonstrativos)
        self.botao_sair.clicked.connect(Demonstrativos.close)
        self.botao_remover.clicked.connect(self.remove)
        self.botao_mover.clicked.connect(self.mover)
        self.botao_renomear.clicked.connect(self.renomear)
        self.botao_abrir.clicked.connect(self.abrir_demonstrativos)

    def retranslateUi(self, Demonstrativos):
        _translate = QtCore.QCoreApplication.translate
        Demonstrativos.setWindowTitle(_translate("Demonstrativos", "Manipulador de Demonstrativos de Caixa"))
        self.botao_remover.setText(_translate("Demonstrativos", "1 - Remover"))
        self.botao_renomear.setText(_translate("Demonstrativos", "2 - Renomear"))
        self.botao_mover.setText(_translate("Demonstrativos", "3 - Mover"))
        self.botao_abrir.setText(_translate("Demonstrativos", "4 - Abrir Demonstrativos"))
        self.botao_sair.setText(_translate("Demonstrativos", "5 - Sair"))
        self.explicaco.setText(_translate("Demonstrativos", "<html><head/><body><p align=\"center\"><span style=\" font-size:10pt; font-weight:600;\">Manipular Demonstrativos de Caixa </span></p><p align=\"center\"><span style=\" font-size:10pt; font-weight:600;\">conforme opções:</span></p></body></html>"))

    def append_message(self, message, error=False):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        separator = "-----------------------"
        if error:
            message = f"<span style='color:red;'>{message}</span>"
        self.mensagem_area.append(f"{separator}<br>{timestamp}<br>{message}")

    def validate_date(self):
        data_demonstrativo = self.input_data.text()
        if len(data_demonstrativo) != 6 or not data_demonstrativo.isdigit():
            QtWidgets.QMessageBox.warning(None, "Erro", "A data deve ter 6 dígitos (DDMMAA) e ser numérica.")
            return False
        return True

    def remove(self):
        if not self.validate_date():
            return

        ano, mes = self.get_ano_mes()
        estados = ['BA', 'NEOS', 'PE', 'RN']
        arquivos_removidos = 0

        for estado in estados:
            folder_remove_demonstrativo = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
            if os.path.exists(folder_remove_demonstrativo):
                for file_to_remove in os.listdir(folder_remove_demonstrativo):
                    if 'DEMONSTRATIVO' in file_to_remove:
                        file_path = os.path.join(folder_remove_demonstrativo, file_to_remove)
                        if os.path.isfile(file_path):
                            try:
                                os.remove(file_path)
                                self.append_message(f"Arquivo {file_to_remove} removido com sucesso.")
                                arquivos_removidos += 1
                            except PermissionError:
                                self.append_message(f"Erro: O arquivo {file_to_remove} não pode ser removido porque está aberto.", error=True)

        if arquivos_removidos > 0:
            self.append_message(f"Total de {arquivos_removidos} arquivos removidos com sucesso!")
        else:
            self.append_message("Nenhum arquivo removido.")

    def mover(self):
        if not self.validate_date():
            return

        ano, mes = self.get_ano_mes()
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        arquivos_movidos = 0

        # Lista todos os arquivos na pasta de Downloads que contêm 'DEMONSTRATIVO'
        arquivos_downloads = [
            file_name for file_name in os.listdir(downloads_folder)
            if 'DEMONSTRATIVO' in file_name
        ]

        # Filtra os caminhos completos dos arquivos
        arquivos_downloads_full_path = [
            os.path.join(downloads_folder, file_name) for file_name in arquivos_downloads
        ]

        # Ordena os arquivos pela data de modificação
        arquivos_downloads_full_path.sort(key=os.path.getmtime, reverse=True)

        if not arquivos_downloads_full_path:
            self.append_message("Nenhum arquivo 'DEMONSTRATIVO' encontrado.")
            return

        # Obtém a data do arquivo mais recente (apenas data, ignorando horas)
        data_mais_recente = int(os.path.getmtime(arquivos_downloads_full_path[0])) // (24 * 3600)

        # Mover os arquivos que têm a mesma data
        for file_path in arquivos_downloads_full_path:
            if int(os.path.getmtime(file_path)) // (24 * 3600) == data_mais_recente:
                estado = None
                if 'BA' in file_path:
                    estado = 'BA'
                elif 'PE' in file_path:
                    estado = 'PE'
                elif 'RN' in file_path:
                    estado = 'RN'
                elif 'NEOS' in file_path:
                    estado = 'NEOS'
                elif 'PGA' in file_path:
                    estado = 'NEOS'  # Considera PGA como NEOS também

                if estado:
                    new_folder = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
                    os.makedirs(new_folder, exist_ok=True)  # Cria a pasta se não existir
                    new_name = os.path.join(new_folder, os.path.basename(file_path))

                    # Verifica se o arquivo já existe no destino
                    if os.path.exists(new_name):
                        self.append_message(f"Erro: O arquivo {os.path.basename(file_path)} já existe em {new_folder}.", error=True)
                    else:
                        # Move o arquivo
                        shutil.move(file_path, new_name)  
                        self.append_message(f"Arquivo {os.path.basename(file_path)} movido para {new_folder}")
                        arquivos_movidos += 1

        if arquivos_movidos > 0:
            self.append_message(f"Total de {arquivos_movidos} arquivos movidos com sucesso!")
        else:
            self.append_message("Nenhum arquivo movido.")

    def renomear(self):
        if not self.validate_date():
            return

        data_demonstrativo = self.input_data.text()
        ano, mes = self.get_ano_mes(data_demonstrativo)

        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        lista_nomes = [
            f'DEMONSTRATIVODECAIXA_BDBA_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_BDPE_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_BDRN_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDBA_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDPE_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDRN_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDNEOS_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_PGA_{data_demonstrativo}.pdf'
        ]

        arquivos_renomeados = 0
        arquivos_downloads = [
            file_name for file_name in os.listdir(downloads_folder)
            if 'DEMONSTRATIVODECAIXA' in file_name
        ]

        # Ordena os arquivos pela data de modificação (sem horas e minutos)
        arquivos_downloads.sort(key=lambda x: os.path.getmtime(os.path.join(downloads_folder, x)) // (24 * 3600), reverse=True)

        # Obtém a data do arquivo mais recente (apenas data, ignorando horas)
        if not arquivos_downloads:
            self.append_message("Nenhum arquivo 'DEMONSTRATIVODECAIXA' encontrado.")
            return

        data_mais_recente = int(os.path.getmtime(os.path.join(downloads_folder, arquivos_downloads[0]))) // (24 * 3600)

        # Renomeia os arquivos com a data mais recente
        for file_name in arquivos_downloads:
            file_path = os.path.join(downloads_folder, file_name)

            # Renomeia apenas os arquivos com a mesma data
            if int(os.path.getmtime(file_path)) // (24 * 3600) == data_mais_recente:
                index = arquivos_downloads.index(file_name)
                if index < len(lista_nomes):
                    new_name = os.path.join(downloads_folder, lista_nomes[index])

                    if os.path.isfile(file_path):
                        if os.path.exists(new_name):
                            self.append_message(f"Erro: O arquivo {lista_nomes[index]} já existe na pasta.", error=True)
                        else:
                            os.rename(file_path, new_name)
                            self.append_message(f"Arquivo {file_name} renomeado para {lista_nomes[index]}")
                            arquivos_renomeados += 1

        if arquivos_renomeados > 0:
            self.append_message(f"Total de {arquivos_renomeados} arquivos renomeados com sucesso!")
        else:
            self.append_message("Nenhum arquivo renomeado.")

    def abrir_demonstrativos(self):
        if not self.validate_date():
            return

        ano, mes = self.get_ano_mes()
        estados = ['BA', 'NEOS', 'PE', 'RN']
        arquivos_abertos = 0

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
            
            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO'
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) if 'DEMONSTRATIVO' in f
                ]

                if arquivos_demonstrativos:
                    # Abre cada arquivo encontrado
                    for arquivo in arquivos_demonstrativos:
                        caminho_completo = os.path.join(pasta_demonstrativos, arquivo)
                        self.append_message(f"Abrindo arquivo: {arquivo}")  # Mensagem sobre o arquivo sendo aberto
                        os.startfile(caminho_completo)  # Abre o arquivo com o programa padrão
                        arquivos_abertos += 1

        if arquivos_abertos > 0:
            self.append_message(f"Abrindo {arquivos_abertos} arquivos.")
        else:
            self.append_message("Nenhum arquivo 'DEMONSTRATIVO' encontrado nas pastas.", error=True)

    def get_ano_mes(self, data=None):
        if data is None:
            data = self.input_data.text()
        ano = "20" + data[-2:]  # Extrai os últimos dois dígitos para o ano
        mes = data[2:]  # Extrai os dígitos do mês (MM)
        return ano, mes

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Demonstrativos = QtWidgets.QMainWindow()
    ui = Ui_Demonstrativos()
    ui.setupUi(Demonstrativos)
    Demonstrativos.show()
    sys.exit(app.exec_())

# transformar em executavel com icone na janelinha e na barra de tarefas: 
# pyinstaller --onefile --windowed --icon=icone.ico --add-data "icone.ico;." manipulador_demonstrativos.py
