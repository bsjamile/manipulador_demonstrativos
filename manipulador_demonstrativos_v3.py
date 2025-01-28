import os
import shutil
import win32com.client as win32
from PyQt5 import QtCore, QtGui, QtWidgets
from datetime import datetime
import sys

"""
Manipulador de Demonstrativos de Caixa

Este aplicativo permite mover, renomear, remover e enviar por email arquivos de demonstrativos de caixa
que são baixados para a pasta Downloads do usuário. 

Especificidades:
- A data deve ser digitada no formato "DDMMAA" (por exemplo, 291024 para 29 de outubro de 2024).
- O aplicativo procura arquivos que contenha a palavra 'DEMONSTRATIVO' na pasta Downloads.
- Ao mover os arquivos, eles serão organizados em subpastas baseadas no estado (BA, PE, RN, NEOS, DF).
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
# transformar em executavel com icone na janelinha e na barra de tarefas: 
# pyinstaller --onefile --windowed --icon=icone.ico --add-data "icone.ico;." manipulador_demonstrativos.py

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

        # Substituindo o botão "Renomear" por um combo box
        self.combo_renomear = QtWidgets.QComboBox(self.tela)
        self.combo_renomear.addItem("2 - Renomear Arquivos: ")
        self.combo_renomear.addItem("Renomear Todos")
        self.combo_renomear.addItem("Renomear Santander")
        self.combo_renomear.addItem("Renomear Itaú")
        self.combo_renomear.currentIndexChanged.connect(self.renomear_selecionado)  # Conectar à função de renomeação
        button_layout.addWidget(self.combo_renomear)  # Adiciona o combo box no mesmo lugar que o botão de renomear

        self.botao_mover = QtWidgets.QPushButton(self.tela)
        self.botao_mover.setObjectName("botao_mover")
        button_layout.addWidget(self.botao_mover)

        # Substituindo o botão "4 - Abrir" por um combo box
        self.combo_abrir = QtWidgets.QComboBox(self.tela)
        self.combo_abrir.addItem("4 - Abrir: ")
        self.combo_abrir.addItem("Abrir Todos")
        self.combo_abrir.addItem("Abrir Santander")
        self.combo_abrir.addItem("Abrir Itaú")
        self.combo_abrir.addItem("Abrir BA")
        self.combo_abrir.addItem("Abrir NEOS")
        self.combo_abrir.addItem("Abrir PE")
        self.combo_abrir.addItem("Abrir RN")        
        self.combo_abrir.currentIndexChanged.connect(self.abrir_selecionado) # Conectar à função para tratar a opção selecionada        
        button_layout.addWidget(self.combo_abrir) # Adicionar o combo box no layout de botões

        # Substituindo o botão "Enviar" por um combo box
        self.combo_envio = QtWidgets.QComboBox(self.tela)
        self.combo_envio.addItem("5 - Enviar Arquivos: ")
        self.combo_envio.addItem("Enviar Todos")
        self.combo_envio.addItem("Enviar Santander")
        self.combo_envio.addItem("Enviar Itaú")
        self.combo_envio.currentIndexChanged.connect(self.enviar_selecionado)  # Conectar à função de envio
        button_layout.addWidget(self.combo_envio)  # Adiciona o combo box no mesmo lugar que o botão de envio

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

    def retranslateUi(self, Demonstrativos):
        _translate = QtCore.QCoreApplication.translate
        Demonstrativos.setWindowTitle(_translate("Demonstrativos", "Manipulador de Demonstrativos de Caixa"))
        self.botao_remover.setText(_translate("Demonstrativos", "1 - Remover"))
        self.botao_mover.setText(_translate("Demonstrativos", "3 - Mover"))
        self.botao_sair.setText(_translate("Demonstrativos", "6 - Sair"))
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
        estados = ['BA', 'DF', 'NEOS', 'PE', 'RN']
        arquivos_removidos = 0

        for estado in estados:
            folder_remove_demonstrativo = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
            if os.path.exists(folder_remove_demonstrativo):
                for file_to_remove in os.listdir(folder_remove_demonstrativo):
                    file_to_remove_upper = file_to_remove.upper()  # Convertendo o nome do arquivo para maiúsculas
                    if 'DEMONSTRATIVO' in file_to_remove_upper:
                        file_path = os.path.join(folder_remove_demonstrativo, file_to_remove_upper)
                        if os.path.isfile(file_path):
                            try:
                                os.remove(file_path)
                                self.append_message(f"Arquivo {file_to_remove_upper} removido com sucesso.")
                                arquivos_removidos += 1
                            except PermissionError:
                                self.append_message(f"Erro: O arquivo {file_to_remove_upper} não pode ser removido porque está aberto.", error=True)

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

        # Lista todos os arquivos na pasta de Downloads que contêm 'DEMONSTRATIVO' ou 'cashstatement'
        arquivos_downloads = [
            file_name for file_name in os.listdir(downloads_folder)  # Converte para maiúsculas
            if 'DEMONSTRATIVO' in file_name.upper() # Verifica as palavras corretamente
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
                elif 'DF' in file_path:
                    estado = 'DF'  # Considera PGA como NEOS também
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

    # Função para renomear com base na seleção do combo box
    def renomear_selecionado(self):
        opcao = self.combo_renomear.currentText()
        if opcao == "Renomear Todos":
            self.renomear_todos()
        elif opcao == "Renomear Santander":
            self.renomear_santander()
        elif opcao == "Renomear Itaú":
            self.renomear_itau()

    def renomear_todos(self):
        if not self.validate_date():
            return

        data_demonstrativo = self.input_data.text()

        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        lista_nomes = [       
            f'DEMONSTRATIVODECAIXA_BDPE_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_BDRN_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDBA_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDNEOS_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_PGA_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_BDBA_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_BDDF_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_SDDF_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDDF_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_PGADF_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_FADF_{data_demonstrativo}.pdf'
        ]

        arquivos_renomeados = 0

        # Lista os arquivos que contêm 'DEMONSTRATIVO' ou 'CASHSTATEMENT' e converte para maiúsculas
        arquivos_downloads = [
            file_name for file_name in os.listdir(downloads_folder)
            if 'DEMONSTRATIVO' in file_name.upper()  # Verifica as palavras corretamente
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

    def renomear_santander(self):
        if not self.validate_date():
            return

        data_demonstrativo = self.input_data.text()

        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        lista_nomes = [       
            f'DEMONSTRATIVODECAIXA_BDPE_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_BDRN_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDBA_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDNEOS_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_PGA_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_BDBA_{data_demonstrativo}.pdf'
        ]

        arquivos_renomeados = 0

        # Lista os arquivos que contêm 'DEMONSTRATIVO' ou 'CASHSTATEMENT' e converte para maiúsculas
        arquivos_downloads = [
            file_name for file_name in os.listdir(downloads_folder)
            if 'DEMONSTRATIVO' in file_name.upper()  # Verifica as palavras corretamente
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

    def renomear_itau(self):
        if not self.validate_date():
            return

        data_demonstrativo = self.input_data.text()

        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        lista_nomes = [
            f'DEMONSTRATIVODECAIXA_BDDF_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_SDDF_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_CDDF_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_PGADF_{data_demonstrativo}.pdf',
            f'DEMONSTRATIVODECAIXA_FADF_{data_demonstrativo}.pdf'
        ]

        arquivos_renomeados = 0

        # Lista os arquivos que contêm 'DEMONSTRATIVO' ou 'CASHSTATEMENT' e converte para maiúsculas
        arquivos_downloads = [
            file_name for file_name in os.listdir(downloads_folder)
            if 'DEMONSTRATIVO' in file_name.upper()  # Verifica as palavras corretamente
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

    def abrir_selecionado(self):
        # Obtém a opção selecionada no combo box
        opcao = self.combo_abrir.currentText()

        if opcao == "Abrir Todos":
            self.abrir_todos_demonstrativos()
        elif opcao == "Abrir Santander":
            self.abrir_demonstrativos_santander()
        elif opcao == "Abrir Itaú":
            self.abrir_demonstrativos_itau()
        elif opcao == "Abrir BA":
            self.abrir_demonstrativos_BA()
        elif opcao == "Abrir NEOS":
            self.abrir_demonstrativos_NEOS()
        elif opcao == "Abrir PE":
            self.abrir_demonstrativos_PE()
        elif opcao == "Abrir RN":
            self.abrir_demonstrativos_RN()
    
    def abrir_todos_demonstrativos(self):
        if not self.validate_date():
            return

        ano, mes = self.get_ano_mes()
        estados = ['BA', 'DF','NEOS', 'PE', 'RN']
        arquivos_abertos = 0

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
            
            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO'
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) if 'DEMONSTRATIVO' in f.upper()
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

    def abrir_demonstrativos_santander(self):
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
                    f for f in os.listdir(pasta_demonstrativos) if 'DEMONSTRATIVO' in f.upper()
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

    def abrir_demonstrativos_itau(self):
        if not self.validate_date():
            return

        ano, mes = self.get_ano_mes()
        estados = ['DF']
        arquivos_abertos = 0

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
            
            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO'
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) if 'DEMONSTRATIVO' in f.upper()
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

    def abrir_demonstrativos_BA(self):
        if not self.validate_date():
            return

        ano, mes = self.get_ano_mes()
        estados = ['BA']
        arquivos_abertos = 0

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
            
            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO'
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) if 'DEMONSTRATIVO' in f.upper()
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

    def abrir_demonstrativos_NEOS(self):
        if not self.validate_date():
            return

        ano, mes = self.get_ano_mes()
        estados = ['NEOS']
        arquivos_abertos = 0

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
            
            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO'
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) if 'DEMONSTRATIVO' in f.upper()
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

    def abrir_demonstrativos_PE(self):
        if not self.validate_date():
            return

        ano, mes = self.get_ano_mes()
        estados = ['PE']
        arquivos_abertos = 0

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
            
            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO'
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) if 'DEMONSTRATIVO' in f.upper()
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

    def abrir_demonstrativos_RN(self):
        if not self.validate_date():
            return

        ano, mes = self.get_ano_mes()
        estados = ['RN']
        arquivos_abertos = 0

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
            
            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO'
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) if 'DEMONSTRATIVO' in f.upper()
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

    def enviar_selecionado(self):
        opcao = self.combo_envio.currentText()
        if opcao == "Enviar Todos":
            self.enviar_todos_demonstrativos()
        elif opcao == "Enviar Santander":
            self.enviar_demonstrativos_santander()
        elif opcao == "Enviar Itaú":
            self.enviar_demonstrativos_itau()

    def enviar_todos_demonstrativos(self):
        if not self.validate_date():
            return

        # Obter a data digitada (DDMMAA)
        data_demonstrativo = self.input_data.text()
        
        # Formatar o assunto para incluir a data digitada
        dia = data_demonstrativo[:2]
        mes_ano = data_demonstrativo[2:4] + data_demonstrativo[4:6]
        mes= data_demonstrativo[2:4]
        ano = "20" + data_demonstrativo[4:6]  # Formato de ano completo (20YY)

        # ano, mes = self.get_ano_mes()
        estados = ['BA', 'DF', 'NEOS', 'PE', 'RN']
        arquivos_abertos = 0

        # Criando o aplicativo do Outlook
        try:
            outlook = win32.Dispatch('Outlook.Application')
            email = outlook.CreateItem(0)  # 0 indica que é um e-mail
        except Exception as e:
            self.append_message(f"Erro ao iniciar o Outlook: {e}", error=True)
            return
        
        email.Subject = f'DEMONSTRATIVOS DE CAIXA | 01 a {dia}/{mes}/{ano}'
        corpo_email = f"Olá!\n\nSegue Demonstrativos de caixa de 01 a {dia}/{mes}/{ano}.\n\nAté mais,"
        email.Body = corpo_email  # Usamos apenas texto simples (sem HTML)
        
        # Destinatário principal e CC
        email.To = 'marcia.valente@neosprevidencia.com.br;vania.barbosa@neosprevidencia.com.br;eron.sampaio@neosprevidencia.com.br;financeiro@neosprevidencia.com.br'
        email.CC = 'gac@neosprevidencia.com.br'

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes_ano}\\{estado}\\'
            self.append_message(f"Verificando pasta: {pasta_demonstrativos}")  # Depuração

            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO' 
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) 
                    if 'DEMONSTRATIVO' in f.upper()
                ]
                
                # Depuração: Verificando quais arquivos foram encontrados
                if arquivos_demonstrativos:
                    self.append_message(f"Arquivos encontrados na pasta {pasta_demonstrativos}: {arquivos_demonstrativos}")
                    # Anexa cada arquivo encontrado ao e-mail
                    for arquivo in arquivos_demonstrativos:
                        caminho_completo = os.path.join(pasta_demonstrativos, arquivo)
                        try:
                            email.Attachments.Add(caminho_completo)  # Adiciona o arquivo como anexo
                            self.append_message(f"Anexando arquivo: {arquivo}")
                            arquivos_abertos += 1
                        except Exception as e:
                            self.append_message(f"Erro ao anexar o arquivo {arquivo}: {e}", error=True)
                else:
                    self.append_message(f"Nenhum arquivo 'DEMONSTRATIVO' encontrado na pasta {pasta_demonstrativos}.", error=True)
            else:
                self.append_message(f"Pasta não encontrada: {pasta_demonstrativos}", error=True)

        if arquivos_abertos > 0:
            # Abrir o e-mail para revisão antes de enviar
            email.Display()  # Isso abrirá o e-mail no Outlook para revisão
            self.append_message(f"Abrindo e-mail com {arquivos_abertos} anexos.")
        else:
            self.append_message("Nenhum arquivo 'DEMONSTRATIVO' encontrado nas pastas.", error=True)

    def enviar_demonstrativos_santander(self):
        if not self.validate_date():
            return

        # Obter a data digitada (DDMMAA)
        data_demonstrativo = self.input_data.text()
        
        # Formatar o assunto para incluir a data digitada
        dia = data_demonstrativo[:2]
        mes_ano = data_demonstrativo[2:4] + data_demonstrativo[4:6]
        mes= data_demonstrativo[2:4]
        ano = "20" + data_demonstrativo[4:6]  # Formato de ano completo (20YY)

        # ano, mes = self.get_ano_mes()
        estados = ['BA','NEOS', 'PE', 'RN']
        arquivos_abertos = 0

        # Criando o aplicativo do Outlook
        try:
            outlook = win32.Dispatch('Outlook.Application')
            email = outlook.CreateItem(0)  # 0 indica que é um e-mail
        except Exception as e:
            self.append_message(f"Erro ao iniciar o Outlook: {e}", error=True)
            return
        
        email.Subject = f'DEMONSTRATIVOS DE CAIXA | 01 a {dia}/{mes}/{ano}'
        corpo_email = f"Olá!\n\nSegue Demonstrativos de caixa de 01 a {dia}/{mes}/{ano}.\n\nAté mais,"
        email.Body = corpo_email  # Usamos apenas texto simples (sem HTML)
        
        # Destinatário principal e CC
        email.To = 'marcia.valente@neosprevidencia.com.br;vania.barbosa@neosprevidencia.com.br;eron.sampaio@neosprevidencia.com.br;financeiro@neosprevidencia.com.br'
        email.CC = 'gac@neosprevidencia.com.br'

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes_ano}\\{estado}\\'
            self.append_message(f"Verificando pasta: {pasta_demonstrativos}")  # Depuração

            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO' 
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) 
                    if 'DEMONSTRATIVO' in f.upper()
                ]
                
                # Depuração: Verificando quais arquivos foram encontrados
                if arquivos_demonstrativos:
                    self.append_message(f"Arquivos encontrados na pasta {pasta_demonstrativos}: {arquivos_demonstrativos}")
                    # Anexa cada arquivo encontrado ao e-mail
                    for arquivo in arquivos_demonstrativos:
                        caminho_completo = os.path.join(pasta_demonstrativos, arquivo)
                        try:
                            email.Attachments.Add(caminho_completo)  # Adiciona o arquivo como anexo
                            self.append_message(f"Anexando arquivo: {arquivo}")
                            arquivos_abertos += 1
                        except Exception as e:
                            self.append_message(f"Erro ao anexar o arquivo {arquivo}: {e}", error=True)
                else:
                    self.append_message(f"Nenhum arquivo 'DEMONSTRATIVO' encontrado na pasta {pasta_demonstrativos}.", error=True)
            else:
                self.append_message(f"Pasta não encontrada: {pasta_demonstrativos}", error=True)

        if arquivos_abertos > 0:
            # Abrir o e-mail para revisão antes de enviar
            email.Display()  # Isso abrirá o e-mail no Outlook para revisão
            self.append_message(f"Abrindo e-mail com {arquivos_abertos} anexos.")
        else:
            self.append_message("Nenhum arquivo 'DEMONSTRATIVO' encontrado nas pastas.", error=True)

    def enviar_demonstrativos_itau(self):
        if not self.validate_date():
            return

        # Obter a data digitada (DDMMAA)
        data_demonstrativo = self.input_data.text()
        
        # Formatar o assunto para incluir a data digitada
        dia = data_demonstrativo[:2]
        mes_ano = data_demonstrativo[2:4] + data_demonstrativo[4:6]
        mes= data_demonstrativo[2:4]
        ano = "20" + data_demonstrativo[4:6]  # Formato de ano completo (20YY)

        # ano, mes = self.get_ano_mes()
        estados = ['DF']
        arquivos_abertos = 0

        # Criando o aplicativo do Outlook
        try:
            outlook = win32.Dispatch('Outlook.Application')
            email = outlook.CreateItem(0)  # 0 indica que é um e-mail
        except Exception as e:
            self.append_message(f"Erro ao iniciar o Outlook: {e}", error=True)
            return
        
        email.Subject = f'DEMONSTRATIVOS DE CAIXA | 01 a {dia}/{mes}/{ano}'
        corpo_email = f"Olá!\n\nSegue Demonstrativos de caixa de 01 a {dia}/{mes}/{ano}.\n\nAté mais,"
        email.Body = corpo_email  # Usamos apenas texto simples (sem HTML)
        
        # Destinatário principal e CC
        email.To = 'marcia.valente@neosprevidencia.com.br;vania.barbosa@neosprevidencia.com.br;eron.sampaio@neosprevidencia.com.br;financeiro@neosprevidencia.com.br'
        email.CC = 'gac@neosprevidencia.com.br'

        for estado in estados:
            pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes_ano}\\{estado}\\'
            self.append_message(f"Verificando pasta: {pasta_demonstrativos}")  # Depuração

            if os.path.exists(pasta_demonstrativos):
                # Procura arquivos que contenham 'DEMONSTRATIVO' 
                arquivos_demonstrativos = [
                    f for f in os.listdir(pasta_demonstrativos) 
                    if 'DEMONSTRATIVO' in f.upper()
                ]
                
                # Depuração: Verificando quais arquivos foram encontrados
                if arquivos_demonstrativos:
                    self.append_message(f"Arquivos encontrados na pasta {pasta_demonstrativos}: {arquivos_demonstrativos}")
                    # Anexa cada arquivo encontrado ao e-mail
                    for arquivo in arquivos_demonstrativos:
                        caminho_completo = os.path.join(pasta_demonstrativos, arquivo)
                        try:
                            email.Attachments.Add(caminho_completo)  # Adiciona o arquivo como anexo
                            self.append_message(f"Anexando arquivo: {arquivo}")
                            arquivos_abertos += 1
                        except Exception as e:
                            self.append_message(f"Erro ao anexar o arquivo {arquivo}: {e}", error=True)
                else:
                    self.append_message(f"Nenhum arquivo 'DEMONSTRATIVO' encontrado na pasta {pasta_demonstrativos}.", error=True)
            else:
                self.append_message(f"Pasta não encontrada: {pasta_demonstrativos}", error=True)

        if arquivos_abertos > 0:
            # Abrir o e-mail para revisão antes de enviar
            email.Display()  # Isso abrirá o e-mail no Outlook para revisão
            self.append_message(f"Abrindo e-mail com {arquivos_abertos} anexos.")
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


