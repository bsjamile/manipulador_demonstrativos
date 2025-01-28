# import os
# import win32com.client as win32

# class EnvioDeDemonstrativos:
#     def __init__(self):
#         pass  # Aqui você pode inicializar atributos se necessário
    
#     def append_message(self, message, error=False):
#         # Apenas para exemplo, você pode usar isso para depuração ou log
#         if error:
#             print(f"[ERRO] {message}")
#         else:
#             print(message)
    
#     def validate_date(self):
#         # Simulação de validação da data
#         return True
    
#     def get_ano_mes(self):
#         # Simulação de obtenção do ano e mês (substitua com a lógica correta)
#         return '2025', '0125'
    
#     def enviar_demonstrativos(self):
#         if not self.validate_date():
#             return

#         ano, mes = self.get_ano_mes()
#         estados = ['BA', 'BSB', 'NEOS', 'PE', 'RN']
#         arquivos_abertos = 0

#         # Criando o aplicativo do Outlook
#         try:
#             outlook = win32.Dispatch('Outlook.Application')
#             email = outlook.CreateItem(0)  # 0 indica que é um e-mail
#         except Exception as e:
#             self.append_message(f"Erro ao iniciar o Outlook: {e}", error=True)
#             return
        
#         email.Subject = 'Assunto do e-mail com anexos'
#         email.Body = 'Olá, este é um e-mail enviado automaticamente pelo Python com os arquivos anexados.'
        
#         # Destinatário principal e CC
#         email.To = 'gac@neosprevidencia.com.br'
#         email.CC = 'marica@meaosoasj.com.br;vania@ahduhausdh.com.br'
#         email.BCC = 'joao@exemplo.com;luana@exemplo.com'

#         for estado in estados:
#             pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
#             self.append_message(f"Verificando pasta: {pasta_demonstrativos}")  # Depuração

#             if os.path.exists(pasta_demonstrativos):
#                 # Procura arquivos que contenham 'DEMONSTRATIVO' ou 'CASHSTATEMENT'
#                 arquivos_demonstrativos = [
#                     f for f in os.listdir(pasta_demonstrativos) 
#                     if 'DEMONSTRATIVO' in f.upper() or 'CASHSTATEMENT' in f.upper()
#                 ]
                
#                 # Depuração: Verificando quais arquivos foram encontrados
#                 if arquivos_demonstrativos:
#                     self.append_message(f"Arquivos encontrados na pasta {pasta_demonstrativos}: {arquivos_demonstrativos}")
#                     # Anexa cada arquivo encontrado ao e-mail
#                     for arquivo in arquivos_demonstrativos:
#                         caminho_completo = os.path.join(pasta_demonstrativos, arquivo)
#                         try:
#                             email.Attachments.Add(caminho_completo)  # Adiciona o arquivo como anexo
#                             self.append_message(f"Anexando arquivo: {arquivo}")
#                             arquivos_abertos += 1
#                         except Exception as e:
#                             self.append_message(f"Erro ao anexar o arquivo {arquivo}: {e}", error=True)
#                 else:
#                     self.append_message(f"Nenhum arquivo 'DEMONSTRATIVO' ou 'CASHSTATEMENT' encontrado na pasta {pasta_demonstrativos}.", error=True)
#             else:
#                 self.append_message(f"Pasta não encontrada: {pasta_demonstrativos}", error=True)

#         if arquivos_abertos > 0:
#             # Abrir o e-mail para revisão antes de enviar
#             email.Display()  # Isso abrirá o e-mail no Outlook para revisão
#             self.append_message(f"Abrindo e-mail com {arquivos_abertos} anexos.")
#         else:
#             self.append_message("Nenhum arquivo 'DEMONSTRATIVO' encontrado nas pastas.", error=True)

# # Criando uma instância da classe
# envio = EnvioDeDemonstrativos()

# # Chamando o método para enviar os demonstrativos
# envio.enviar_demonstrativos()

import os
import win32com.client as win32

def enviar_demonstrativos(self):
    if not self.validate_date():
        return

    ano, mes = self.get_ano_mes()
    estados = ['BA', 'BSB', 'NEOS', 'PE', 'RN']
    arquivos_abertos = 0

    # Criando o aplicativo do Outlook
    try:
        outlook = win32.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)  # 0 indica que é um e-mail
    except Exception as e:
        self.append_message(f"Erro ao iniciar o Outlook: {e}", error=True)
        return
    
    email.Subject = f'DEMONSTRATIVOS DE CAIXA | {mes}/{ano}'
    email.Body = 'Prezados, Segue .'
    
    # Destinatário principal e CC
    email.To = 'gac@neosprevidencia.com.br'
    email.CC = 'marica@meaosoasj.com.br;vania@ahduhausdh.com.br'
    email.BCC = 'joao@exemplo.com;luana@exemplo.com'

    for estado in estados:
        pasta_demonstrativos = f'H:\\GAC\\Relatórios Santander\\Carteira Custódia\\{ano}\\{mes}\\{estado}\\'
        self.append_message(f"Verificando pasta: {pasta_demonstrativos}")  # Depuração

        if os.path.exists(pasta_demonstrativos):
            # Procura arquivos que contenham 'DEMONSTRATIVO' ou 'CASHSTATEMENT'
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
                self.append_message(f"Nenhum arquivo 'DEMONSTRATIVO' ou 'CASHSTATEMENT' encontrado na pasta {pasta_demonstrativos}.", error=True)
        else:
            self.append_message(f"Pasta não encontrada: {pasta_demonstrativos}", error=True)

    if arquivos_abertos > 0:
        # Abrir o e-mail para revisão antes de enviar
        email.Display()  # Isso abrirá o e-mail no Outlook para revisão
        self.append_message(f"Abrindo e-mail com {arquivos_abertos} anexos.")
    else:
        self.append_message("Nenhum arquivo 'DEMONSTRATIVO' encontrado nas pastas.", error=True)