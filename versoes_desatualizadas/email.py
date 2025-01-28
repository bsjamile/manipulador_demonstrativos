import os
import win32com.client as win32

try:
    # Criando o aplicativo do Outlook
    outlook = win32.Dispatch('Outlook.Application')
    print("Outlook iniciado com sucesso")

    # Criando o e-mail
    email = outlook.CreateItem(0)  # 0 indica que é um e-mail
    email.Subject = 'Assunto do e-mail com anexo'
    email.Body = 'Olá, este é um e-mail enviado automaticamente pelo Python com um arquivo anexo.'
    email.To = 'gac@neosprevidencia.com.br'
    email.CC = 'marica@meaosoasj.com.br;vania@ahduhausdh.com.br'

    # Caminho da pasta onde os arquivos estão
    pasta = 'C:/Users/jamile.santos/Downloads/'  # Modifique para o caminho correto da sua pasta

    for arquivo in os.listdir(pasta):
        # Verifica se o nome do arquivo contém a palavra "DEMONSTRATIVO"
        if 'DEMONSTRATIVO' in arquivo:
            caminho_arquivo = os.path.join(pasta, arquivo)
            email.Attachments.Add(caminho_arquivo)  # Adicionando o anexo
            print(f'Anexo adicionado: {arquivo}')

    # Abrindo o e-mail para verificação
    email.Display()  # Isso abrirá o e-mail no Outlook para revisão

    print('E-mail aberto para verificação.')

except Exception as e:
    print(f'Ocorreu um erro: {e}')