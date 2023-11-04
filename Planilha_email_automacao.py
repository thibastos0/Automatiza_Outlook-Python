# pandas -> bases de dados
# os -> arquivos do computador
# pywin32 -> enviar email
import os
from datetime import datetime
# import time
import pandas as pd
import win32com.client as win32 # pip install pywin32

caminho = "dados"
arquivos = os.listdir(caminho)
print(arquivos)

# apenas aproveitando a lib win32 e testando ferramenta de leitura
# speaker = win32.Dispatch("SAPI.SpVoice")
# speaker.Speak("Enviando os seus dados!")

tabela_consolidada = pd.DataFrame()

for nome_arquivo in arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho, nome_arquivo))
    # tratar as datas no csv que estão apresentadas como número. Data "1" = 01/01/1900, no Excel.
    tabela_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas["Data de Venda"], unit="d")
    # juntando os arquivos csv em um único
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda") # ordenar por data
tabela_consolidada = tabela_consolidada.reset_index(drop=True) # reinicia contagem índice
tabela_consolidada.to_excel("Vendas.xlsx", index=False) # desconsiderando o index para salvar o arquivo Excel

# envio do e-mail
outlook = win32.Dispatch('outlook.application')

# verificar contas de e-mail disponíveis
for accounts in outlook.Session.Accounts:
    print(f'Conta de e-mail disponível: {accounts}')

# Para enviar a partir de outra conta no Outlook que não seja a principal:
conta_email = outlook.Session.Accounts['thiago.luiz@rio.br']

email = outlook.CreateItem(0)

# vincular o objeto e-mail a conta selecionada no objeto Outlook
email._oleobj_.Invoke(*(64209, 0, 8, 0, conta_email))

# preparando os detalhes para o envio da mensagem
email.To = "thiago.luiz@teste.gov"
data_hoje = datetime.today().strftime("%d/%m/%Y")
email.Subject = f"Relatório de Vendas {data_hoje}"
# .HTMLBody, caso queira utilizar html
email.Body = f"""
Prezados,

Segue em anexo o Relatório de Vendas de {data_hoje} atualizado.
Qualquer coisa estou à disposição.
Atenciosamente,
Thiago.
"""

# pegando o caminho da pasta atual (onde está o .py) e anexando o arquivo salvo ao e-mail
caminho = os.getcwd()
anexo = os.path.join(caminho, "Vendas.xlsx")
email.Attachments.Add(anexo)

# getting default folder of used Account
# myNamespace = outlook.GetNamespace("MAPI")
# myFolder = myNamespace.GetDefaultFolder(6)
# myFolder.Display()

# Exibir o e-mail antes de enviar
# email.Display()
# time.sleep(5)
email.Send()