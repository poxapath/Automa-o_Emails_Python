#1 Importar tabela pelo drive
from google.colab import drive
drive.mount('/content/drive')

import os
planilha = os.listdir('/content/drive/MyDrive/Colab Notebooks/Minicurso de Automação')

# e colocar biblioteca que para ler
import pandas as pd
tabela_vendas = pd.read_excel('/content/drive/MyDrive/Colab Notebooks/Minicurso de Automação/Vendas.xlsx')

#2 Melhor visualização da base de dados
pd.set_option('display.max_columns', None)
tabela_vendas = tabela_vendas.rename(columns={0:'Vendas'})
display(tabela_vendas)

#3 Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
faturamento = faturamento.rename(columns={0:'Faturamento'})
display(faturamento)

#4 Qnt de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
display(quantidade)

#5 Ticket médio (faturamento/qnt de produtos)
# ticket= faturamento/quantidade
ticket = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket = ticket.rename(columns={0:'Ticket Médio'})
display(ticket)

import smtplib
import email.message

def enviar_email():
    corpo_email = f'''
    <p>Prezados,</p>

    <p>Segue o Relatório de Vendas por cada Loja.</p>

    <p>Faturamento:</p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

    <p>Quantidade Vendida:</p>
    {quantidade.to_html()}

    <p>Ticket Médio dos Produtos em cada Loja:</p>
    {ticket.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

    <p>Qualquer dúvida estou à disposição!</p>

    <p>Att.,</p>
    <p>Path :)</p>
    '''

    msg = email.message.Message()
    msg['Subject'] = "Assunto"
    msg['From'] = 'raul.oliveira369@gmail.com'
    msg['To'] = 'raul.oliveira369@gmail.com'
    password = 'nlddkxsaueemleto'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

enviar_email()