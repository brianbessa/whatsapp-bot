"""
CLIENTE NECESSITA DE UM BOT PARA AUTOMATIZAR MENSAGENS DE COBRANÇA DE BOLETO BASEADO EM UMA TABELA DE CLIENTES COM VENCIMENTOS EM DIAS 
DIFERENTES

"""

# ler a planilha de dados e guardar as informações do cliente
# entrar no wpp web
# autenticar a conta
# criar e enviar a mensagem personalizada para cada cliente
# clicar no botão

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

planilha = openpyxl.load_workbook('clientes.xlsx')  
clientes = planilha['Plan1']

for linha in clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    data_vencimento = linha[2].value
    #t print(nome, telefone, data_vencimento)
    #msg-base: https://web.whatsapp.com/send?phone=&text=

    mensagem = f'Olá, {nome}\nVimos que seu boleto com vencimento para o dia {data_vencimento.strftime('%d/%m/%Y')} está proximo.\nRealize o pagamento a tempo para evitar multas por atraso.'

    try:
        link = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}' 
        webbrowser.open(link)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('send.png')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}{os.linesep}')
#t print(planilha['Plan1'])
