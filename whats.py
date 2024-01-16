"""
Automatizar mensagens de whatsapp para todos os contatos da planilha clientes
"""

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


# webbrowser.get('firefox').open_new_tab('https://web.whatsapp.com/')   
# webbrowser.open('https://drogariauniao.com')
webbrowser.open('https://web.whatsapp.com/')
sleep(30)
workbook = openpyxl.load_workbook('clientes.xlsx')

clientes = workbook['Planilha1']

for linha in clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mensagem = f'Olá {nome}! Conforme combinado, estamos te lembrano que sua conta venceu em {vencimento.strftime("%d/%m/%Y")}.'
        
try:
    link_mensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    # webbrowser.get('firefox').open_new_tab(link_mensagem)   
    webbrowser.open(link_mensagem)
    sleep(10)
    seta = pyautogui.locateCenterOnScreen('seta.png')
    sleep(5)
    pyautogui.click(seta[0], seta[1])
    sleep(5)
    pyautogui.hotkey('ctrl','w')
    sleep(5)
except: 
    print(f'Não foi possível enviar para o cliente{nome}, telefone{telefone}')
    with open('erros.csv','a',newline='',encoding='utf-8')as arquivo:
        arquivo.write(f'{nome}, {telefone}')    
    
    