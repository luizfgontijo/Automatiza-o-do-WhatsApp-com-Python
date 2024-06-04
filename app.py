import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

workbook = openpyxl.load_workbook('planilha.xlsx')
pag = workbook['Planilha1']

for linha in pag.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value

    mensagem = f'Olá {nome} estou testando o python!'
    
    try:
        link_msg = f'https://web.whatsapp.com/send?phone={
        telefone}&text={quote(mensagem)}'
        webbrowser.open(link_msg)
        sleep(15)
        pyautogui.hotkey('enter')
        sleep(10)
        pyautogui.hotkey('ctrl', 'w')
        sleep(10)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
