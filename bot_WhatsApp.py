import openpyxl
from urllib.parse import quote
from time import sleep

import pyautogui
import webbrowser

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

workbook = openpyxl.load_workbook('teste.xlsx')
pagina_clientes = workbook['Plan1']
pyautogui.hotkey('ctrl','w')

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    pix = 'pix celular 92991200836'

    mensagem = f'Olá {nome} estou fazendo um teste no meu programa de automação do WhatsApp para meus clientes de IPTV seu vencimento é dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar com o {pix}'

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        
    except:
        print(f'Não foi possivel enviar mensagem para {nome}')
        with open('erros.cvs','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')
    pyautogui.hotkey('ctrl','w')