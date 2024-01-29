import openpyxl
from urllib.parse import quote
from time import sleep
import pyautogui
import webbrowser
from datetime import datetime, timedelta


webbrowser.open('https://web.whatsapp.com/')
sleep(30)
pyautogui.hotkey('ctrl','w')
workbook = openpyxl.load_workbook('teste1.xlsx')
pag_clientes = workbook['Plan1']
pix = 'pix celular 92991200836'
hoje = datetime.now()


for linha in pag_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha [1].value
    vencimento = linha [2].value
    aviso = vencimento - timedelta(days=2)
    if aviso <= hoje:
        mensagem = f'Olá {nome}, estou passando para informar que sua lista vence dia {vencimento.strftime('%d/%m/%Y')} para não ter interrupções no serviço faça a renovação com o {pix}'
    
        try:
            link_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_whatsapp)
            sleep(10)
            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(5)
            pyautogui.click(seta[0],seta[1])
            sleep(5)

        except:
            print(f'não foi possivel enviar mensagem para {nome}')
            with open('erros.cvs','a',newline='',encoding='uft-8') as arquivo:
                arquivo.write(f'{nome},{telefone}')
    else:
        pass
    pyautogui.hotkey('ctrl','w')