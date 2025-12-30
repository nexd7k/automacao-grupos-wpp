import openpyxl
import pywhatkit as kit
from time import sleep
import pyautogui as pag
import os
import pyperclip

def msg():
    mensagem_def = f"Insira uma mensagem padrão e mude o conforme precise na planilha {mensagem}"

    #Tentativa de enviar mensagem para cliente
    try:
        print(f'Enviando mensagem para {nome}...')
        if mensagem is None:
            print(f"Não havia mensagem para enviar para {nome}")
        else:
            kit.sendwhatmsg_to_group_instantly(f'{id_grupo}', ' ')
            sleep(2)
            pyperclip.copy(mensagem_def)
            pag.hotkey('ctrl', 'v')
            pag.press('enter')
            print(f'Mensagem enviada para {nome}')
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{os.linesep}')
    sleep(3)
    with pag.hold('ctrl'):
        pag.press(['w'])


print("Iniciando o programa.")
# Ler planilha e guardar informações sobre nome, id do grupo e mensagem
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']
for linha in pagina_clientes.iter_rows(min_row=2):
    # id_grupo, nome, mensagem
    id_grupo = linha[0].value
    nome = linha[1].value
    mensagem = linha[2].value
    
    #Chamar a função da mensagem
    msg()

