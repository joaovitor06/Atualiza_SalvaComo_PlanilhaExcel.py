import win32com.client as win32
import PySimpleGUI as simplegui
import relatorios_carrefour
import relatorios_mapfre
import relatorios_unisa

outlook = win32.Dispatch('Outlook.application')
excel = win32.Dispatch('Excel.Application')
nome = r''
anexo = r''
destinatario = r''
copia = r''
assunto = r''
corpo = r''   

simplegui.theme('DarkBlue')
layout = [
    [simplegui.Text('Relatórios Carrefour', size=(20, 1)),simplegui.Button('Acessar', key='acessar_relatorios_carrefour')],
    [simplegui.Text('Relatórios Mapfre', size=(20, 1)),simplegui.Button('Acessar', key='acessar_relatorios_mapfre')],
    [simplegui.Text('Relatórios UNISA', size=(20, 1)),simplegui.Button('Acessar', key='acessar_relatorios_unisa')]
    ]

janela = simplegui.Window('Menu',layout, finalize=True)    
while True:
    eventos, valores = janela.read()
    if eventos in (None, 'Exit'):
        break
    if eventos == 'acessar_relatorios_carrefour':
        janela.hide()
        relatorios_carrefour.interface_grafica()
        janela.un_hide()
    elif eventos == 'acessar_relatorios_mapfre':
        janela.hide()
        relatorios_mapfre.interface_grafica()
        janela.un_hide()
    elif eventos == 'acessar_relatorios_unisa':
        janela.hide()
        relatorios_unisa.interface_grafica()
        janela.un_hide()
    else:
        print('Não configurado')        

