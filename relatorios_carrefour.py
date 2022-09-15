import PySimpleGUI as simplegui
import datetime
from dateutil.relativedelta import relativedelta

import atualizador_de_arquivo
import disparador_de_email
import verificador_pasta_mes_vigente

diretorio = r'\\SaoPaulo\Rezende\Jarezende\Administrativo\MIS\Planejamento_MIS\MIS\CARREFOUR'
dt = datetime.datetime.today()
data = dt.date() - relativedelta(months=1)

if dt.strftime('%d') == '1':
    pastaanexo = data.strftime('%Y%m')

else:
    pastaanexo = dt.strftime('%Y%m')
    
arquivo = ''
anexo = ''
arquivo_maluco = ''
data_maluca = data.strftime('%Y%m')

relatorios = {
    'TelefonesCarrefour':{
        'nome': 'Ludiana',
        'arquivo': f'telefones_carrefour_{pastaanexo}.xlsx',
        'arquivo_maluco': f'Telefones_Carrefour_{data_maluca}.xlsx',
        'destinatario': 'email1; email2',
        'copia': 'email3',
        'assunto': f'Telefones Carrefour {pastaanexo}',
        'corpo': f'base com telefones ativos do Carrefour referente ao período {pastaanexo}'        
        }
    }

def interface_grafica():
    simplegui.theme('DarkTeal')
    layout = [
        [simplegui.Text('Telefones Carrefour', size=(20, 1)),simplegui.Button('Atualizar', key='AttTelefonesCarrefour'),simplegui.Button('Enviar e-mail', key='EmailTelefonesCarrefour'), simplegui.Text('', key=('txt_atualizacao')), simplegui.Text('', key=('txt_email'))]
        ]
    
    janela = simplegui.Window('Relatórios Carrefour',layout, finalize=True)
     
    while True:
        eventos, valores = janela.read()
        if eventos in (None, 'Exit'):
            break
        if eventos == 'AttTelefonesCarrefour':
            arquivo = relatorios['TelefonesCarrefour']['arquivo']
            arquivo_maluco = relatorios['TelefonesCarrefour']['arquivo_maluco']
            anexo = r'{}\{}\{}'.format(diretorio, pastaanexo, arquivo)
            verificador_pasta_mes_vigente.verificador_de_pasta(diretorio, pastaanexo, arquivo, arquivo_maluco, anexo)
            tempo_de_execucao = (atualizador_de_arquivo.atualizar_relatorio(anexo))
            janela['txt_atualizacao'].update(f'Relatório atualizado em {tempo_de_execucao} segundos!')
        elif eventos == 'EmailTelefonesCarrefour':
            nome = relatorios['TelefonesCarrefour']['nome']
            arquivo = relatorios['TelefonesCarrefour']['arquivo']
            arquivo_maluco = relatorios['TelefonesCarrefour']['arquivo_maluco']
            anexo = r'{}\{}\{}'.format(diretorio, pastaanexo, arquivo)
            destinatario = relatorios['TelefonesCarrefour']['destinatario']
            copia = relatorios['TelefonesCarrefour']['copia']
            assunto = relatorios['TelefonesCarrefour']['assunto']
            corpo = relatorios['TelefonesCarrefour']['corpo']
            disparador_de_email.envia_email(nome, anexo, destinatario, copia, assunto, corpo)
            janela['txt_email'].update(f'E-mail {assunto} enviado com sucesso!')                 
        else:
            print('Não configurado') 
            
