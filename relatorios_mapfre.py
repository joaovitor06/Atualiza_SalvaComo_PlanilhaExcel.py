import PySimpleGUI as simplegui
import datetime
from dateutil.relativedelta import relativedelta

import atualizador_de_arquivo
import disparador_de_email
import verificador_pasta_mes_vigente

diretorio = r'\\SaoPaulo\Rezende\Jarezende\Administrativo\MIS\Planejamento_MIS\MIS\MAPFRE'
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
    'AcionamentosMapfreInterno':{
        'nome': 'Prezados',
        'arquivo': f'ACIONAMENTOS_MAPFRE_{pastaanexo} - Interno.xlsb',
        'arquivo_maluco': f'ACIONAMENTOS_MAPFRE_{data_maluca} - Interno.xlsb',
        'destinatario': 'joao.santos@jarezende.com.br', 
        'copia': 'joao.santos@jarezende.com.br', 
        'assunto': f'Acionamentos MAPFRE INTERNO {pastaanexo}',
        'corpo': f'Acionamentos MAPFRE referente ao período {pastaanexo}'        
        }
    }

def interface_grafica():
    simplegui.theme('DarkRed1')
    layout = [
        [simplegui.Text('Acionamentos Mapfre INTERNO (Enviar às 16:30)'),simplegui.Button('Atualizar', key='AttAcionamentosMapfreInterno'),simplegui.Button('Enviar e-mail', key='EmailAcionamentosMapfreInterno'), simplegui.Text('', key=('txt_atualizacao')), simplegui.Text('', key=('txt_email'))]
        ]
    
    janela = simplegui.Window('Relatórios Mapfre',layout, finalize=True)
     
    while True:
        eventos, valores = janela.read()
        if eventos in (None, 'Exit'):
            break
        if eventos == 'AttAcionamentosMapfreInterno':
            arquivo = relatorios['AcionamentosMapfreInterno']['arquivo']
            arquivo_maluco = relatorios['AcionamentosMapfreInterno']['arquivo_maluco']
            anexo = r'{}\{}\{}'.format(diretorio, pastaanexo, arquivo)
            verificador_pasta_mes_vigente.verificador_de_pasta(diretorio, pastaanexo, arquivo, arquivo_maluco, anexo)
            tempo_de_execucao = (atualizador_de_arquivo.atualizar_relatorio(anexo))
            janela['txt_atualizacao'].update(f'Relatório atualizado em {tempo_de_execucao} segundos!')
        elif eventos == 'EmailAcionamentosMapfreInterno':
            nome = relatorios['AcionamentosMapfreInterno']['nome']
            arquivo = relatorios['AcionamentosMapfreInterno']['arquivo']
            arquivo_maluco = relatorios['AcionamentosMapfreInterno']['arquivo_maluco']
            anexo = r'{}\{}\{}'.format(diretorio, pastaanexo, arquivo)
            destinatario = relatorios['AcionamentosMapfreInterno']['destinatario']
            copia = relatorios['AcionamentosMapfreInterno']['copia']
            assunto = relatorios['AcionamentosMapfreInterno']['assunto']
            corpo = relatorios['AcionamentosMapfreInterno']['corpo']
            disparador_de_email.envia_email(nome, anexo, destinatario, copia, assunto, corpo)
            janela['txt_email'].update(f'E-mail {assunto} enviado com sucesso!')                 
        else:
            print('Não configurado') 
            
