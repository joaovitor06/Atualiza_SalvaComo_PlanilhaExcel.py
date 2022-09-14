import PySimpleGUI as simplegui
import datetime
from dateutil.relativedelta import relativedelta

import atualizador_de_arquivo
import disparador_de_email
import verificador_pasta_mes_vigente

diretorio = r'\\SaoPaulo\Rezende\Jarezende\Administrativo\MIS\Planejamento_MIS\MIS\UNISA'
dt = datetime.datetime.today()
data = dt.date() - relativedelta(months=1)

if dt.strftime('%d') == '01':
    pastaanexo = data.strftime('%Y%m')

else:
    pastaanexo = dt.strftime('%Y%m')
    
arquivo = ''
anexo = ''
arquivo_maluco = ''
data_maluca = data.strftime('%Y%m')

relatorios = {
    'ComparativoUnisa':{
        'nome': 'Alzira',
        'arquivo': f'COMPARATIVO_UNISA_{pastaanexo}.xlsm',
        'arquivo_maluco': f'COMPARATIVO_UNISA_{data_maluca}.xlsm',
        'destinatario': 'email1; email2; email3', 
        'copia': 'email4; email5; email6', 
        'assunto': f'Comparativo UNISA {pastaanexo}',
        'corpo': 'Comparativo UNISA'        
        },
        'ComparativoUnisaFechamento':{
            'nome': 'Alzira',
            'arquivo': f'COMPARATIVO_UNISA_FECHAMENTO_{pastaanexo}.xlsm',
            'arquivo_maluco': f'COMPARATIVO_UNISA_FECHAMENTO_{data_maluca}.xlsm',
            'destinatario': 'email1; email2; email3', 
            'copia': 'email4; email5; email6', 
            'assunto': f'Comparativo UNISA FECHAMENTO {pastaanexo}',
            'corpo': 'Comparativo UNISA FECHAMENTO'        
            }
    }

def interface_grafica():
    simplegui.theme('DarkBlue12')
    layout = [
        [simplegui.Text('Comparativo UNISA (Sexta-feira)'),simplegui.Button('Atualizar', key='AttComparativoUnisa'),simplegui.Button('Enviar e-mail', key='EmailComparativoUnisa'), simplegui.Text('', key=('txt_atualizacao_ComparativoUnisa')), simplegui.Text('', key=('txt_email_ComparativoUnisa'))],
        [simplegui.Text('Comparativo UNISA FECHAMENTO (primeira Sexta-feira do mês)'),simplegui.Button('Atualizar', key='AttComparativoUnisaFechamento'),simplegui.Button('Enviar e-mail', key='AttComparativoUnisaFechamento'), simplegui.Text('', key=('txt_atualizacao_ComparativoUnisaFechamento')), simplegui.Text('', key=('txt_email_ComparativoUnisaFechamento'))]
        ]
    
    janela = simplegui.Window('Relatórios UNISA',layout, finalize=True)
     
    while True:
        eventos, valores = janela.read()
        if eventos in (None, 'Exit'):
            break
        if eventos == 'AttComparativoUnisa':
            arquivo = relatorios['ComparativoUnisa']['arquivo']
            arquivo_maluco = relatorios['ComparativoUnisa']['arquivo_maluco']
            anexo = r'{}\{}\{}'.format(diretorio, pastaanexo, arquivo)
            verificador_pasta_mes_vigente.verificador_de_pasta(diretorio, pastaanexo, arquivo, arquivo_maluco, anexo)
            tempo_de_execucao = (atualizador_de_arquivo.atualizar_relatorio(anexo))
            janela['txt_atualizacao_ComparativoUnisa'].update(f'Relatório atualizado em {tempo_de_execucao} segundos!')
        elif eventos == 'EmailComparativoUnisa':
            nome = relatorios['ComparativoUnisa']['nome']
            arquivo = relatorios['ComparativoUnisa']['arquivo']
            arquivo_maluco = relatorios['ComparativoUnisa']['arquivo_maluco']
            anexo = r'{}\{}\{}'.format(diretorio, pastaanexo, arquivo)
            destinatario = relatorios['ComparativoUnisa']['destinatario']
            copia = relatorios['ComparativoUnisa']['copia']
            assunto = relatorios['ComparativoUnisa']['assunto']
            corpo = relatorios['ComparativoUnisa']['corpo']
            disparador_de_email.envia_email(nome, anexo, destinatario, copia, assunto, corpo)
            janela['txt_email_ComparativoUnisa'].update(f'E-mail {assunto} enviado com sucesso!')
        elif eventos == 'AttComparativoUnisaFechamento':
            arquivo = relatorios['ComparativoUnisaFechamento']['arquivo']
            arquivo_maluco = relatorios['ComparativoUnisaFechamento']['arquivo_maluco']
            anexo = r'{}\{}\{}'.format(diretorio, pastaanexo, arquivo)
            verificador_pasta_mes_vigente.verificador_de_pasta(diretorio, pastaanexo, arquivo, arquivo_maluco, anexo)
            tempo_de_execucao = (atualizador_de_arquivo.atualizar_relatorio(anexo))
            janela['txt_atualizacao_ComparativoUnisaFechamento'].update(f'Relatório atualizado em {tempo_de_execucao} segundos!')
        elif eventos == 'AttComparativoUnisaFechamento':
            nome = relatorios['ComparativoUnisaFechamento']['nome']
            arquivo = relatorios['ComparativoUnisaFechamento']['arquivo']
            arquivo_maluco = relatorios['ComparativoUnisaFechamento']['arquivo_maluco']
            anexo = r'{}\{}\{}'.format(diretorio, pastaanexo, arquivo)
            destinatario = relatorios['ComparativoUnisaFechamento']['destinatario']
            copia = relatorios['ComparativoUnisaFechamento']['copia']
            assunto = relatorios['ComparativoUnisaFechamento']['assunto']
            corpo = relatorios['ComparativoUnisaFechamento']['corpo']
            disparador_de_email.envia_email(nome, anexo, destinatario, copia, assunto, corpo)
            janela['txt_email_ComparativoUnisaFechamento'].update(f'E-mail {assunto} enviado com sucesso!')                          
        else:
            print('Não configurado') 

