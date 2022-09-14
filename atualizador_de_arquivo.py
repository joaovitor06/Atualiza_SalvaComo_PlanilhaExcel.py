import win32com.client as win32
import time
import historico_de_eventos

excel = win32.Dispatch('Excel.Application')
tipo = 'atualizacao'

def atualizar_relatorio(anexo):
    inicial = time.time()
    wb = excel.Workbooks.Open(anexo)
    wb.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()
    wb.Save()
    wb.Close(SaveChanges=False)
    excel.Quit()
    final = time.time()
    tempo_de_execucao = int(final - inicial)
    historico_de_eventos.importa_historico(anexo, tipo)
    return tempo_de_execucao

    