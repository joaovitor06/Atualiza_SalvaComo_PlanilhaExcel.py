import shutil
import datetime

def pega_arquivo_mes_anterior(diretorio, arquivo, arquivo_maluco, anexo, pastaanexo):
    today = datetime.date.today()
    first = today.replace(day=1)
    lastMonth = first - datetime.timedelta(days=1)
    final = lastMonth.strftime("%Y%m")
    #shutil.copyfile(f'{diretorio}\{final}\{arquivo_maluco}', anexo)
    shutil.copyfile(f'{diretorio}\{final}\{arquivo_maluco}', f'{diretorio}\{pastaanexo}\{arquivo}')
    print(f'{arquivo} copiado para {anexo}')   


