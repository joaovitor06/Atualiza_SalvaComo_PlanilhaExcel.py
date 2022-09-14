
import os
import pega_arquivo

def verificador_de_pasta(diretorio, pastaanexo, arquivo, arquivo_maluco, anexo):
        
    if os.path.exists(f'{diretorio}\{pastaanexo}'):
        print('Pasta já existe!')
        
        if os.path.exists(f'{diretorio}\{pastaanexo}\{arquivo}'):
            print('Arquivo já existe!')
            
        else:
            pega_arquivo.pega_arquivo_mes_anterior(diretorio, arquivo, arquivo_maluco, anexo, pastaanexo)
            print(f'Arquivo {arquivo} copiado para a pasta!')
    else:
        os.mkdir(f"{diretorio}\{pastaanexo}")
        print(f"{diretorio}\{pastaanexo} criada")
        pega_arquivo.pega_arquivo_mes_anterior(diretorio, arquivo, arquivo_maluco, anexo, pastaanexo)
        print(f'Arquivo {arquivo} copiado para a pasta!')
