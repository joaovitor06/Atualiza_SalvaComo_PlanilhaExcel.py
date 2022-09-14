import pyodbc

def importa_historico(anexo, tipo):
    server = 'SERVIDOR'
    database = 'BANCO DE DADOS'
    username = 'USUARIO'
    password = 'SENHA'
    cnn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnn.cursor()

    sql = f"INSERT INTO MIS.DBO.HISTORICO_MENU_JP SELECT '{anexo}', GETDATE(), '{tipo}'"

    cursor.execute(sql)
    cursor.close()
    cnn.commit()

