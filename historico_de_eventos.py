import pyodbc

def importa_historico(anexo, tipo):
    server = '192.168.200.182\DB2014'
    database = 'MIS'
    username = 'MIS'
    password = 'm1s'
    cnn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnn.cursor()

    sql = f"INSERT INTO MIS.DBO.HISTORICO_MENU_JP SELECT '{anexo}', GETDATE(), '{tipo}'"

    cursor.execute(sql)
    cursor.close()
    cnn.commit()

