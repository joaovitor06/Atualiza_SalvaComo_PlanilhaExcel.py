import win32com
import datetime

#Formata a data de hoje para anomêsdia
data = datetime.datetime.today().strftime('%Y%m%d')

#Define o produto que será utilizado
xlapp = win32com.client.DispatchEx("Excel.Application")

#Busca arquivo no diretório
wb = xlapp.Workbooks.Open("J:\\Comercial\\MIS\\Yamaha\\BaseYAMAHAWEDOO_2022802.xlsm")
#Executa a atualização dos dados
wb.RefreshAll()
#Aguarda a conclusão da atualização para prosseguir
xlapp.CalculateUntilAsyncQueriesDone()
#Salva um novo arquivo, alterando a data do nome para o dia de hoje
wb.SaveAs("J:\\Comercial\\MIS\\Yamaha\\BaseYAMAHAWEDOO_{}.xlsm".format(data))
#Fecha a planilha 
map(lambda book: book.Close(False), xlapp.Workbooks)
xlapp.Quit()

