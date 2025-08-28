import os
import win32com.client as win32

pasta = r"F:\Direcção Comercial\Tratamento Telecom\Doc Pagamento\Inicial"


excel = win32.Dispatch("Excel.Application")
excel.DisplayAlerts = False  # evita mensagens do Excel

for arquivo in os.listdir(pasta):
    if arquivo.lower().endswith(".xls") and not arquivo.lower().endswith(".xlsx"):
        caminho_arquivo = os.path.join(pasta, arquivo)
        workbook = excel.Workbooks.Open(caminho_arquivo)
        
        # Novo nome .xlsx
        novo_nome = arquivo.rsplit(".", 1)[0] + ".xlsx"
        caminho_novo = os.path.join(pasta, novo_nome)
        
        workbook.SaveAs(caminho_novo, FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
        workbook.Close()

excel.Quit()
print("Conversão concluída!")
