from pyodbc import connect
from random import randrange, random
from datetime import datetime as dt

driver = "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
driver = "{" + driver + "}"

filename1 = "./Excel/Write 1.xlsx"
filename2 = "./Excel/Write 2.xls"
filename3 = "./Excel/Write 3.xlsx"

# ReadOnly=0 => Modo de escrita/leitura
connection = f"DRIVER={driver};DBQ={filename1};ReadOnly=0"

with connect(connection, autocommit=True) as excel:
    cursor = excel.cursor()
    
    # Insere linhas na planilha
    linhas = [
        [randrange(6, 1000), "Meu novo produto", 1000, random(), dt.today(), random()]
    ]
    
    for linha in linhas:
        cursor.execute("INSERT INTO [Produtos$] (Indice, Produtos, Valor, Desconto,	Data, Compras) VALUES(?, ?, ?, ?, ?, ?)", *linha)
    
    # Atualiza uma linha
    cursor.execute(f"UPDATE [Produtos$] SET Valor = ? WHERE Indice = ?", randrange(2000, 2500), 1)
    
    # Adiciona uma coluna : Não é possível
    # cursor.execute(f"ALTER TABLE [Produtos$] ADD COLUMN 'Disponivel' TEXT")
    # cursor.execute(f"UPDATE [Produtos$] SET Disponivel = ? WHERE Indice = ?", "SIM", 2)
    
    # Limpa uma célula
    cursor.execute(f"UPDATE [Produtos$] SET Valor = ? WHERE Indice = ?", None, 5)
    
    # Exclui uma linha
    clear = [ None for number in range(0, 6) ]
    cursor.execute(f"UPDATE [Produtos$] SET Indice = ?, Produtos = ?, Valor = ?, Desconto = ?,	Data = ?, Compras = ? WHERE Indice = ?", *clear, 4)
    