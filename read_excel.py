from pyodbc import connect

driver = "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
driver = "{" + driver + "}"

filename1 = "./Excel/Read 1.xlsx"
filename2 = "./Excel/Read 2.xls"
filename3 = "./Excel/Read 3.xlsx"

# ReadOnly=0 => Modo somente leitura
connection = f"DRIVER={driver};DBQ={filename3};ReadOnly=1"

with connect(connection, autocommit=True) as excel:
    cursor = excel.cursor()
    
    # Mostra tabelas/planilhas disponíveis
    for table in cursor.tables():
        print(table)
    
    # Realiza a query
    cursor.execute("SELECT * FROM [Produtos$]")
    
    # Resgata o nome das colunas
    col_names = [ column[0] for column in cursor.description ]
    
    for row in cursor:
        print(row)
        
        # Colunas com uma palavra de título
        print(row.Produtos, row.Valor, row.Desconto, row.Data)
        
        # Colunas com mais de uma palavra de título devem usar índice numérico
        print(row[4]) # Compras por cliques
        
        # Outro método para colunas com mais de uma palavra no título
        dict_row = dict(zip(col_names, row))
        
        print(dict_row["Compras por cliques"])
        
        
        
    