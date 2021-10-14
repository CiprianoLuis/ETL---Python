import pyodbc
#from xlrd.compdoc import 

server = 'tcp:kikosandb.cajpqxmeoicq.us-east-1.rds.amazonaws.com' 
database = 'Oficina_Mecanica' 
username = 'Admin' 
password = 'FourData' 

STRConect = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password

class DataAcess:
    """def OpnConect(_init_): 
        cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+_init_.server+';DATABASE='+_init_.database+';UID='+_init_.username+';PWD='+ _init_.password)
        cursor = cnxn.cursor()
        return cursor"""

    def OpnConect2(conect): 
        #Instance
        cnxn = pyodbc.connect(conect)
        cursor = cnxn.cursor()
        return cursor
    def CommitBD(conect):
        cnxn = pyodbc.connect(conect)
        cnxn.commit()
        return 

# Ações
class DataFunctions:
    def RegisterCar(Veiculo):
        try:
            cursor = DataAcess.OpnConect2(STRConect)
            query =   """
                INSERT INTO Veiculo (Marca, Modelo, Motor, Placa,Km,Ano)
                      VALUES ('{}','{}','{}','{}','{}','{}')
                """.format(Veiculo.Marca,Veiculo.Modelo,Veiculo.Motor,Veiculo.Placa,Veiculo.Km,Veiculo.Ano)
            cursor.execute(query) 
            cursor.commit()
            return True
        except Exception as Error:
            print(Error)
            return Error
    def SearchCar():
        return

    def RegisterClient(Cliente):
        try:
            cursor = DataAcess.OpnConect2(STRConect)
            #Select Query
            stringquery =  """
                    INSERT INTO _Cliente (Nome,Telefone,Email)
                        VALUES ('{}','{}','{}')
                """.format(Cliente.Nome,Cliente.Telefone,Cliente.Email)

            cursor.execute(stringquery) 
            cursor.commit()
            return True
        except Exception as Error:
            print(Error)
            return Error
    def SearchClient():
        return

    def RegisterOS():
        try:
            cursor = DataAcess.OpnConect2(STRConect)
            #Select Query
            cursor.execute(
                """
                    INSERT INTO OS (ID_Funcinario_FK, ID_Prod_FK, ID_Clinte_FK, ID_Veiculo_FK, Data_Entrada, Pago, Observacao, Mao_De_Obra)
                        VALUES ({},{},{},{},{},{},{},{})
                """
            ) 
            row = cursor.fetchone() 
            return True
        except Exception as Error:
            print(Error)
            return Error
    def SearchOS():
        return

    def RegisterItem(Produto):
        try:
            cursor = DataAcess.OpnConect2(STRConect)
            #Select Query
            stringquery = """
                    INSERT INTO Produtos (Descricao, Produtos_Cliente, Produtos_Loja,QTD, Valor)
                        VALUES ('{}',{},{},{},{})
                """.format(Produto.Descricao,Produto.Produtos_Cliente,Produto.Produtos_Loja,Produto.QTD,Produto.Valor)

            cursor.execute(stringquery) 
            cursor.commit()
            return True
        except Exception as Error:
            print(Error)
            return Error   
    def SearchItem():
        return



    def RegisterFalhaLeitura(Erro):
        cursor = DataAcess.OpnConect2(STRConect)
        #Select Query
        cursor.execute(
                """
                    Insert Historico_ETL (ID) 
                    VALUES (?,?,?,?,?)
                """
            ) 
        row = cursor.fetchone() 
        return True
    def SearchFalhaLeitura():
        return

    #Metodo de Cadastramento 
    def Inserir(Dados):
        try:
            #AtuantesFizicos
            _Cliente = DataFunctions.RegisterClient(Dados[1])
            _Veiculo = DataFunctions.RegisterCar(Dados[3])
            
            #ProdutosServiço
            for Produto in Dados[5]:
                _Produto = DataFunctions.RegisterItem(Produto)
                
            #RegistrarOS
            _OS = DataFunctions.RegisterOS(Dados[4])

            return('Rows inserted: ')
        
        except Exception as Erro:
            return [False,Erro]


