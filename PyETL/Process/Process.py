import xlrd
import os
from Model import Models

class Read: 
    def LoadFiles(Local):
        GlobalFiles = [os.path.join(Local,nome) for nome in os.listdir(Local)]
        Files = [arq for arq in GlobalFiles if os.path.isfile(arq)]
        jpgs = [arq for arq in Files if arq.lower().endswith(".xls")]
        return jpgs

    def ReadFiles():

        return

class AutoXrdl:
    def Readxls(Filepath):
        try:
            #Tabela Instancia 
            Cliente = Models.Cliente 
            Funcionario = Models.Funcionario
            OS = Models.Os
            Produtos_OS = [] 
            Veiculo  =  Models.Veiculo

            #Carregando XLSX
            book = xlrd.open_workbook(Filepath) #("Filepath.xls")
            sh = book.sheet_by_index(1)
            print("Planilha {0}".format(sh.name))
        #print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
        #print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))

        #Carga
        #print("Nome: {0}".format(sh.cell_value(rowx=11,colx=4)))
        #print(sh.row(11))
            ex = sh.row(11)

        #Cliente
            Cliente = Cliente(
                0,
                sh.cell_value(rowx=11,colx=4),
                sh.cell_value(rowx=14,colx=4),
                '')
        
        #Veiculo    
            Veiculo = Veiculo(
                0,
                "Carro",
                sh.cell_value(rowx=12,colx=11),
                sh.cell_value(rowx=13,colx=4),
                sh.cell_value(rowx=13,colx=11),
                sh.cell_value(rowx=14,colx=11),
                sh.cell_value(rowx=12,colx=4))

        #Veiculo().Placa = sh.cell_value(rowx=13,colx=12)
        #Veiculo().Modelo = sh.cell_value(rowx=12,colx=12)
        #Veiculo().Ano = sh.cell_value(rowx=12,colx=4)
        #Veiculo().Km = sh.cell_value(rowx=14,colx=12)
        #Veiculo().Motor = sh.cell_value(rowx=13,colx=4)

            #OS
            OS(
                0,
                0,
                0,
                0,
                sh.cell_value(rowx=11,colx=4),
                0,
                '',
                sh.cell_value(rowx=11,colx=4))

        #OS().DataEntrada = sh.cell_value(rowx=11,colx=4)
        #OS().MaoObra = sh.cell_value(rowx=11,colx=4)

        #Produtos Lista consumo 
            i = 17
            while (i < 48):
                Produtos = Models.Produtos
                Produtos = Produtos(
                    0,
                    0,
                    sh.cell_value(rowx=i,colx=4),
                    0,
                    0,
                    sh.cell_value(rowx=i,colx=3),
                    sh.cell_value(rowx=i,colx=10))
                
                #Produtos().Descricao = sh.cell_value(rowx=i,colx=5)
                #Produtos().Valor = sh.cell_value(rowx=i,colx=11)
                #Produtos().QTD = sh.cell_value(rowx=i,colx=4)
                if Produtos.Valor != "" and Produtos.Descricao != "":
                    Produtos_OS.append(Produtos)
                i += 1

        #Funcionario
            Funcionario = Funcionario(0,sh.cell_value(rowx=49,colx=9),"","Mecanico","")
        
            Dados = [True,Cliente,Funcionario,Veiculo,OS,Produtos_OS]


            return Dados
        except Exception as Erro:
            return [False,Erro]
