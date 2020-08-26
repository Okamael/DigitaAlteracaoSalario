from tir import Webapp
from tir import ApwInternal
from datetime  import datetime
import pandas  as pd
import locale as lo

filial="01010001"
dataalt="01/08/2020"
tipoalt="007"
dataAdimissao = ""


lo.setlocale(lo.LC_MONETARY,'pt_BR.UTF-8')
DataSytem = datetime.today().strftime('%d/%m/%Y')
oHelper = Webapp("C:\\Users\\TI\\Documents\\branche04\\.teste\\dissidio_automatizacao\\config.json")

Listagem= pd.read_excel("C:\\Users\\TI\\Documents\\branche04\\.teste\\dissidio_automatizacao\\dissidio.xlsx")

newListagem = Listagem[['Matricula','correcao']]


'''Entra na tela de funcionarios'''
oHelper.Setup("SIGAGPE",DataSytem,"10","01010001","07")
oHelper.SetBranch
oHelper.Program("GPEA010")

for x in range (len(newListagem)):
    matricula = int(newListagem.values[x][0])
    salarioNaoFormatado = newListagem.values[x][1]
    salarioFormatado = lo.currency(salarioNaoFormatado)
    print(f'Matricula:{matricula}, Salario:{salarioFormatado}')


    '''Realiza a alteração de dissidio'''
    oHelper.SearchBrowse(f'{filial}{matricula}',"Filial+matricula")
    oHelper.SetButton("Alterar")
    oHelper.ClickFolder("Funcionais")
    '''Valida data de  adimissao do  funcionario'''
    oHelper.SetValue("RA_SALARIO",salarioFormatado,check_value=False)
    oHelper.SetValue("RA_ANTEAUM",salarioFormatado,check_value=False)
    oHelper.SetValue("RA_DATAALT",dataalt, check_value=False)
    oHelper.SetValue("RA_TIPOALT",tipoalt,check_value=False)
    oHelper.SetButton("Salvar")
    oHelper.SetButton("OK")