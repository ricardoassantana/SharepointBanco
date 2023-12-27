# %%
import pandas as pd
from sqlalchemy import create_engine
import urllib
from time import sleep, time
import os
import win32com.client
from datetime import datetime


# %%
pasta = r'C:User/SeusDocumentos/PastacomExcel' #Adicione o caminho onde estão localizados os arquivos em Excel

# %%
""" ------------------------------------ PROCESSO DE ATUALIZAÇÃO DOS SHAREPOINTS POR MEIO DOS ARQUIVOS EM EXCEL ------------------------------------"""

max_tentativas = 3 


arquivos_excel = [arquivo for arquivo in os.listdir(pasta) if arquivo.endswith('.xlsx')and not arquivo.startswith('~$')]

for arquivo_excel in arquivos_excel:
    caminho_arquivo = os.path.join(pasta, arquivo_excel)
    tentativas = 0

    print(f"Lendo o arquivo: {caminho_arquivo}")

    while tentativas < max_tentativas:
        tentativas += 1

        try:
            excel = None
            workbook = None
            excel = win32com.client.Dispatch("Excel.Application")
            sleep(2)
            excel.Visible = False
            workbook = excel.Workbooks.Open(caminho_arquivo)
            start_time = time()
            break # Se der tudo certo, saia do loop de tentativas
            print("Atualização do excel via sharepoint...")
        except Exception as erro:
            print(f"Erro ao abrir o arquivo: {caminho_arquivo}")
            print(f"Tentativa {tentativas}/{max_tentativas}")
            print(f"O erro: {erro}")
            sleep(3)  
            if excel is not None:
                del excel

    if tentativas == max_tentativas:
        print(f"Não foi possível abrir o arquivo após {max_tentativas} tentativas. Pulando para o próximo.")
        continue

    NomeConexaoExcel = []
    for connection in workbook.Connections:
        NomeConexaoExcel.append(connection.Name)

    """ ------------------------------------ PROCESSO DE ATUALIZAÇÃO DAS CONEXÕES COM O SHAREPOINT ------------------------------------"""

    for nome_conexao in NomeConexaoExcel:
        print(f"Atualização da conexão: {nome_conexao}")
        start_time = time()
        print("Atualização do excel via sharepoint...")
        connection = workbook.Connections[nome_conexao]
        connection.OLEDBConnection.BackgroundQuery = False
        connection.Refresh()
        end_time = time()
        duracao_tempo = end_time - start_time
        print("Tempo de atualização: ", duracao_tempo)
        print("Fim da atualização")
        sleep(3)

    # workbook.Save()
    # print("Excel salvo")
    workbook.Close(SaveChanges=True)
    print("Excel salvo e fechado")
    excel.Quit()
    print("Aplicativo Excel fechado")
    del excel
    sleep(2)

print("App Excel fechado!")
   
""" ------------------------------------ FIM DO PROCESSO DE ATUALIZAÇÃO DOS SHAREPOINTS POR MEIO DOS ARQUIVOS EM EXCEL ------------------------------------"""    


# %%

""" ------------------------------------ PROCESSO DE CONEXÃO COM O BANCO DE DADOS ------------------------------------""" 

params = urllib.parse.quote_plus("DRIVER={ODBC Driver 17 for SQL Server};"
                                 "SERVER=NOME DO SERVIDOR;" # Adicione o seu servidor
                                 "DATABASE=NOME DO BANCO;" # Adicione o nome do banco de dados
                                 "Trusted_Connection=yes")

engine = create_engine("mssql+pyodbc:///?odbc_connect={}".format(params), fast_executemany=True)

# %%

""" ------------------------------------ FUNÇÃO QUE ADICIONA UMA TABELA NO BANCO DE DADOS ------------------------------------""" 
def Inserir_Tabela_Banco(df, nome_tabela, engine):
    df.to_sql(name=nome_tabela, con=engine, schema='dbo', index=False, if_exists='append')

def conectar_banco():
    conn = engine.connect()
    trans = conn.begin()
    print("Banco conectado com sucesso!") 
    return conn, trans

def desconectar_banco(conn):
    conn.close()
    print('Conexão com o banco fechada')

# Função abaixo para escolher o nome da tabela que será inserida no banco de dados + o nome do arquivo excel
def nome_da_tabela(arquivo_excel):
    nome_tabela = "NOME_PADRAO_DA_TABELA_" + arquivo_excel.replace(".xlsx","")
    print(nome_tabela)
    return nome_tabela
    
def excluir_linhas_tabela(arquivo_excel, conn, nome_tabela):
    print(f'Inserindo no banco a planilha {arquivo_excel}...')
    VerificaLinhasTabela = f"SELECT COUNT(*) FROM {nome_tabela}"
    result = conn.execute(VerificaLinhasTabela)
    contador = result.fetchone()[0]
    if contador > 0:
            print(f"A tabela {nome_tabela} possui", contador, "linhas.")
            print(f"{nome_tabela}: Executando o comando de exclusão...")
            # Executa o comando de exclusão
            DeletaLinhasTabela = f"DELETE FROM {nome_tabela}"
            result = conn.execute(DeletaLinhasTabela)
            print(f"{nome_tabela}: Linhas da tabela deletadas com sucesso!")
            VerificaLinhasDeNovo = conn.execute(VerificaLinhasTabela)
            
            if VerificaLinhasDeNovo.rowcount > 0:
                print(f'{nome_tabela}:Nada foi deletado do banco')
                print(VerificaLinhasDeNovo.rowcount)
            else:
                print(f"{nome_tabela}: Linhas deletadas do banco")
            # Verifica o número de linhas afetadas
            linhas_afetadas = result.rowcount
            
            if linhas_afetadas > 0:
                print(f"{nome_tabela}: Linhas da tabela deletadas com sucesso. Total de linhas afetadas:", linhas_afetadas)
            else:
                print(f"{nome_tabela}: Nenhuma linha foi deletada.")
    else:
        print(f"{nome_tabela}: A tabela está vazia. Nenhum comando de exclusão necessário.")
    

def planilha_duplicatas_datacarga(caminho_arquivo, nome_tabela, arquivo_excel):
    df = pd.read_excel(caminho_arquivo)
    print(f"{nome_tabela}: Planilha importada para o Python!")
    #removendo as duplicidades da base 
    df = df.drop_duplicates()
    print(f"{nome_tabela}: Duplicatas removidas")
    # Adiciona uma coluna de Data de Carga para validação da importação
    df['DataCarga'] = datetime.today()
    print(f"{nome_tabela}: Data de carga inserida!")
    print(f"{nome_tabela}: Inserindo no banco a planilha...")
    return df


# %%
""" ------------------------------------ PROCESSO DE LEITURA DOS ARQUIVOS QUE ESTÃO NO CAMINHO OFICIAL DAS BASES ------------------------------------""" 

arquivos_excel = [arquivo for arquivo in os.listdir(pasta) if arquivo.endswith('.xlsx')and not arquivo.startswith('~$')]
nome_tabela_banco = []
for arquivo_excel in arquivos_excel:
    caminho_arquivo = os.path.join(pasta, arquivo_excel)
    nome_tabela = nome_da_tabela(arquivo_excel)
    nome_tabela_banco.append(nome_tabela)

    # ----------------- Conectando com o banco de dados -----------------
    conn, trans = conectar_banco()
    excluir_linhas_tabela(arquivo_excel, conn, nome_tabela)
    trans.commit()
    print('Realizado o commit para o banco')   
    desconectar_banco(conn)
    sleep(1)
    print('Conexão fechada com o banco!')
    # ----------------- Conectando com o banco de dados -----------------
    conn, trans = conectar_banco()
    # Removendo as duplicidades da base
    df = planilha_duplicatas_datacarga(caminho_arquivo,nome_tabela,arquivo_excel)
    # Inserindo tabela no banco
    Inserir_Tabela_Banco(df, nome_tabela, engine)

    print(f'Pronto. Importação da tabela {nome_tabela} realizada com sucesso!')
    desconectar_banco(conn)     

print('Todas as planilhas do sharepoint foram importadas com sucesso para o banco!')


