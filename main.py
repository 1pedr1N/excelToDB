import pandas as pd
import mysql.connector



def conectar_mysql():
    try:
        conexao = mysql.connector.connect(
            host='localhost',
            user='root',
            password='passwordpassword',
            database='test'
        )
        return conexao
    except mysql.connector.Error as erro:
        print(f'Erro ao conectar ao MySQL: {erro}')



def ler_excel_e_enviar_para_mysql(arquivo_excel):
    df = pd.read_excel(arquivo_excel, sheet_name='CIDADE')
    py = df.keys()
    for i in py:
        print(i)

    conexao = conectar_mysql()


    try:
        cursor = conexao.cursor()

        for _, row in df.iterrows():
            coluna1 = row['Code']
            coluna2 = row['Name']
            coluna3 = row['StateId']

            sql = "INSERT INTO City (code, name, state_id) VALUES (%s, %s,%s)"  

            valores = (coluna1, coluna2, coluna3)
            print(valores)
            cursor.execute(sql, valores)

        conexao.commit()

        print('Dados inseridos com sucesso!')

    except mysql.connector.Error as erro:
        print(f'Erro ao inserir dados no MySQL: {erro}')
    finally:
        cursor.close()
        conexao.close()
        print('Conexão fechada.')
arquivo_excel = './seuexcelaqui.xlsx'
ler_excel_e_enviar_para_mysql(arquivo_excel)
