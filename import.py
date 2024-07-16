import pandas as pd
import pyodbc

def load_excel_data(file_path):
    df = pd.read_excel(file_path)
    return df

# Função para conectar ao banco de dados SQL Server usando pyodbc
def connect_to_db():
    server = 'SERVDB4\\SQLEXPRESS'
    database = 'futebol'
    username = 'sa'
    password = ''
    driver = 'ODBC Driver 17 for SQL Server' ##Atualizar drive conforme o SQL instalado na maquina
    conn_str = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'
    conn = pyodbc.connect(conn_str)
    return conn

def upsert_data(conn, df):
    cursor = conn.cursor()

    for index, row in df.iterrows():
        team = row['team']
        liga = row['liga']
        country = row['country']

        cursor.execute(f"SELECT COUNT(*) FROM clube WHERE team = ?", team)
        if cursor.fetchone()[0] > 0:
            # Atualiza o registro
            cursor.execute(f"UPDATE clube SET liga = ?, country = ? WHERE team = ?", liga, country, team)
        else:
            # Insere um novo registro
            cursor.execute(f"INSERT INTO clube (team, liga, country) VALUES (?, ?, ?)", team, liga, country)

    conn.commit()

file_path = 'C:\\Users\\thorodin\\Desktop\\share\\base.xls'

try:
    df = load_excel_data(file_path)
    conn = connect_to_db()
    upsert_data(conn, df)
    print("Dados atualizados com sucesso!")

except Exception as e:
    print(f"Erro ao processar dados: {e}")
