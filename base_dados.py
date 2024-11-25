import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine

def get_bank_transactions(conta, mes, ano=2024):
    # Configurações de conexão
    server = r'phc\sqlexpress'
    database = 'BDRECACTIV'
    username = 'Tiago'
    password = 'Catarina03'

    try:
        # Criar engine SQLAlchemy
        engine = create_engine(f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server')

        # Executar query
        query = f"""SELECT
        CONVERT(VARCHAR(10), data, 103) AS 'Data'
        , ml.dinome AS 'Diário'
        , ml.dilno AS 'Nº'
        , ml.adoc AS 'Documento'
        , ml.descritivo AS 'Descrição'
        , ml.edeb AS 'Débito'
        , ml.ecre AS 'Crédito'
        , ml.cct AS 'Centro Custo'
        , ml.conta AS 'Conta'
        , ml.descricao AS 'Nome Conta'
        , ml.edeb - ml.ecre AS 'Valor'
        , ABS(ml.edeb - ml.ecre) AS 'ABS'
        , ml.intid AS 'Id Interna'
        FROM ml
        WHERE conta LIKE '{conta}'
        AND YEAR(data) = {ano}
        AND MONTH(data) = {mes}"""

        df = pd.read_sql(query, engine)
        return df
    except Exception as e:
        print("Erro:", e)
        return None

# Exemplo de como chamar a função
# get_bank_transactions('120501', 5)  # Para maio de 2024
# get_bank_transactions('120501', 5, 2023)  # Para maio de 2023