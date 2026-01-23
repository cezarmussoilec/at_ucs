import pandas as pd

# Abre documento Excel e busca somente pelos usuários com Status diferentes de "C"
def carrega_pendentes(caminho: str):
    df = pd.read_excel(caminho)
    usuarios_pendentes = df[(df["Status"] != "Ok") & (df["Status"] != "N Ok")]
    return df, usuarios_pendentes

def carrega_planilha(caminho: str):
    df = pd.read_excel(caminho)
    return df
