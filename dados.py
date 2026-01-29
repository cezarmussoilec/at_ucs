import pandas as pd
from utils import get_logger

logger = get_logger()

# Abre documento Excel e busca somente pelos usuários com Status diferentes de "C"
def carrega_pendentes(caminho: str):
    df = pd.read_excel(caminho)
    usuarios_pendentes = df[(df["Status"] != "Ok") & (df["Status"] != "N Ok")]

    if usuarios_pendentes.empty:
        logger.info("Nenhum usuário pendente encontrado")
    return df, usuarios_pendentes

def carrega_planilha(caminho: str):
    df = pd.read_excel(caminho)
    return df
