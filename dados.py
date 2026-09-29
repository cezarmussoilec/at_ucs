import pandas as pd
from utils import get_logger, is_url
from graph_api import load_sheet_as_dataframe, write_dataframe_to_workbook

logger = get_logger()

def carrega_pendentes(caminho_or_df):
    if isinstance(caminho_or_df, pd.DataFrame):
        df = caminho_or_df.copy()
    else:
        df = pd.read_excel(caminho_or_df, sheet_name=0)

    # Padroniza nomes das colunas
    df.columns = df.columns.astype(str).str.strip()

    # Garante que a coluna "Status" existe
    if "Status" not in df.columns:
        logger.error("Planilha não contém coluna 'Status'. Colunas encontradas: %s", ", ".join(df.columns))
        return df, pd.DataFrame()  # Sem pendentes se a coluna não existir

    # Usuários pendentes: Status diferentes de "Ok" e "N Ok"
    usuarios_pendentes = df[(df["Status"] != "Ok") & (df["Status"] != "N Ok")]

    return df, usuarios_pendentes

def carrega_planilha_sharepoint(workbook_url, sheet_name='Sheet1'):
    df = load_sheet_as_dataframe(workbook_url, sheet_name)
    return carrega_pendentes(df)

def salva_planilha(df, origem, sheet_name='Sheet1'):
    if is_url(origem):
        write_dataframe_to_workbook(df, origem, sheet_name=sheet_name)
        return

    df.to_excel(origem, index=False)
