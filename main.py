import os
import sys
from dados import carrega_pendentes, carrega_planilha_sharepoint
from ucs import executar_script
from utils import get_logger, is_url

logger = get_logger()

def main():
    sharepoint_path = (os.getenv("SHAREPOINT_PATH") or "").strip()
    local_file = sys.argv[1] if len(sys.argv) > 1 else os.getenv("LOCAL_FILE_PATH")

    if sharepoint_path:
        if not is_url(sharepoint_path):
            logger.error("SHAREPOINT_PATH deve ser a URL completa da planilha do SharePoint.")
            logger.error("Para arquivo local, deixe SHAREPOINT_PATH vazio e configure LOCAL_FILE_PATH ou passe o arquivo por argumento.")
            sys.exit(1)

        sheet_name = os.getenv("SHAREPOINT_SHEET", "Sheet1")
        logger.info(f"Detectado: SharePoint - {sharepoint_path} (aba: {sheet_name})")
        df, usuarios_pendentes = carrega_planilha_sharepoint(sharepoint_path, sheet_name)
        arquivo_origem = sharepoint_path

    elif local_file and os.path.isfile(local_file):
        logger.info(f"Detectado: Arquivo local - {local_file}")
        df, usuarios_pendentes = carrega_pendentes(local_file)
        arquivo_origem = local_file

    else:
        logger.error("Nenhuma planilha configurada.")
        logger.error("Configure no .env: SHAREPOINT_PATH ou LOCAL_FILE_PATH")
        sys.exit(1)

    if usuarios_pendentes.empty:
        logger.info("Nenhum usuário pendente. Encerrando.")
        return

    executar_script(arquivo_origem, df, usuarios_pendentes)


if __name__ == "__main__":
    main()
