import logging
import os, sys
from datetime import datetime
from urllib.parse import urlparse
from dotenv import load_dotenv

load_dotenv()

def is_url(value):
    parsed = urlparse(value)
    return parsed.scheme in ("http", "https") and bool(parsed.netloc)

# Formata nome do usuário antes da criação
def formatar_nome(nome: str) -> str:
    partes = nome.lower().split()
    preposicoes = {"de", "da", "do", "dos", "das"}
    return " ".join([p if p in preposicoes else p.capitalize() for p in partes])

# Busca as credenciais da UCS nas variáveis de ambiente
def credenciais_ucs():
    usuario_ucs = os.getenv("UCS_USER")
    senha_ucs = os.getenv("UCS_PASSWORD")

    faltando = []
    if not usuario_ucs:
        faltando.append("UCS_USER")
    if not senha_ucs:
        faltando.append("UCS_PASSWORD")

    if faltando:
        raise RuntimeError(f"Credenciais UCS ausentes no .env: {', '.join(faltando)}")

    return usuario_ucs, senha_ucs

# Busca a senha padrão para os novos usuários nas variáveis de ambiente
def credenciais_padrao():
    senha_padrao = os.getenv("STD_PASS")

    if not senha_padrao:
        raise RuntimeError("Senha padrão ausente no .env: STD_PASS")

    return senha_padrao


# Função auxiliar para o caminho da planilha
def caminho_relativo(nome_arquivo: str) -> str:
    if getattr(sys, 'frozen', False):
        pasta_base = os.path.dirname(sys.executable)
    else:
        pasta_base = os.path.dirname(__file__)
    return os.path.join(pasta_base, nome_arquivo)

def get_logger():
    # Cria pasta "logs" caso não existir
    logs_dir = os.path.join(os.path.dirname(__file__), "logs")
    os.makedirs("logs", exist_ok=True)

    # Gera nome do arquivo com data e hora da execução
    data_atual = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    nome_arquivo = f"AT_UCS_{data_atual}.log"
    caminho_arquivo = os.path.join(logs_dir, nome_arquivo)

    logger = logging.getLogger("automacao_ucs")
    logger.setLevel(logging.INFO)

    # Evita duplicar handlers
    if not logger.handlers:
        formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s",
                                      datefmt="%Y-%m-%d %H:%M:%S")

        # Handler para arquivo
        file_handler = logging.FileHandler(caminho_arquivo, encoding="utf-8")
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)

        # Handler para console
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    return logger
