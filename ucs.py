from navegador import navegador, login, fecha_nav
from usuarios import cadastro
from utils import get_logger, credenciais_ucs, credenciais_padrao

# Executa script de automação
def executar_script(caminho, df, usuarios_pendentes):
    logger = get_logger()
    try:
        nav, espera = navegador() #navegador.py
        usuario_ucs, senha_ucs = credenciais_ucs() #utils.py
        login(nav, espera, usuario_ucs, senha_ucs) #navegador.py
        cadastro(nav, espera, df, usuarios_pendentes, credenciais_padrao(), caminho) #usuarios.py
        logger.info("Execução concluída com sucesso")
    except Exception as e:
        logger.error(f"Erro durante execução: {e}", exc_info=True)
    finally:
        try:
            fecha_nav(nav)
            logger.info("Navegador fechado com sucesso")
        except Exception as e:
            logger.warning(f"Erro ao fechar navegador, já estava fechado ou não foi iniciado: {e}", exc_info=True)
