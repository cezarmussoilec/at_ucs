from navegador import navegador, login, credenciais_login, fecha_nav
from usuarios import cadastro
from utils import senha_padrao, get_logger

# Executa script de automação
def executar_script(caminho, df, usuarios_pendentes):
    logger = get_logger()
    try:
        nav, espera = navegador()
        usuario_adm, senha_adm = credenciais_login()
        login(nav, espera, usuario_adm, senha_adm)
        cadastro(nav, espera, df, usuarios_pendentes, senha_padrao(), caminho)
        logger.info("Execução concluída com sucesso")
    except Exception as e:
        logger.error(f"Erro durante execução: {e}", exc_info=True)
    finally:
        try:
            fecha_nav(nav)
            logger.info("Navegador fechado com sucesso")
        except Exception as e:
            logger.warning(f"Erro ao fechar navegador, já estava fechado ou não foi iniciado: {e}", exc_info=True)
