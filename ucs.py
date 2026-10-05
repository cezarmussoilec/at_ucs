from navegador import navegador, login, fecha_nav
from usuarios import cadastro
from utils import get_logger, credenciais_ucs, credenciais_padrao, salvar_screenshot_erro

# Executa script de automação
def executar_script(caminho, df, usuarios_pendentes):
    logger = get_logger()
    nav = None
    try:
        nav, espera = navegador() #navegador.py
        usuario_ucs, senha_ucs = credenciais_ucs() #utils.py
        login(nav, espera, usuario_ucs, senha_ucs) #navegador.py
        resumo = cadastro(nav, espera, df, usuarios_pendentes, credenciais_padrao(), caminho) #usuarios.py
        logger.info("Execução concluída com sucesso")
        return resumo
    except Exception as e:
        logger.error(f"Erro durante execução: {e}", exc_info=True)
        screenshot_path = salvar_screenshot_erro(nav, "erro_fatal")
        if screenshot_path:
            logger.info(f"Screenshot do erro fatal salvo em: {screenshot_path}")
        raise
    finally:
        try:
            if nav:
                fecha_nav(nav)
                logger.info("Navegador fechado com sucesso")
        except Exception as e:
            logger.warning(f"Erro ao fechar navegador, já estava fechado ou não foi iniciado: {e}", exc_info=True)
