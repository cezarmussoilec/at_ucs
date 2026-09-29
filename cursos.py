import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from utils import get_logger

logger = get_logger()

# Chama as funções para inserir em cada curso
def insere_cursos(nav, espera, nome, login_cpf, cursos_faltantes):
    try:
        curso_otimizacao_yms = False

        for curso in cursos_faltantes:
            if curso == "Gestão de Pessoas | HCM":
                acessa_catalogo(nav, espera)
                gestao_pessoas(nav, nome, login_cpf, curso)
            elif curso == "Otimização Logística":
                if not curso_otimizacao_yms:
                    acessa_catalogo(nav, espera)
                    otimizacao_yms(nav, nome, login_cpf, curso)
                    curso_otimizacao_yms = True
            elif curso == "Gestão de Armazenagem | WMS Senior":
                acessa_catalogo(nav, espera)
                gestao_wms(nav, nome, login_cpf, curso)
            elif curso == "Gestão de Mão de Obra no Armazém":
                acessa_catalogo(nav, espera)
                gestao_obra(nav, nome, login_cpf, curso)
            elif curso == "Gestão Empresarial | ERP XT - Geral":
                acessa_catalogo(nav, espera)
                gestao_erp(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Cadastros Iniciais":
                acessa_catalogo(nav, espera)
                erp_cadastros(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Mercado":
                acessa_catalogo(nav, espera)
                erp_mercado(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Manufatura":
                acessa_catalogo(nav, espera)
                erp_manufatura(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Qualidade":
                acessa_catalogo(nav, espera)
                erp_qualidade(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Finanças":
                acessa_catalogo(nav, espera)
                erp_financas(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Custos":
                acessa_catalogo(nav, espera)
                erp_custos(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Integrações":
                acessa_catalogo(nav, espera)
                erp_integracoes(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Suprimentos":
                acessa_catalogo(nav, espera)
                erp_suprimentos(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Serviços":
                acessa_catalogo(nav, espera)
                erp_servicos(nav, nome, login_cpf, curso)
            elif curso == "ERP XT Controladoria":
                acessa_catalogo(nav, espera)
                erp_controladoria(nav, nome, login_cpf, curso)
            else:
                return False
    except Exception as e:
        logger.error(f"Erro ao inserir o usuário {nome} no curso {curso}: {e}")

# Acessa catálogo de cursos
def acessa_catalogo(nav, espera):
    try:
        logger.info("Acessando o catálogo de cursos")
        aprendizagem = espera.until(ec.presence_of_element_located(("id", "btn_toggle_dropdown_menu_module-learning")))
        aprendizagem.click()
        catalogo = espera.until(ec.presence_of_element_located(("id", "menu_item_link-learning-catalog_manager")))
        catalogo.click()
        nossas_solucoes = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[1]/button')))
        nossas_solucoes.click()
        time.sleep(1)
        logger.info("Catálogo acessado com sucesso")
    except Exception as e:
        logger.error(f"Erro ao acessar o catálogo: {e}", exc_info=True)

# Curso: Gestão de Pessoas | HCM
def gestao_pessoas(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        hcm = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[1]/button')))
        hcm.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-Gestão de Pessoas | HCM")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: Otimização Logística
def otimizacao_yms(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        log = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[7]/button')))
        log.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-Otimização Logística")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: Gestão de Armazenagem | WMS Senior
def gestao_wms(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        wms = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[7]/button')))
        wms.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-Gestão de Armazenagem | WMS Senior")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: Gestão de Mão de Obra no Armazém
def gestao_obra(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        obra = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[7]/button')))
        obra.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-Gestão de Mão de Obra no Armazém")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: Gestão Empresarial | ERP XT - Geral
def gestao_erp(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-Gestão Empresarial | ERP XT - Geral")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Cadastros Iniciais
def erp_cadastros(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Cadastros Iniciais")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Mercado
def erp_mercado(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Mercado")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Manufatura
def erp_manufatura(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Manufatura")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Qualidade
def erp_qualidade(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Qualidade")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Finanças
def erp_financas(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Finanças")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Custos
def erp_custos(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Custos")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Integrações
def erp_integracoes(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Integrações")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Suprimentos
def erp_suprimentos(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Suprimentos")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Serviços
def erp_servicos(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Serviços")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Curso: ERP XT Controladoria
def erp_controladoria(nav, nome, login_cpf, curso):
    try:
        logger.info(f"Acessando o curso {curso}")
        espera = WebDriverWait(nav, 10)
        erp = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
        erp.click()
        saiba_mais = espera.until(ec.presence_of_element_located(("id", "btn-sm-ERP XT Controladoria")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
        saiba_mais.click()
        associa_cursos(nav, login_cpf)
        logger.info(f"Usuário {nome} ({login_cpf}) associado com sucesso ao curso {curso}")
    except Exception as e:
        logger.error(f"Erro ao acessar o curso {curso}: {e}")
        raise

# Associa o usuário ao curso em questão
def associa_cursos(nav, login_cpf):
    try:
        logger.info("Associando usuário ao curso")
        espera = WebDriverWait(nav, 20)
        time.sleep(5)
        associar_usuario = espera.until(ec.presence_of_element_located((By.XPATH,
            '/html/body/div[1]/div[3]/div/div[4]/div/div/div/div/div/div/div/div/div[2]/div[4]/div[1]/a')))
        associar_usuario.click()
        busca_usuario = espera.until(ec.presence_of_element_located(("id", "input_search_field")))
        busca_usuario.send_keys(login_cpf)
        buscar = espera.until(ec.presence_of_element_located(("id", "btn_search_field")))
        buscar.click()
        time.sleep(2)
        check = espera.until(ec.presence_of_element_located(("id", "checkboxID_0")))
        check.click()
        associar = espera.until(ec.presence_of_element_located((By.XPATH,
            '//*[@id="konviva_content_root"]/div/div[4]/div/div[2]/div/div/div/div/div/div/div/div[3]/div/button')))
        associar.click()
        # Confirma
        botao_confirmar = espera.until(ec.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Confirmar')]")))
        botao_confirmar.click()
        time.sleep(1)
        nav.refresh()
        time.sleep(3)
    except Exception as e:
        logger.error(f"Erro ao associar o usuário: {e}", exc_info=True)
        raise
