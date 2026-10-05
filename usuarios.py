import time
import os

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from cursos import insere_cursos
from utils import formatar_nome, get_logger, salvar_screenshot_erro
from informativo import envia_email, envia_email_usuario_existente
from dados import salva_planilha

logger = get_logger()

def env_bool(nome, padrao=True):
    valor = os.getenv(nome)
    if valor is None:
        return padrao

    return valor.strip().lower() in ("1", "true", "sim", "s", "yes", "y")

CURSOS_ATRIBUIDOS_XPATH = (
    '//*[@id="konviva_content_root"]/div/div[4]/div[2]/div/div/section/div[2]/div/div/div/div[2]'
    '/div/div/div[2]/section/div'
)
PAGINADOR_CURSOS_XPATH = "//ul[contains(@class, 'pagination-detalhe-equipe')]"

def lista_cursos_da_celula(valor):
    if valor is None or valor != valor:
        return []

    return [curso.strip() for curso in str(valor).split(";") if curso.strip()]

def valor_coluna(usuario, coluna):
    if coluna not in usuario.index:
        raise KeyError(f"Coluna obrigatória não encontrada: {coluna}. Colunas disponíveis: {', '.join(usuario.index)}")

    return usuario[coluna]

def cursos_da_pagina_atual(espera):
    div_trilhas = espera.until(ec.presence_of_element_located((By.XPATH, CURSOS_ATRIBUIDOS_XPATH)))
    busca_cursos = div_trilhas.find_elements(By.XPATH, ".//h2[@title]")
    return [el.get_attribute("title").strip() for el in busca_cursos if el.get_attribute("title")]

def pagina_atual_cursos(nav):
    paginas_ativas = nav.find_elements(
        By.XPATH,
        f"{PAGINADOR_CURSOS_XPATH}//li[contains(@class, 'active')]/a"
    )
    if not paginas_ativas:
        return 1

    texto_pagina = paginas_ativas[0].text.strip()
    return int(texto_pagina) if texto_pagina.isdigit() else 1

def botao_proxima_pagina_cursos(nav):
    botoes = nav.find_elements(
        By.XPATH,
        f"{PAGINADOR_CURSOS_XPATH}//a[contains(@ng-click, 'currentPage + 1')]"
    )
    if not botoes:
        return None

    botao = botoes[0]
    item_paginacao = botao.find_element(By.XPATH, "./ancestor::li[1]")
    classes = item_paginacao.get_attribute("class") or ""
    if "disabled" in classes.split() or botao.get_attribute("disabled"):
        return None

    return botao

def cursos_atribuidos_usuario(nav, espera):
    cursos_encontrados = []
    paginas_visitadas = set()

    while True:
        pagina_atual = pagina_atual_cursos(nav)
        if pagina_atual in paginas_visitadas:
            break

        paginas_visitadas.add(pagina_atual)

        for curso in cursos_da_pagina_atual(espera):
            if curso not in cursos_encontrados:
                cursos_encontrados.append(curso)

        logger.info(f"Cursos coletados da aba {pagina_atual}")

        proxima_pagina = botao_proxima_pagina_cursos(nav)
        if not proxima_pagina:
            break

        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", proxima_pagina)
        proxima_pagina.click()

        logger.info("Aguardando carregamento dos elementos da próxima aba")
        time.sleep(30)

        try:
            WebDriverWait(nav, 10).until(lambda driver: pagina_atual_cursos(driver) != pagina_atual)
        except TimeoutException:
            logger.warning("Tempo limite ao aguardar troca da aba de cursos; encerrando coleta de cursos atribuídos")
            break

        time.sleep(2)

    return cursos_encontrados

# Verifica existência do usuário
def usuario_existe(nav, espera, login_cpf):
    # Acessa tela de usuários
    administracao = espera.until(ec.presence_of_element_located(("id", "btn_toggle_dropdown_menu_module-users")))
    administracao.click()
    nav.find_element("id", "menu_item_link-users-user").click()
    time.sleep(2)
    # Busca pelo CPF
    campo_busca = nav.find_element("id", "input_search_field")
    campo_busca.clear()
    campo_busca.send_keys(login_cpf)
    time.sleep(1)
    nav.find_element("id", "btn_search_field").click()
    time.sleep(3)

    # Verifica se há resultados
    celulas = nav.find_elements(By.XPATH, "//table[contains(@class, 'dataTable')]/tbody/tr/td[2]")
    for celula in celulas:
        if login_cpf in celula.text:
            return True
    return False

def form_usuario(nav, nome, login_cpf, senha_padrao, email):
    try:
        # Preenche o formulário
        espera = WebDriverWait(nav, 10)
        input_name = espera.until(ec.presence_of_element_located(("id", "input_name")))
        input_name.send_keys(nome)
        nav.find_element("id", "input_login").send_keys(login_cpf)
        nav.find_element("id", "input_pass").send_keys(senha_padrao)
        nav.find_element("id", "input_pass_2").send_keys(senha_padrao)
        nav.find_element("id", "input_email").send_keys(email)
        nav.find_element("id", "input_CPF").send_keys(login_cpf)
        aluno = espera.until(ec.presence_of_element_located((
            By.XPATH,
            '/html/body/div[1]/div[3]/div/div[4]/div/div/div/div/div[2]/div/div/div[1]/div/div[1]/div/form/div/div/span/ng-include/div/div/div/span[26]/span/ng-include/span/span/div[2]/div[2]/table/tbody/tr[2]/td[1]/div/div/label'
        )))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", aluno)
        aluno.click()
        uni_org = espera.until(ec.presence_of_element_located(("id", "btn_modal_select_2")))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", uni_org)
        uni_org.click()
        leclair = espera.until(ec.presence_of_element_located((By.XPATH, '//*[@id="input_radio_24731"]')))
        leclair.click()
        confirma = espera.until(ec.presence_of_element_located(("id", "btn_select_unit")))
        confirma.click()
        # Salva
        salva = espera.until(ec.presence_of_element_located(("id", "usuario_botao_salvar_usuario")))
        salva.click()
        time.sleep(2)
    except Exception as e:
        logger.error(f"Erro na criação do usuário {nome} | {login_cpf} | {email}. Erro: {e}")
        raise

def cadastro(nav, espera, df, usuarios_pendentes, senha_padrao, caminho):
    sheet_name = os.getenv("SHAREPOINT_SHEET", "Sheet1")
    resumo = {
        "total_pendentes": len(usuarios_pendentes),
        "ok": [],
        "n_ok": [],
    }
    logger.info("Iniciando processamento dos usuários")
    for u, usuario in usuarios_pendentes.iterrows():
        # Busca os dados dos usuários
        nome = formatar_nome(usuario["Nome Completo"])
        cpf = usuario["CPF"]
        email = usuario["E-mail"]
        email = email.lower()
        ajuste_cpf = []
        # Padroniza o cpf para somente números
        for digit in cpf:
            if digit.isnumeric():
                ajuste_cpf.append(digit)
            else:
                continue
        login_cpf = "".join(ajuste_cpf)
        logger.info(f"Processando usuário: {nome} | CPF: {login_cpf} | Email: {email}")
        usuario_resumo = {
            "nome": nome,
            "cpf": login_cpf,
            "email": email,
        }

        try:
            # Lista de outros cursos selecionados
            lista_outros_cursos = lista_cursos_da_celula(valor_coluna(usuario, "Outros cursos"))
            logger.info(f"Outros cursos solicitados: {lista_outros_cursos}")

            # Lista de cursos do ERP SENIOR
            lista_cursos_erp = lista_cursos_da_celula(
                valor_coluna(usuario, "Cursos Disponíveis dentro da\xa0Gestão Empresarial | ERP SENIOR")
            )
            logger.info(f"Cursos ERP solicitados: {lista_cursos_erp}")

            # Equivalência de nomes de "Outros cursos"
            equivalencia_outros_cursos = {
                "Gestão de Pessoas | Recursos Humanos": "Gestão de Pessoas | HCM",
                "Otimização | Logística": "Otimização Logística",
                "Gestão de Pátio YMS | Logística": "Otimização Logística",
                "Gestão de Armazenagem WMS | Logística": "Gestão de Armazenagem | WMS Senior",
                "Gestão de Mão de Obra na Armazenagem | Logística": "Gestão de Mão de Obra no Armazém",
                "Gestão Empresarial | ERP - (NOVO)": "Gestão Empresarial | ERP XT - Trilha"
            }
            lista_outros_cursos = [equivalencia_outros_cursos.get(curso, curso) for curso in lista_outros_cursos]

            # Equivalência de nomes de Cursos ERP
            equivalencia_cursos_erp = {
                "Cadastros": "ERP XT Cadastros Iniciais",
                "Mercado": "ERP XT Mercado",
                "Manufatura": "ERP XT Manufatura",
                "Qualidade": "ERP XT Qualidade",
                "Finanças": "ERP XT Finanças",
                "Custos": "ERP XT Custos",
                "Integrações": "ERP XT Integrações",
                "Suprimentos": "ERP XT Suprimentos",
                "Serviços": "ERP XT Serviços",
                "Controladoria": "ERP XT Controladoria"
            }
            lista_cursos_erp = [equivalencia_cursos_erp.get(curso, curso) for curso in lista_cursos_erp]
            lista_cursos = list(dict.fromkeys(lista_outros_cursos + lista_cursos_erp))
            logger.info(f"Cursos finais para atribuição: {lista_cursos}")
            usuario_resumo["cursos_solicitados"] = lista_cursos
            usuario_resumo["cursos_ja_atribuidos"] = []
            usuario_resumo["cursos_atribuidos_execucao"] = []

            # Verifica se o usuário já existe
            if usuario_existe(nav, espera, login_cpf):
                logger.info(f"Usuário {nome} ({login_cpf}) já existe")
                hist = espera.until(ec.presence_of_element_located(("id", f"btn_history-login={login_cpf}")))
                hist.click()
                logger.info("Verificando cursos já atribuídos ao usuário")
                logger.info("Aguardando carregamento dos elementos da página")
                time.sleep(45)

                # Busca cursos já atribuídos
                cursos_encontrados = cursos_atribuidos_usuario(nav, espera)
                cursos_faltantes = [curso for curso in lista_cursos if curso not in cursos_encontrados]
                usuario_resumo["cursos_ja_atribuidos"] = [
                    curso for curso in lista_cursos if curso in cursos_encontrados
                ]
                logger.info(f"Cursos já atribuídos a {nome}: {cursos_encontrados}")

                # Insere nos cursos faltantes
                if cursos_faltantes:
                    logger.info(f"Cursos ainda necessários para {nome}: {cursos_faltantes}")
                    insere_cursos(nav, espera, nome, login_cpf, cursos_faltantes)
                    usuario_resumo["cursos_atribuidos_execucao"] = cursos_faltantes
                    if env_bool("SEND_EXISTING_USER_EMAIL", True):
                        envia_email_usuario_existente(nome, email, cursos_faltantes)
                else:
                    logger.info(f"Todos cursos já atribuídos a {nome} ({login_cpf})")

            else:
                # Adiciona Usuário
                logger.info(f"Usuário {nome} ({login_cpf}) não encontrado")
                logger.info("Iniciando criação do usuário")
                nav.find_element("id", "btn_toggle_dropdown_menu_module-users").click()
                nav.find_element("id", "menu_item_link-users-user").click()
                adiciona = espera.until(ec.presence_of_element_located(("id", "btn_add-user")))
                adiciona.click()
                form_usuario(nav, nome, login_cpf, senha_padrao, email)
                logger.info(f"Usuário {nome} ({login_cpf}) criado com sucesso")
                cursos_faltantes = lista_cursos
                insere_cursos(nav, espera, nome, login_cpf, cursos_faltantes)
                usuario_resumo["cursos_atribuidos_execucao"] = cursos_faltantes

                # Envia email com informativo
                envia_email(nome, email, login_cpf, senha_padrao)

            # Atualiza planilha
            df.at[usuario.name, "Status"] = "Ok"

            try:
                salva_planilha(df, caminho, sheet_name=sheet_name)
                logger.info("Planilha atualizada")
            except Exception as e:
                logger.error(f"Erro ao salvar planilha: {e}")
                raise
            logger.info("Processo de cadastro finalizado com sucesso")
            resumo["ok"].append(usuario_resumo)

        except Exception as e:
            logger.error(f"Erro ao processar o usuário {nome} ({login_cpf}): {e}", exc_info=True)
            screenshot_path = salvar_screenshot_erro(nav, f"usuario_{login_cpf}")
            if screenshot_path:
                usuario_resumo["screenshot"] = screenshot_path
                logger.info(f"Screenshot do erro salvo em: {screenshot_path}")
            usuario_resumo["erro"] = str(e)
            resumo["n_ok"].append(usuario_resumo)
            # Atualiza planilha
            df.at[usuario.name, "Status"] = "N Ok"
            try:
                salva_planilha(df, caminho, sheet_name=sheet_name)
                logger.info("Planilha atualizada")
            except Exception as e:
                logger.error(f"Erro ao salvar planilha: {e}")

    return resumo
