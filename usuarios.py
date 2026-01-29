import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from cursos import insere_cursos
from utils import formatar_nome, get_logger
from informativo import envia_email

logger = get_logger()

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

def form_usuario(nav, nome, login_cpf, senha_pad, email):
    try:
        # Preenche o formulário
        espera = WebDriverWait(nav, 10)
        input_name = espera.until(ec.presence_of_element_located(("id", "input_name")))
        input_name.send_keys(nome)
        nav.find_element("id", "input_login").send_keys(login_cpf)
        nav.find_element("id", "input_pass").send_keys(senha_pad)
        nav.find_element("id", "input_pass_2").send_keys(senha_pad)
        nav.find_element("id", "input_email").send_keys(email)
        nav.find_element("id", "input_CPF").send_keys(login_cpf)
        aluno = espera.until(ec.presence_of_element_located((
            By.XPATH,
            '//*[@id="usuario_fieldsetcontroller_ng"]/div/div/span[26]/span/ng-include/span/span/div[2]/div[2]/table/tbody/tr[2]/td[1]/div'
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

def cadastro(nav, espera, df, usuarios_pendentes, senha_pad, caminho):
    logger.info("Iniciando processamento dos usuários")
    for u, usuario in usuarios_pendentes.iterrows():
        # Busca os dados dos usuários
        nome = formatar_nome(usuario["Nome Completo"])
        cpf = usuario["CPF"]
        email = usuario["E-mail"]
        ajuste_cpf = []
        # Padroniza o cpf para somente números
        for digit in cpf:
            if digit.isnumeric():
                ajuste_cpf.append(digit)
            else:
                continue
        login_cpf = "".join(ajuste_cpf)
        logger.info(f"Processando usuário: {nome} | CPF: {login_cpf} | Email: {email}")

        try:
            # Lista de cursos selecionados
            cursos_str = usuario["Cursos Disponíveis"].split(";")
            lista_cursos = [curso.strip() for curso in cursos_str if curso.strip()]
            logger.info(f"Cursos solicitados: {lista_cursos}")

            # Equivalência de nomes
            equivalencia_cursos = {
                "Gestão de Pessoas | Recursos Humanos": "Gestão de Pessoas | HCM",
                "Otimização | Logística": "Otimização Logística",
                "Gestão de Pátio YMS | Logística": "Otimização Logística",
                "Gestão de Armazenagem WMS | Logística": "Gestão de Armazenagem | WMS Senior",
                "Gestão de Mão de Obra na Armazenagem | Logística": "Gestão de Mão de Obra no Armazém",
                "Gestão Empresarial | ERP - (NOVO)": "Gestão Empresarial | ERP XT - Geral"
            }
            lista_cursos = [equivalencia_cursos.get(curso, curso) for curso in lista_cursos]

            # Verifica se o usuário já existe
            if usuario_existe(nav, espera, login_cpf):
                logger.info(f"Usuário {nome} ({login_cpf}) já existe")
                hist = espera.until(ec.presence_of_element_located(("id", f"btn_history-login={login_cpf}")))
                hist.click()
                logger.info("Verificando cursos já atribuídos ao usuário")
                logger.info("Aguardando carregamento dos elementos da página")
                time.sleep(45)

                # Busca cursos já atribuídos
                div_trilhas = espera.until(ec.presence_of_element_located((By.XPATH,
                    '//*[@id="konviva_content_root"]/div/div[4]/div[2]/div/div/section/div[2]/div/div/div/div[2]'
                    '/div/div/div[2]/section/div')))
                busca_cursos = div_trilhas.find_elements(By.XPATH, ".//h2[@title]")
                cursos_encontrados = [el.get_attribute("title").strip() for el in busca_cursos if el.get_attribute("title")]
                cursos_encontrados = [equivalencia_cursos.get(curso, curso) for curso in cursos_encontrados]
                cursos_faltantes = [curso for curso in lista_cursos if curso not in cursos_encontrados]
                logger.info(f"Cursos já atribuídos a {nome}: {cursos_encontrados}")

                # Insere nos cursos faltantes
                if cursos_faltantes:
                    logger.info(f"Cursos ainda necessários para {nome}: {cursos_faltantes}")
                    insere_cursos(nav, espera, nome, login_cpf, cursos_faltantes)
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
                form_usuario(nav, nome, login_cpf, senha_pad, email)
                logger.info(f"Usuário {nome} ({login_cpf}) criado com sucesso")
                cursos_faltantes = lista_cursos
                insere_cursos(nav, espera, nome, login_cpf, cursos_faltantes)

            # Envia email com informativo e atualiza planilha
            envia_email(nome, email, login_cpf, senha_pad)
            df.at[usuario.name, "Status"] = "Ok"
            df.to_excel(caminho, index=False)
            logger.info("Processo de cadastro finalizado com sucesso")

        except Exception as e:
            logger.error(f"Erro ao processar o usuário {nome} ({login_cpf}): {e}", exc_info=True)
            # Atualiza planilha
            df.at[usuario.name, "Status"] = "N Ok"
            df.to_excel(caminho, index=False)
