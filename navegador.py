import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from utils import get_logger

logger = get_logger()

# Instancia do navegador (Chrome)
def navegador():
    try:
        logger.info("Abrindo navegador")
        nav = webdriver.Chrome()
        nav.get("https://ucsonline.senior.com.br/lms/index.html#/home")
        nav.maximize_window()
        espera = WebDriverWait(nav, timeout=10)
        logger.info("Navegador aberto com sucesso")
        return nav, espera
    except Exception as e:
        logger.error(f"Erro na abertura do navegador: {e}")

# Fecha o navegador após o processamento
def fecha_nav(nav):
    nav.quit()

# Define o arquivo que contém os dados para login de supervisor
def credenciais_login():
    with open("credenciais_login.txt", "r", encoding="utf-8") as credenciais:
        linhas = credenciais.readlines()
        usuario_adm = linhas[0].strip()
        senha_adm = linhas[1].strip()
        return usuario_adm, senha_adm

# Realiza login
def login(nav, espera, usuario, senha):
    try:
        logger.info("Iniciando login na UCS")
        cookies = espera.until(ec.presence_of_element_located((
            By.XPATH, '//*[@id="modal-warning-politica"]/button')))
        nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", cookies)
        cookies.click()

        time.sleep(2)
        cliente = espera.until(ec.presence_of_element_located((
            By.XPATH, '/html/body/div[1]/div[2]/div/div[4]/div/div/main/div/div[1]/section[2]/div/div/div/div/div[2]/div/div/form/div/a[4]')))
        cliente.click()

        nav.find_element("id", "input_username-login").send_keys(usuario)
        nav.find_element("id", "input_password-login").send_keys(senha)
        nav.find_element("id", "btn_enter-login").click()
        logger.info("Login realizado com sucesso")

        # Acessa perfil de supervisor
        logger.info("Acessando perfil de supervisor")
        time.sleep(5)
        troca_perfil = espera.until(ec.presence_of_element_located(("id", "btn_changeprofile")))
        troca_perfil.click()
        supervisor = espera.until(ec.presence_of_element_located(("id", "menu_perfil_SUP")))
        supervisor.click()
        logger.info("Perfil de supervisor acessado com sucesso")
    except Exception as e:
        logger.error(f"Erro no login na UCS: {e}")
