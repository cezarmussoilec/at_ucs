import os
import sys
import time
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Instancia do navegador (Chrome)
def navegador():
    nav = webdriver.Chrome()
    nav.get("https://ucsonline.senior.com.br/lms/index.html#/home")
    nav.maximize_window()
    return nav

# Define o arquivo que contém os dados para login de supervisor
def credenciais_login():
    with open("credenciais_login.txt", "r", encoding="utf-8") as credenciais:
        linhas = credenciais.readlines()
        usuario_adm = linhas[0].strip()
        senha_adm = linhas[1].strip()
        return usuario_adm, senha_adm

# Realiza login
def login(nav, usuario, senha):
    # Cookies
    espera = WebDriverWait(nav, 10)
    cookies = espera.until(EC.presence_of_element_located((
        By.XPATH,
        '//*[@id="modal-warning-politica"]/button'
    )))
    nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", cookies)
    cookies.click()
    # Cliente
    cliente = espera.until(EC.presence_of_element_located((
        By.XPATH,
        '/html/body/div[1]/div[2]/div/div[4]/div/div/main/div/div[1]/section[2]/div/div/div/div/div[2]/div/div/form/div/a[4]'
    )))
    cliente.click()
    # Login
    nav.find_element("id", "input_username-login").send_keys(usuario)
    nav.find_element("id", "input_password-login").send_keys(senha)
    nav.find_element("id", "btn_enter-login").click()
    # Perfil Superior
    troca_perfil = espera.until(EC.presence_of_element_located((
        "id", "btn_changeprofile")))
    troca_perfil.click()
    nav.find_element("id", "menu_perfil_SUP").click()

# Abre documento Excel e busca somente pelos usuários com Status diferentes de "C"
def carrega_planilha(caminho):
    df = pd.read_excel(caminho)
    usuarios_pendentes = df[df["Status"] != "C"]
    return df, usuarios_pendentes

def cadastro(nav, df, usuarios_pendentes, senha_pad, caminho):
    espera = WebDriverWait(nav, 10)
    for u, usuario in usuarios_pendentes.iterrows():
        # Busca os dados dos usuários
        nome = formatar_nome(usuario["Nome Completo"])
        CPF = usuario["CPF"]
        email = usuario["E-mail"]
        ajusteCPF = []
        # Padroniza o CPF para somente números
        for digit in CPF:
            if digit.isnumeric():
                ajusteCPF.append(digit)
            else:
                continue
        loginCPF = "".join(ajusteCPF)
        cursos_str = usuario["Cursos Disponíveis"].split(";")
        lista_cursos = [curso.strip() for curso in cursos_str if curso.strip()]
        equivalencia_cursos = {
            "Gestão de Pessoas | Recursos Humanos": "Gestão de Pessoas | HCM",
            "Otimização | Logística": "Otimização Logística",
            "Gestão de Pátio YMS | Logística": "Otimização Logística",
            "Gestão de Armazenagem WMS | Logística": "Gestão de Armazenagem | WMS Senior",
            "Gestão de Mão de Obra na Armazenagem | Logística": "Gestão de Mão de Obra no Armazém"
        }
        lista_cursos = [equivalencia_cursos.get(curso, curso) for curso in lista_cursos]
        # Verifica se o usuário já existe
        if usuario_existe(nav, loginCPF):
            espera = WebDriverWait(nav, 10)
            hist = espera.until(EC.presence_of_element_located(("id", f"btn_history-login={loginCPF}")))
            hist.click()
            time.sleep(5)
            busca_cursos = nav.find_elements(By.XPATH, "//h2[@title]")
            cursos_encontrados = [el.get_attribute("title").strip() for el in busca_cursos if el.get_attribute("title")]
            cursos_encontrados = [equivalencia_cursos.get(curso, curso) for curso in cursos_encontrados]
            cursos_faltantes = [curso for curso in lista_cursos if curso not in cursos_encontrados]

            # Insere nos cursos faltantes e gera log
            if cursos_faltantes:
                insere_cursos(nav, nome, cursos_faltantes)

        else:
            # Adiciona Usuário
            nav.find_element("id", "btn_toggle_dropdown_menu_module-users").click()
            nav.find_element("id", "menu_item_link-users-user").click()
            adiciona = espera.until(EC.presence_of_element_located(("id", "btn_add-user")))
            adiciona.click()
            form_usuario(nav, nome, loginCPF, senha_pad, email)
            insere_cursos(nav, nome, lista_cursos)
            cursos_faltantes = lista_cursos

        df.at[usuario.name, "Status"] = "C"
        df.to_excel(caminho, index=False)
        with open("log.txt", "a", encoding="utf-8") as log:
            log.write(
                f"{usuario['Nome Completo']} ({loginCPF}):\n"
                f"  Cursos escolhidos: {', '.join(lista_cursos)}\n"
                f"  Cursos atribuídos: {', '.join(cursos_faltantes) if cursos_faltantes else 'Todos já atribuídos'}\n")

# Verifica existência do usuário
def usuario_existe(nav, loginCPF):
    # Acessa tela de usuários
    nav.find_element("id", "btn_toggle_dropdown_menu_module-users").click()
    nav.find_element("id", "menu_item_link-users-user").click()
    time.sleep(1)
    # Busca pelo CPF
    campo_busca = nav.find_element("id", "input_search_field")
    campo_busca.clear()
    campo_busca.send_keys(loginCPF)
    nav.find_element("id", "btn_search_field").click()
    time.sleep(1)

    # Verifica se há resultados
    celulas = nav.find_elements(By.XPATH, "//table[contains(@class, 'dataTable')]/tbody/tr/td[2]")
    for celula in celulas:
        if loginCPF in celula.text:
            return True
    return False

def form_usuario(nav, nome, loginCPF, senha_pad, email):
    # Preenche o formulário
    espera = WebDriverWait(nav, 10)
    input = espera.until(EC.presence_of_element_located(("id", "input_name")))
    input.send_keys(nome)
    nav.find_element("id", "input_login").send_keys(loginCPF)
    nav.find_element("id", "input_pass").send_keys(senha_pad)
    nav.find_element("id", "input_pass_2").send_keys(senha_pad)
    nav.find_element("id", "input_email").send_keys(email)
    nav.find_element("id", "input_CPF").send_keys(loginCPF)
    aluno = espera.until(EC.presence_of_element_located((
        By.XPATH,
        '//*[@id="usuario_fieldsetcontroller_ng"]/div/div/span[26]/span/ng-include/span/span/div[2]/div[2]/table/tbody/tr[2]/td[1]/div'
    )))
    nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", aluno)
    aluno.click()
    uni_org = espera.until(EC.presence_of_element_located(("id", "btn_modal_select_2")))
    nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", uni_org)
    uni_org.click()
    leclair = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="input_radio_24731"]')))
    leclair.click()
    confirma = espera.until(EC.presence_of_element_located(("id", "btn_select_unit")))
    confirma.click()
    # Salva
    salva = espera.until(EC.presence_of_element_located(("id", "usuario_botao_salvar_usuario")))
    salva.click()
    time.sleep(2)

# Acessa catálogo de cursos
def acessa_catalogo(nav):
    nav.find_element("id", "btn_toggle_dropdown_menu_module-learning").click()
    nav.find_element("id", "menu_item_link-learning-catalog_manager").click()
    espera = WebDriverWait(nav, 10)
    nossas_solucoes = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[2]/button')))
    nossas_solucoes.click()
    time.sleep(1)

# Chama as funções para inserir em cada curso
def insere_cursos(nav, nome, cursos_faltantes):
    curso_otimizacao_yms = False

    for curso in cursos_faltantes:
        if curso == "Gestão de Pessoas | HCM":
            acessa_catalogo(nav)
            gestao_pessoas(nav, nome)
        elif curso == "Otimização Logística":
            if not curso_otimizacao_yms:
                acessa_catalogo(nav)
                otimizacao_yms(nav, nome)
                curso_otimizacao_yms = True
        elif curso == "Gestão de Armazenagem | WMS Senior":
            acessa_catalogo(nav)
            gestao_wms(nav, nome)
        elif curso == "Gestão de Mão de Obra no Armazém":
            acessa_catalogo(nav)
            gestao_obra(nav, nome)
        else:
            return False

# Curso: Gestão de Pessoas | HCM
def gestao_pessoas(nav, nome):
    espera = WebDriverWait(nav, 10)
    hcm = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[1]/button')))
    hcm.click()
    saiba_mais = espera.until(EC.presence_of_element_located(("id", "btn-sm-Gestão de Pessoas | HCM")))
    nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
    saiba_mais.click()
    time.sleep(5)
    associar_usuario = espera.until(EC.presence_of_element_located((By.XPATH,
         '//*[@id="konviva_content_root"]/div/div[4]/div/div/div/div/div/div/div[2]/div[3]/div[1]/a')))
    associar_usuario.click()
    busca_nome = espera.until(EC.presence_of_element_located(("id", "input_search_field")))
    busca_nome.send_keys(nome)
    buscar = espera.until(EC.presence_of_element_located(("id", "btn_search_field")))
    buscar.click()
    time.sleep(2)
    check = espera.until(EC.presence_of_element_located(("id", "checkboxID_0")))
    check.click()
    associar = espera.until(EC.presence_of_element_located((By.XPATH,
        '//*[@id="konviva_content_root"]/div/div[4]/div/div[2]/div/div/div/div/div/div/div/div[3]/div/button')))
    associar.click()
    # Confirma
    botao_confirmar = espera.until(EC.presence_of_element_located((By.XPATH,"//button[contains(text(), 'Confirmar')]")))
    botao_confirmar.click()
    time.sleep(1)
    nav.refresh()
    time.sleep(3)

# Curso: Otimização Logística
def otimizacao_yms(nav, nome):
    espera = WebDriverWait(nav, 10)
    log = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[5]/button')))
    log.click()
    saiba_mais = espera.until(EC.presence_of_element_located(("id", "btn-sm-Otimização Logística")))
    nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
    saiba_mais.click()
    time.sleep(5)
    associar_usuario = espera.until(EC.presence_of_element_located((By.XPATH,
         '//*[@id="konviva_content_root"]/div/div[4]/div/div/div/div/div/div/div[2]/div[3]/div[1]/a')))
    associar_usuario.click()
    busca_nome = espera.until(EC.presence_of_element_located(("id", "input_search_field")))
    busca_nome.send_keys(nome)
    buscar = espera.until(EC.presence_of_element_located(("id", "btn_search_field")))
    buscar.click()
    time.sleep(2)
    check = espera.until(EC.presence_of_element_located(("id", "checkboxID_0")))
    check.click()
    associar = espera.until(EC.presence_of_element_located((By.XPATH,
        '//*[@id="konviva_content_root"]/div/div[4]/div/div[2]/div/div/div/div/div/div/div/div[3]/div/button')))
    associar.click()
    # Confirma
    botao_confirmar = espera.until(
        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Confirmar')]")))
    botao_confirmar.click()
    time.sleep(1)
    nav.refresh()
    time.sleep(3)

# Curso: Gestão de Armazenagem | WMS Senior
def gestao_wms(nav, nome):
    espera = WebDriverWait(nav, 10)
    wms = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="categories_list-catalog"]/li[5]/button')))
    wms.click()
    saiba_mais = espera.until(EC.presence_of_element_located(("id", "btn-sm-Gestão de Armazenagem | WMS Senior")))
    nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
    saiba_mais.click()
    time.sleep(5)
    associar_usuario = espera.until(EC.presence_of_element_located((By.XPATH,
        '//*[@id="konviva_content_root"]/div/div[4]/div/div/div/div/div/div/div[2]/div[3]/div[1]/a')))
    associar_usuario.click()
    busca_nome = espera.until(EC.presence_of_element_located(("id", "input_search_field")))
    busca_nome.send_keys(nome)
    buscar = espera.until(EC.presence_of_element_located(("id", "btn_search_field")))
    buscar.click()
    time.sleep(2)
    check = espera.until(EC.presence_of_element_located(("id", "checkboxID_0")))
    check.click()
    associar = espera.until(EC.presence_of_element_located((By.XPATH,
        '//*[@id="konviva_content_root"]/div/div[4]/div/div[2]/div/div/div/div/div/div/div/div[3]/div/button')))
    associar.click()
    # Confirma
    botao_confirmar = espera.until(
        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Confirmar')]")))
    botao_confirmar.click()
    time.sleep(1)
    nav.refresh()
    time.sleep(3)

# Curso: Gestão de Mão de Obra no Armazém
def gestao_obra(nav, nome):
    espera = WebDriverWait(nav, 10)
    obra = espera.until(EC.presence_of_element_located((By.XPATH,
        '//*[@id="categories_list-catalog"]/li[5]/button')))
    obra.click()
    saiba_mais = espera.until(EC.presence_of_element_located((
        "id", "btn-sm-Gestão de Mão de Obra no Armazém")))
    nav.execute_script("arguments[0].scrollIntoView({block: 'center'});", saiba_mais)
    saiba_mais.click()
    time.sleep(5)
    associar_usuario = espera.until(EC.presence_of_element_located((By.XPATH,
         '//*[@id="konviva_content_root"]/div/div[4]/div/div/div/div/div/div/div[2]/div[3]/div[1]/a')))
    associar_usuario.click()
    busca_nome = espera.until(EC.presence_of_element_located(("id", "input_search_field")))
    busca_nome.send_keys(nome)
    buscar = espera.until(EC.presence_of_element_located(("id", "btn_search_field")))
    buscar.click()
    time.sleep(2)
    check = espera.until(EC.presence_of_element_located(("id", "checkboxID_0")))
    check.click()
    associar = espera.until(EC.presence_of_element_located((By.XPATH,
        '//*[@id="konviva_content_root"]/div/div[4]/div/div[2]/div/div/div/div/div/div/div/div[3]/div/button')))
    associar.click()
    # Confirma
    botao_confirmar = espera.until(
        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Confirmar')]")))
    botao_confirmar.click()
    time.sleep(1)
    nav.refresh()
    time.sleep(3)

# Formata nome do usuário antes da criação
def formatar_nome(nome):
    partes = nome.lower().split()
    preposicoes = {"de", "da", "do", "dos", "das"}
    nome_formatado = " ".join([p if p in preposicoes else p.capitalize() for p in partes])
    return nome_formatado

# Define o arquivo que irá possuir a senha padrão para criação dos usuários
def senha_padrao():
    with open("senha_padrao.txt", "r", encoding="utf-8") as senha_pad:
        return senha_pad.read().strip()

# Função auxiliar para o caminho da planilha
def caminho_relativo(nome_arquivo):
    if getattr(sys, 'frozen', False):
        pasta_base = os.path.dirname(sys.executable)
    else:
        pasta_base = os.path.dirname(__file__)
    return os.path.join(pasta_base, nome_arquivo)

# Selecionaa planilha contendo informações dos usuários
def selecionar_planilha():
    caminho = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Excel files", "*.xlsx")])
    if caminho:
        executar_script(caminho)

# Executa script de automação
def executar_script(caminho):
    try:
        nav = navegador()
        usuario_adm, senha_adm = credenciais_login()
        login(nav, usuario_adm, senha_adm)
        df, usuarios_pendentes = carrega_planilha(caminho)
        cadastro(nav, df, usuarios_pendentes, senha_padrao(), caminho)
        messagebox.showinfo("Finalizado", "Processamento concluído com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Inicia interface
def iniciar_interface():
    root = tk.Tk()
    root.title("Automação UCS")
    largura = 275
    altura = 150
    root.geometry(f"{largura}x{altura}")
    root.resizable(False, False)
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura // 2)
    pos_y = (altura_tela // 2) - (altura // 2)
    root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")
    label = tk.Label(root, text="Selecione a planilha para iniciar:", pady=20)
    label.pack()
    botao = tk.Button(root, text="Selecionar Planilha", command=selecionar_planilha)
    botao.pack()
    root.mainloop()

def main():
    iniciar_interface()

if __name__ == '__main__':
    main()