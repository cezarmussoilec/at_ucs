import sys
import tkinter as tk
from tkinter import filedialog
from ucs import executar_script
from dados import carrega_pendentes
from utils import get_logger

logger = get_logger()

# Finaliza a interface e execução
def finalizar_interface(root):
    root.destroy()
    sys.exit(0)

# Seleciona a planilha contendo informações dos usuários
def selecionar_planilha(root):
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if caminho:
        logger.info(f"Iniciando execução da automação UCS - Universidade Corporativa Senior")
        logger.info("Criação de usuários e atribuição de cursos.")
        logger.info(f"Arquivo: {caminho}")
        df, usuarios_pendentes = carrega_pendentes(caminho)
        if usuarios_pendentes.empty:
            logger.info("Finalizando execução da automação")
            finalizar_interface(root)
        else:
            executar_script(caminho, df, usuarios_pendentes)
            finalizar_interface(root)

# Inicia interface
def iniciar_interface():
    root = tk.Tk()
    root.title("Automação UCS")
    largura, altura = 275, 150
    root.geometry(f"{largura}x{altura}")
    root.resizable(False, False)

    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura // 2)
    pos_y = (altura_tela // 2) - (altura // 2)
    root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    label = tk.Label(root, text="Selecione a planilha para iniciar:", pady=20)
    label.pack()
    botao = tk.Button(root, text="Selecionar Planilha", command=lambda: selecionar_planilha(root))
    botao.pack()
    botao_sair = tk.Button(root, text="Sair", command=lambda: finalizar_interface(root))
    botao_sair.pack(pady=10)

    root.mainloop()

    return root
