import tkinter as tk
from tkinter import filedialog
from ucs import executar_script

# Seleciona a planilha contendo informações dos usuários
def selecionar_planilha():
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if caminho:
        executar_script(caminho)

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
    botao = tk.Button(root, text="Selecionar Planilha", command=selecionar_planilha)
    botao.pack()

    root.mainloop()
