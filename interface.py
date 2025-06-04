import tkinter as tk
from tkinter import messagebox
import pandas as pd
from criarHTML import processa_aba_gera_html
from atualizador_WP import atualizar_pagina_wp

# Lê o arquivo com os links
links_df = pd.read_excel("links_startups.xlsx")
abas_links = dict(zip(links_df['ABA'], links_df['LINK']))

# Cria a janela principal
root = tk.Tk()
root.title("Atualizar Páginas de Startups")
root.geometry("500x500")  # Janela mais larga

tk.Label(root, text="QUAIS ABAS DESEJA ATUALIZAR?", font=("Arial", 12)).pack(pady=10)

# Frame para conter os checkboxes
checkbox_frame = tk.Frame(root)
checkbox_frame.pack()

# Dicionário para armazenar variáveis de estado dos checkboxes
checkbox_vars = {}

# Número de colunas desejado
num_colunas = 2
abas = list(abas_links.keys())
num_linhas = (len(abas) + num_colunas - 1) // num_colunas

# Cria os checkboxes organizados em colunas
for i, aba in enumerate(abas):
    var = tk.BooleanVar()
    checkbox = tk.Checkbutton(checkbox_frame, text=aba, variable=var, anchor='w', width=25)
    checkbox.grid(row=i % num_linhas, column=i // num_linhas, sticky='w')
    checkbox_vars[aba] = var

def on_submit():
    selecionadas = [aba for aba, var in checkbox_vars.items() if var.get()]
    
    if not selecionadas:
        messagebox.showwarning("Aviso", "Selecione ao menos uma aba.")
        return

    erros = []
    for aba in selecionadas:
        try:
            
            html = processa_aba_gera_html(aba)
            if html is None:
                erros.append(f"{aba}: Erro ao gerar HTML.")
                continue

            link = abas_links[aba]
            resposta = atualizar_pagina_wp(link, html)

            if resposta is not True:
                erros.append(f"{aba}: Falha ao atualizar a página.")

        except Exception as e:
            erros.append(f"{aba}: {str(e)}")

    if not erros:
        messagebox.showinfo("Sucesso", "Todas as abas selecionadas foram atualizadas com sucesso.")
    else:
        mensagem = "Alguns erros ocorreram:\n" + "\n".join(erros)
        messagebox.showerror("Erros encontrados", mensagem)

# Botão de executar
tk.Button(root, text="Executar Selecionadas", command=on_submit).pack(pady=20)

root.mainloop()
