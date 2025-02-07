import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import messagebox

# Função para encontrar o caminho correto do arquivo `dados.xlsx`
def get_file_path():
    if getattr(sys, 'frozen', False):  # Se estiver rodando como .exe
        diretorio_base = os.path.dirname(sys.executable)
    else:  # Se estiver rodando como script .py
        diretorio_base = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(diretorio_base, "dados.xlsx")

file_path = get_file_path()

# Verifica se a planilha existe antes de tentar carregar
if not os.path.exists(file_path):
    messagebox.showerror("Erro", f"O arquivo 'dados.xlsx' não foi encontrado.\nColoque-o na pasta do programa: {file_path}")
    sys.exit()

try:
    df = pd.read_excel(file_path, engine="openpyxl")
    df.columns = df.columns.str.strip()
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
except Exception as e:
    messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")
    sys.exit()

# Criar a interface Tkinter
root = tk.Tk()
root.title("Dados ADP")
root.geometry("420x250")

# Criar frame para organizar os elementos
frame_top = tk.Frame(root)
frame_top.pack(pady=10)

# Campo de entrada para matrícula
matricula_entry = tk.Entry(frame_top)
matricula_entry.pack(side=tk.LEFT, padx=5)

# Botão de busca ao lado do campo
btn_buscar = tk.Button(frame_top, text="Buscar", command=lambda: buscar_dados())
btn_buscar.pack(side=tk.LEFT)

# Caixa de texto copiável para exibição dos resultados
resultado_text = tk.Text(root, height=7, width=50, wrap="word")
resultado_text.pack(pady=10)
resultado_text.config(state=tk.DISABLED)  # Impede edição pelo usuário

# Função para buscar os dados
def buscar_dados(event=None):  # Permite ENTER chamar a função
    global df  # Sempre usa a planilha mais atual

    # Recarregar os dados da planilha toda vez que buscar
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        df.columns = df.columns.str.strip()
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao recarregar a planilha: {e}")
        return

    matricula = matricula_entry.get().strip()
    
    if not matricula.isdigit():
        messagebox.showerror("Erro", "Digite apenas números na matrícula.")
        return

    # Buscar a linha correspondente
    resultado = df[df["Matrícula"].astype(str).str.strip() == matricula]

    if resultado.empty:
        texto = "Matrícula não encontrada."
    else:
        dados = resultado.iloc[0]  # Pegamos a primeira ocorrência encontrada
        texto = (
            f"Matrícula: {dados['Matrícula']}\n"
            f"Nome: {dados['Nome']}\n"
            f"Situação: {dados['Situação']}\n"
            f"Data Admissão: {dados['Data de Admissão']}\n"
            f"Data Desligamento: {dados['Data de Desligamento']}\n"
            f"Área: {dados['Área']}\n"
            f"Gestor: {dados['Gestor']}\n"
        )

    # Atualizar a caixa de texto copiável
    resultado_text.config(state=tk.NORMAL)
    resultado_text.delete("1.0", tk.END)
    resultado_text.insert(tk.END, texto)
    resultado_text.config(state=tk.DISABLED)

# Atalhos de teclado
root.bind("<Return>", buscar_dados)

root.mainloop()
