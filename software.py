import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, IntVar, Checkbutton, Scrollbar, Canvas, Frame
import os

# Variáveis globais
colunas_desejadas = []
colunas_encontradas = []
checkboxes_vars = []

# processar o arquivo Excel
def process_excel(input_file, output_file, colunas_selecionadas):
    try:
        df = pd.read_excel(input_file, skiprows=1)
        df_filtrado = df[colunas_selecionadas]
        df_filtrado.to_excel(output_file, index=False)
        messagebox.showinfo("Sucesso", "Arquivo processado e salvo com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

#selecionar o arquivo de entrada
def selecionar_arquivo_entrada():
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if arquivo:
        entrada_var.set(arquivo)
        nome_saida_padrao = os.path.join(os.path.dirname(arquivo), "arquivo_processado.xlsx")
        saida_var.set(nome_saida_padrao)
        
        try:
            df = pd.read_excel(arquivo, skiprows=1)
            global colunas_encontradas
            colunas_encontradas = df.columns.tolist()
            exibir_colunas(colunas_encontradas)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo: {e}")

# exibir as colunas encontradas com checkboxes
def exibir_colunas(colunas):
    for widget in colunas_frame.winfo_children():
        widget.destroy()

    global checkboxes_vars
    checkboxes_vars = []

    # Configuração de scroll para colunas longas
    canvas = Canvas(colunas_frame, bg="#f0f0f0")
    scroll_frame = Frame(canvas, bg="#f0f0f0")
    scrollbar = Scrollbar(colunas_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    
    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    for i, coluna in enumerate(colunas):
        var = IntVar(value=1 if coluna in colunas_desejadas_padrao else 0)
        chk = Checkbutton(scroll_frame, text=coluna, variable=var, onvalue=1, offvalue=0, bg="#f0f0f0")
        chk.grid(row=i, column=0, sticky='w', padx=5, pady=2)
        checkboxes_vars.append((coluna, var))

# local de saída
def selecionar_local_saida():
    arquivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
        title="Salvar arquivo como"
    )
    if arquivo:
        saida_var.set(arquivo)

# iniciar o processamento
def iniciar_processamento():
    input_file = entrada_var.get()
    output_file = saida_var.get()
    
    colunas_selecionadas = [coluna for coluna, var in checkboxes_vars if var.get() == 1]
    
    if input_file and output_file and colunas_selecionadas:
        process_excel(input_file, output_file, colunas_selecionadas)
    else:
        messagebox.showwarning("Aviso", "Selecione o arquivo de entrada, o local de saída e pelo menos uma coluna.")

# Interface gráfica com Tkinter
janela = tk.Tk()
janela.title("Processar Excel")
janela.geometry("750x600")

# Variáveis para armazenar caminhos dos arquivos
entrada_var = tk.StringVar()
saida_var = tk.StringVar()

colunas_desejadas_padrao = [
    'OS', 'ProcessoFilho', 'Atividade', 'Status', 
    'Item / Localidade', 'Solicitante', 'Executante',
    'Competência', 'Criticidade Item', 'Aberta Por'
]

janela.configure(bg="#f0f0f0")

frame = tk.Frame(janela, bg="#f0f0f0")
frame.pack(pady=10)

btn_entrada = tk.Button(frame, text="Selecionar Arquivo de Entrada", command=selecionar_arquivo_entrada, bg="#4CAF50", fg="white", font=("Arial", 12), width=30)
btn_entrada.grid(row=0, column=0, padx=10, pady=10)

entrada_entry = tk.Entry(frame, textvariable=entrada_var, width=50, font=("Arial", 10))
entrada_entry.grid(row=0, column=1, padx=10, pady=10)

btn_saida = tk.Button(frame, text="Selecionar Local de Saída", command=selecionar_local_saida, bg="#2196F3", fg="white", font=("Arial", 12), width=30)
btn_saida.grid(row=1, column=0, padx=10, pady=10)

saida_entry = tk.Entry(frame, textvariable=saida_var, width=50, font=("Arial", 10))
saida_entry.grid(row=1, column=1, padx=10, pady=10)

colunas_frame = tk.Frame(janela, bg="#f0f0f0")
colunas_frame.pack(pady=20, fill="both", expand=True)

btn_processar = tk.Button(janela, text="Processar", command=iniciar_processamento, bg="#FF5722", fg="white", font=("Arial", 14), width=20)
btn_processar.pack(pady=20)

janela.mainloop()
