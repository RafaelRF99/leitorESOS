import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook

# Variável global para armazenar os dados da planilha
dados_indexados = {}

# Função para carregar os dados na memória ao iniciar o programa
def carregar_dados():
    global dados_indexados
    try:
        wb = load_workbook(r"C:\Program Files\leitorESOS\PPH-Rendimento.xlsm", data_only=True)
        aba = wb["Rendimento programa"]

        # Criar um dicionário indexado pela coluna F (SD3010.Num OP)
        dados_indexados = {}
        for row in aba.iter_rows(min_row=2, max_row=aba.max_row, values_only=True):
            codigo = str(row[5]).strip() if row[5] else None  # Coluna F (SD3010.Num OP)
            if codigo:
                dados_indexados[codigo] = (row[4], row[0], row[1], row[2], row[3], row[5])  # Linha, Código, Raio, MP, REND, SD3010.Num OP

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar os dados do Excel: {e}")

# Função para buscar os dados rapidamente no dicionário
def buscar_dados(codigo):
    if codigo in dados_indexados:
        linha, codigo, raio, mp, rend, num_op = dados_indexados[codigo]
        return [(linha, codigo, raio, mp, rend, num_op)]  # Alterado para retornar a nova ordem dos campos
    return None

# Função para exibir os dados corretamente formatados
def on_enter(event=None):
    codigo = entry_codigo.get().strip()

    # Fecha o programa se digitar 'exit'
    if codigo.lower() == 'exit':
        root.destroy()
        return

    if len(codigo) == 11:
        resultados = buscar_dados(codigo)

        if resultados:
            texto_resultado.delete(1.0, tk.END)  # Limpar área de resultado
            col_sizes = [15, 15, 15, 10, 10, 15]  # Ajustar tamanhos para acomodar os novos campos

            # Criar o cabeçalho incluindo os campos na ordem correta
            header = f"{'Linha':^{col_sizes[0]}}{'Código':^{col_sizes[1]}}{'Raio':^{col_sizes[2]}}{'MP':^{col_sizes[3]}}{'REND':^{col_sizes[4]}}{'SD3010.Num OP':^{col_sizes[5]}} \n"
            texto_resultado.insert(tk.END, header)
            texto_resultado.insert(tk.END, "-" * sum(col_sizes) + "\n")

            # Adicionar os resultados
            for resultado in resultados:
                linha_formatada = f"{str(resultado[0]):^{col_sizes[0]}}{str(resultado[1]):^{col_sizes[1]}}{str(resultado[2]):^{col_sizes[2]}}{str(resultado[3]):^{col_sizes[3]}}{str(resultado[4]):^{col_sizes[4]}}{str(resultado[5]):^{col_sizes[5]}} \n"
                texto_resultado.insert(tk.END, linha_formatada)

        else:
            texto_resultado.delete(1.0, tk.END)
            texto_resultado.insert(tk.END, "Código não encontrado.")

# Função para limitar a entrada de texto
def limitar_entrada(event):
    codigo = entry_codigo.get().strip()

    if codigo.lower() == "exit":
        on_enter()
        return

    if len(codigo) > 11:
        entry_codigo.delete(0, tk.END)
        entry_codigo.insert(0, codigo[-1])

    if len(codigo) == 11:
        if codigo[10] == "2":
            entry_codigo.delete(0, tk.END)
            entry_codigo.insert(0, codigo[:10] + "1")
        on_enter()

# Configuração da interface gráfica
root = tk.Tk()
root.title("Busca no Excel")
root.geometry("900x600")  # Aumentando o tamanho da janela

# Caixa de entrada
tk.Label(root, text="Digite o código (11 caracteres):").pack(pady=10)
entry_codigo = tk.Entry(root, font=("Arial", 14), width=20)
entry_codigo.pack(pady=10)
entry_codigo.focus()

entry_codigo.bind("<KeyRelease>", limitar_entrada)

# Área de texto para exibir o resultado
texto_resultado = tk.Text(root, font=("Courier", 12), width=80, height=15)
texto_resultado.pack(pady=20)

# Carregar os dados na inicialização
carregar_dados()

root.mainloop()
