import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook

# Caminho correto do Excel
EXCEL_PATH = r"\\srv4flabeg\DOCS2\ENGENHARIA\GERENCIAMENTO DE PROGRAMAS\Tabela link.xlsm"
SHEET_NAME = "Rendimento programa"

# Variável global para armazenar os resultados e código de barra
resultados = []
ultimo_codigo = ""

def buscar_dados(codigo_barra):
    try:
        wb = load_workbook(EXCEL_PATH, data_only=True)
        aba = wb[SHEET_NAME]

        codigo_encontrado = None
        resultados_local = []

        for row in aba.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 6 and str(row[5]).strip() == codigo_barra.strip():
                codigo_encontrado = row[0]
                break

        if not codigo_encontrado:
            return None

        for row in aba.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 6 and str(row[0]).strip() == str(codigo_encontrado).strip():
                resultados_local.append(row)

        return resultados_local if resultados_local else None

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")
        return None

def filtrar_resultado(resultados, filtro):
    if filtro:
        resultados_filtrados = []
        for row in resultados:
            if any(filtro.lower() in str(cell).lower() for cell in row):
                resultados_filtrados.append(row)
        return resultados_filtrados
    return resultados

def exibir_resultado(resultados):
    texto_resultado.delete(1.0, tk.END)
    if resultados:
        col_sizes = [10, 15, 15, 10, 10]
        header = f"{'Linha':^{col_sizes[0]}}{'Código':^{col_sizes[1]}}{'Raio':^{col_sizes[2]}}{'MP':^{col_sizes[3]}}{'REND':^{col_sizes[4]}}\n"
        texto_resultado.insert(tk.END, header, "header")
        texto_resultado.insert(tk.END, "-" * sum(col_sizes) + "\n")

        for i, row in enumerate(resultados):
            linha_formatada = f"{str(row[4]):^{col_sizes[0]}}{str(row[0]):^{col_sizes[1]}}{str(row[1]):^{col_sizes[2]}}{str(row[2]):^{col_sizes[3]}}{str(row[3]):^{col_sizes[4]}}\n"
            tag = "even" if i % 2 == 0 else "odd"
            texto_resultado.insert(tk.END, linha_formatada, tag)
    else:
        texto_resultado.insert(tk.END, "Código não encontrado.")

def on_enter(event=None):
    global ultimo_codigo
    codigo = codigo_var.get().strip()
    if len(codigo) != 11:
        return
    if codigo.lower() == 'exit':
        root.destroy()
        return
    if codigo.lower() == "help":
        texto_resultado.delete(1.0, tk.END)
        texto_resultado.insert(tk.END, "Criado e desenvolvido por: Rafael Ferreira\n")
        texto_resultado.insert(tk.END, "LinkedIn: https://www.linkedin.com/in/rafael-ferreira99/\n")
        texto_resultado.insert(tk.END, "WhatsApp: (19) 998255728\n")
        return
    # Salva o último código digitado
    ultimo_codigo = codigo
    global resultados
    resultados = buscar_dados(codigo)
    exibir_resultado(resultados)
    codigo_var.set("")
    entry_codigo.focus()

def filtrar_por_termo(*args):
    filtro = filtro_var.get()
    if ultimo_codigo:  # Se houver código de barra salvo, realiza a busca com o filtro
        resultados_filtrados = filtrar_resultado(resultados, filtro)
        exibir_resultado(resultados_filtrados)

def limitar_tamanho(*args):
    texto = codigo_var.get()
    if len(texto) > 11:
        codigo_var.set(texto[:11])
    elif len(texto) == 11:
        root.after(100, on_enter)

root = tk.Tk()
root.title("Busca no Excel")
root.geometry("800x600")

# Adiciona a Label com texto explicativo para código
tk.Label(root, text="Digite o código ou 'help':", font=("Arial", 12)).pack(pady=10)

# Criando o Entry para código com borda
codigo_var = tk.StringVar()
codigo_var.trace_add("write", limitar_tamanho)
entry_codigo = tk.Entry(root, font=("Arial", 14), textvariable=codigo_var, width=20, bd=2, relief="solid", justify="center")
entry_codigo.pack(pady=10)
entry_codigo.focus()

# Adiciona o campo de filtro pequeno, alinhado à esquerda, com o texto "Linha:"
frame_filtro = tk.Frame(root)
frame_filtro.pack(pady=10, anchor="w", padx=10)
tk.Label(frame_filtro, text="Linha:", font=("Arial", 12)).pack(side=tk.LEFT)
filtro_var = tk.StringVar()
filtro_var.trace_add("write", filtrar_por_termo)  # Chama a função de filtro ao digitar
entry_filtro = tk.Entry(frame_filtro, font=("Arial", 14), textvariable=filtro_var, width=8, bd=2, relief="solid", justify="center")  # Reduzido ainda mais
entry_filtro.pack(side=tk.LEFT)

# Texto de resultado (não é editável)
texto_resultado = tk.Text(root, font=("Courier", 12), width=70, height=15, wrap=tk.WORD)
texto_resultado.pack(pady=20)
texto_resultado.tag_configure("even", background="#F0F0F0")
texto_resultado.tag_configure("odd", background="#D0D0D0")
texto_resultado.tag_configure("header", font=("Courier", 12, "bold"))

# Label com texto de ajuda
tk.Label(root, text="Caso dificuldade, digitar 'help'", font=("Arial", 10)).pack(side=tk.BOTTOM, fill="x", padx=20)

root.mainloop()
