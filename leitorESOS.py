import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook

# Função para ler os dados do Excel
def buscar_dados(codigo):
    try:
        wb = load_workbook("./leitorESOS.xlsx")
        aba1 = wb["DB"]
        aba2 = wb["Rendimento programas"]

        # Procurar o calibrador na primeira aba
        calibrador = None
        for row in aba1.iter_rows(min_row=2, max_row=aba1.max_row, values_only=True):
            if row[0] == codigo:
                calibrador = row[1]

        if calibrador is None:
            print("Código não encontrado na 1º aba.")
            return None

        # Buscar os dados na segunda aba
        resultados = []
        for row in aba2.iter_rows(min_row=2, max_row=aba2.max_row, values_only=True):
            if row[0] == calibrador:
                linha = row[4]
                raio = row[1]
                mp = row[2]
                rend = row[3]
                resultados.append((linha, calibrador, raio, mp, rend))

        return resultados if resultados else None

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")
        return None

# Função para exibir os dados corretamente formatados
def on_enter(event=None):
    codigo = entry_codigo.get().strip()

    # Força o fechamento da aplicação
    if codigo.lower() == 'exit':
        root.destroy()
        return

    # Se o usuário digitou "help", exibe as informações de contato
    if codigo.lower() == "help":
        texto_resultado.delete(1.0, tk.END)
        texto_resultado.insert(tk.END, "Criado e desenvolvido por: Rafael Ferreira\n")
        texto_resultado.insert(tk.END, "LinkedIn: https://www.linkedin.com/in/rafael-ferreira99/\n")
        texto_resultado.insert(tk.END, "WhatsApp: (19) 998255728\n")
        return

    if len(codigo) == 11:
        resultados = buscar_dados(codigo)

        if resultados:
            texto_resultado.delete(1.0, tk.END)  # Limpar área de resultado
            # Definir o tamanho das colunas
            col_sizes = [10, 15, 15, 10, 10]  # Ajuste conforme necessário

            # Criar o cabeçalho centralizado
            header = f"{'Linha':^{col_sizes[0]}}{'Calibrador':^{col_sizes[1]}}{'Raio':^{col_sizes[2]}}{'MP':^{col_sizes[3]}}{'REND':^{col_sizes[4]}}\n"
            texto_resultado.insert(tk.END, header)
            texto_resultado.insert(tk.END, "-" * sum(col_sizes) + "\n")  # Linha separadora

            # Adicionar os resultados centralizados
            for resultado in resultados:
                linha_formatada = f"{str(resultado[0]):^{col_sizes[0]}}{str(resultado[1]):^{col_sizes[1]}}{str(resultado[2]):^{col_sizes[2]}}{str(resultado[3]):^{col_sizes[3]}}{str(resultado[4]):^{col_sizes[4]}}\n"
                texto_resultado.insert(tk.END, linha_formatada)

        else:
            texto_resultado.delete(1.0, tk.END)
            texto_resultado.insert(tk.END, "Código não encontrado.")

# Função para limitar a entrada de texto, reiniciar ao ultrapassar 11 caracteres e corrigir "2" para "1" na posição 11
def limitar_entrada(event):
    codigo = entry_codigo.get().strip()
    
    if codigo.lower() == "exit":
        on_enter()
        return

    if codigo.lower() == "help":
        on_enter()
        return

    if len(codigo) > 11:
        # Guarda o último caractere digitado
        ultimo_caractere = codigo[-1]
        # Limpa o campo e reinicia com o último caractere
        entry_codigo.delete(0, tk.END)
        entry_codigo.insert(0, ultimo_caractere)
    
    # Se o usuário digitou 11 caracteres
    if len(codigo) == 11:
        # Verifica se o 11º caractere é "2"
        if codigo[10] == "2":
            codigo = codigo[:10] + "1"  # Substitui "2" por "1" na posição 11
            entry_codigo.delete(0, tk.END)
            entry_codigo.insert(0, codigo)

        on_enter()

# Configuração da interface gráfica
root = tk.Tk()
root.title("Busca no Excel")
root.geometry("700x500")

# Caixa de entrada
tk.Label(root, text="Digite o código (11 caracteres) ou 'help':").pack(pady=10)
entry_codigo = tk.Entry(root, font=("Arial", 14), width=20)
entry_codigo.pack(pady=10)
entry_codigo.focus()

entry_codigo.bind("<KeyRelease>", limitar_entrada)

# Área de texto para exibir o resultado
texto_resultado = tk.Text(root, font=("Courier", 12), width=70, height=15)  # Usar fonte monoespaçada melhora alinhamento
texto_resultado.pack(pady=20)

rodape = tk.Label(root, text="Caso dificuldade, digitar 'help'", font=("Arial", 10), anchor="center")
rodape.pack(side=tk.BOTTOM, fill="x", padx=20)

root.mainloop()
