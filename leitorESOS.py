import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook

# Função para ler os dados do Excel
def buscar_dados(codigo):
    # Carregar o arquivo Excel
    try:
        wb = load_workbook("./leitorESOS.xlsx")  # Substitua pelo caminho correto
        aba1 = wb["DB"]
        aba2 = wb["Rendimento programas"]

        # Procurar o calibrador na primeira aba
        calibrador = None
        for row in aba1.iter_rows(min_row=2, max_row=aba1.max_row):
            if row[0].value == codigo:
                calibrador = row[1].value  # Supondo que a segunda coluna tenha o calibrador
                break

        if calibrador is None:
            return None  # Não encontrou o calibrador na primeira aba

        # Agora procurar os dados na segunda aba
        resultados = []
        for row in aba2.iter_rows(min_row=2, max_row=aba2.max_row):
            if row[0].value == calibrador:
                linha = row[4].value
                raio = row[1].value
                mp = row[2].value
                rend = row[3].value
                resultados.append((linha, calibrador, raio, mp, rend))

        return resultados

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")
        return None

# Função para lidar com a entrada de dados
def on_enter(event=None):
    codigo = entry_codigo.get()
    if len(codigo) == 11:  # Verifica se tem 11 caracteres
        resultados = buscar_dados(codigo)
        if resultados:
            texto_resultado.delete(1.0, tk.END)  # Limpar área de resultado
            for resultado in resultados:
                texto_resultado.insert(tk.END, f"Linha: {resultado[0]}, Calibrador: {resultado[1]}, Raio: {resultado[2]}, MP: {resultado[3]}, REND: {resultado[4]}\n")
        else:
            texto_resultado.delete(1.0, tk.END)
            texto_resultado.insert(tk.END, "Código não encontrado.")
    else:
        messagebox.showwarning("Aviso", "Digite exatamente 11 caracteres.")

# Configuração da interface gráfica
root = tk.Tk()
root.title("Busca no Excel")
root.geometry("600x400")

# Caixa de entrada para o código
tk.Label(root, text="Digite o código (11 caracteres):").pack(pady=10)
entry_codigo = tk.Entry(root, font=("Arial", 14), width=20)
entry_codigo.pack(pady=10)
entry_codigo.bind("<Return>", on_enter)  # Quando pressionar Enter

# Área de texto para exibir o resultado
texto_resultado = tk.Text(root, font=("Arial", 12), width=70, height=10)
texto_resultado.pack(pady=20)

# Botão de buscar (opcional, já estamos usando o Enter)
# tk.Button(root, text="Buscar", command=on_enter).pack(pady=10)

root.mainloop()
