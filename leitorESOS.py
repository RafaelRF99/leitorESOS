import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook
import pytesseract
import cv2
import numpy as np
import os

# Variável global para armazenar os dados da planilha
dados_indexados = {}

# Função para carregar os dados na memória ao iniciar o programa
def carregar_dados():
    global dados_indexados
    try:
        # Caminho correto para o arquivo
        arquivo_excel = r"C:\Program Files\leitorESOS\PPH-Rendimento.xlsm"
        
        # Verificando se o arquivo existe
        if not os.path.exists(arquivo_excel):
            messagebox.showerror("Erro", "Arquivo Excel não encontrado!")
            return
        
        # Carregar o arquivo Excel
        wb = load_workbook(arquivo_excel, data_only=True)
        aba = wb["Rendimento programa"]

        # Criar um dicionário indexado pela coluna "CODIGO"
        dados_indexados = {}
        for row in aba.iter_rows(min_row=2, max_row=aba.max_row, values_only=True):
            codigo = str(row[0]).strip() if row[0] else None  # Coluna CODIGO
            if codigo:
                if codigo not in dados_indexados:
                    dados_indexados[codigo] = []
                dados_indexados[codigo].append((row[4], row[0], row[1], row[2], row[3]))  # LINHA, CODIGO, RAIO, MP, REND

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar os dados do Excel: {e}")

# Função para buscar os dados rapidamente no dicionário
def buscar_dados(codigo):
    # Retorna todos os registros encontrados com esse código
    return dados_indexados.get(codigo, None)

# Função para exibir os dados corretamente formatados
def on_enter(event=None):
    codigo = entry_codigo.get().strip()

    # Fecha o programa se digitar 'exit'
    if codigo.lower() == 'exit':
        root.destroy()
        return

    if codigo:  # Sem validar a quantidade de caracteres
        resultados = buscar_dados(codigo)

        texto_resultado.delete(1.0, tk.END)  # Limpar área de resultado

        if resultados:
            col_sizes = [15, 15, 15, 10, 10]  # Ajustar tamanhos para acomodar os novos campos

            # Criar o cabeçalho incluindo os campos na ordem correta
            header = f"{'LINHA':^{col_sizes[0]}}{'CODIGO':^{col_sizes[1]}}{'RAIO':^{col_sizes[2]}}{'MP':^{col_sizes[3]}}{'REND':^{col_sizes[4]}} \n"
            texto_resultado.insert(tk.END, header)
            texto_resultado.insert(tk.END, "-" * sum(col_sizes) + "\n")

            # Adicionar os resultados
            for resultado in resultados:
                linha_formatada = f"{str(resultado[0]):^{col_sizes[0]}}{str(resultado[1]):^{col_sizes[1]}}{str(resultado[2]):^{col_sizes[2]}}{str(resultado[3]):^{col_sizes[3]}}{str(resultado[4]):^{col_sizes[4]}} \n"
                texto_resultado.insert(tk.END, linha_formatada)

        else:
            texto_resultado.insert(tk.END, "Código não encontrado.")

# Função para limitar a entrada de texto
def limitar_entrada(event):
    codigo = entry_codigo.get().strip()

    if codigo.lower() == "exit":
        on_enter()
        return

    # Removemos as validações de limite de caracteres (não há limite agora)
    on_enter()

# Função para capturar o número da câmera e preencher no campo de entrada
def capturar_numero_da_camera():
    # Abertura da câmera com índice 1 (ajustado conforme seu dispositivo)
    cap = cv2.VideoCapture(1)  # Altere para o índice correto da sua câmera Trust USB

    if not cap.isOpened():
        messagebox.showerror("Erro", "Câmera não encontrada!")
        return

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        # Exibindo a imagem da câmera
        cv2.imshow("Captura de Imagem - Pressione Enter para Congelar", frame)

        # Aguardar pressionamento da tecla "Enter" para congelar a imagem
        key = cv2.waitKey(1) & 0xFF
        if key == ord('\r'):  # Se pressionar Enter (caractere '\r')
            # Congelar a imagem para seleção
            cv2.destroyAllWindows()
            # Seleção de ROI (Região de Interesse) com o mouse
            r = cv2.selectROI("Selecione o número", frame)
            if r != (0, 0, 0, 0):  # Se uma área for selecionada
                x, y, w, h = r
                roi = frame[y:y + h, x:x + w]

                # Usando o Tesseract para extrair texto da área selecionada
                numero_detectado = pytesseract.image_to_string(roi).strip()

                # Preencher o campo de entrada com o número detectado
                entry_codigo.delete(0, tk.END)
                entry_codigo.insert(0, numero_detectado)

                # Realiza a busca no Excel com o número detectado
                on_enter()

            break

    cap.release()
    cv2.destroyAllWindows()

# Configuração da interface gráfica
root = tk.Tk()
root.title("Busca no Excel")
root.geometry("900x600")  # Aumentando o tamanho da janela

# Caixa de entrada
tk.Label(root, text="Digite o código:").pack(pady=10)  # Removido o "11 caracteres"
entry_codigo = tk.Entry(root, font=("Arial", 14), width=20)
entry_codigo.pack(pady=10)
entry_codigo.focus()

entry_codigo.bind("<KeyRelease>", limitar_entrada)

# Área de texto para exibir o resultado
texto_resultado = tk.Text(root, font=("Courier", 12), width=80, height=15)
texto_resultado.pack(pady=20)

# Botão para capturar o número da câmera
botao_camera = tk.Button(root, text="Capturar Número da Câmera", font=("Arial", 14), command=capturar_numero_da_camera)
botao_camera.pack(pady=20)

# Carregar os dados na inicialização
carregar_dados()

root.mainloop()
