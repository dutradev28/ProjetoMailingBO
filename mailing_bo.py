import tkinter as tk
from tkinter import filedialog
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import time
import os
import datetime


def importar_para_excel():
    if arquivo_txt:
        try:
            # Leia o arquivo de texto para um DataFrame do pandas com codificação 'latin1' e delimiter ';'
            df = pd.read_csv(arquivo_txt, delimiter=';', encoding='latin1', on_bad_lines='skip')

            # Adicione uma nova coluna "OI_CHAMADO" após a coluna "DATA_CRIACAO"
            df.insert(df.columns.get_loc("DATA_CRIACAO") + 1, "OI_CHAMADO", "")

            # Adicione uma nova coluna "CPF/CNPJ" após a coluna "STATUS_PEDIDO"
            df.insert(df.columns.get_loc("STATUS_PEDIDO") + 1, "CPF/CNPJ", "")
            
            # Remover linhas com valores vazios na coluna "CODIGO_REJEICAO"
            df = df.dropna(subset=["CODIGO_REJEICAO"])
            
            # Classificar o DataFrame com base na coluna "NUMERO_ATIVO" em ordem alfabética
            df = df.sort_values(by=["NUMERO_ATIVO"])
            
            # Verifique se a coluna 'Unnamed: 20' existe antes de tentar removê-la
            if 'Unnamed: 20' in df.columns:
                df.drop(columns=['Unnamed: 20'], inplace=True)
                
            # Excluir linhas em que a coluna "TIPO_ATIVIDADE" contém "Port. Doadora" 
            df = df[~df["TIPO_ATIVIDADE"].str.contains("Port. Doadora")]

            # Defina o nome do arquivo CSV de destino
            arquivo_csv = arquivo_txt.replace('.txt', '.csv')

            # Salvar o DataFrame como arquivo .csv
            arquivo_csv = arquivo_txt.replace('.txt', '.csv')
            df.to_csv(arquivo_csv, index=False, sep=';', encoding='latin1')

            # Renomear o arquivo .csv conforme a data atual
            data_atual = datetime.datetime.now()
            novo_nome = f'BO PORTABILIDADE - Base Fibra_{data_atual.strftime("%Y%m%d")}.csv'

            # Verificar se o novo nome já existe e adicionar um número sequencial se necessário
            contador = 1
            while os.path.exists(novo_nome):
                novo_nome = f'BO PORTABILIDADE - Base Fibra_{data_atual.strftime("%Y%m%d")}_{contador}.csv'
                contador += 1

            os.rename(arquivo_csv, novo_nome)

            # Configurar a mensagem de sucesso
            mensagem_sucesso_animada(arquivo_csv, sucesso=True)
        except Exception as e:
            mostrar_mensagem(str(e), sucesso=False)

# Função para selecionar o arquivo
def selecionar_arquivo():
    global arquivo_txt
    # Abrir a janela de seleção de arquivo
    arquivo_txt = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt")])
# Função para mostrar a mensagem de erro com quebras de linha
def mostrar_mensagem(mensagem, sucesso=False):
    output_text.config(state="normal")  # Ativar o widget Text
    output_text.delete(1.0, tk.END)  # Limpar o conteúdo anterior

    # Dividir a mensagem em linhas e adicionar quebras de linha entre elas
    linhas_mensagem = mensagem.splitlines()
    mensagem_formatada = "\n".join(linhas_mensagem)

    output_text.insert(tk.END, mensagem_formatada)  # Inserir a mensagem formatada no widget Text
    output_text.config(state="disabled")  # Desativar o widget Text
    output_text.yview_moveto(1.0)  # Rolar para o final da mensagem

    # Definir a cor do texto com base no sucesso ou erro
    if sucesso:
        output_text.config(fg="green")
    else:
        output_text.config(fg="red")

# Função para mostrar a mensagem de sucesso com quebras de linha
def mensagem_sucesso_animada(arquivo, sucesso=False):
    mensagem = f'O arquivo foi importado para {arquivo} com sucesso.' if sucesso else f'Ocorreu um erro ao importar o arquivo para {arquivo}.'

    # Chame a função mostrar_mensagem com base no sucesso ou erro
    mostrar_mensagem(mensagem, sucesso)

# Função para configurar a barra de rolagem horizontal
def configurar_barra_de_rolagem_horizontal(*args):
    output_text.xview(*args)

# Configuração da janela
root = tk.Tk()
root.title("BO PORTABILIDADE-BASE FIBRA EMP")

# Definir o tamanho da janela como 500x500 pixels
root.geometry("500x500")

# Configurar a fonte tkinter para incluir todos os caracteres
font = ("Arial", 12)
root.option_add("*Font", font)

# Botão para selecionar arquivo
select_button = tk.Button(root, text="Selecionar Arquivo .txt", command=selecionar_arquivo)
select_button.pack(pady=20)

# Botão "OK" para iniciar a importação
ok_button = tk.Button(root, text="Iniciar Processo", command=importar_para_excel)
ok_button.pack(pady=10)

# Widget Text
output_text = tk.Text(root, wrap="none", height=6)
output_text.pack(pady=10, padx=10, fill="both", expand=True)

# Barra de rolagem horizontal personalizada
scrollbar_x = tk.Scrollbar(root, orient="horizontal", command=configurar_barra_de_rolagem_horizontal)
scrollbar_x.pack(fill="x")

output_text.config(xscrollcommand=scrollbar_x.set)

# Iniciar a janela principal
root.mainloop()
