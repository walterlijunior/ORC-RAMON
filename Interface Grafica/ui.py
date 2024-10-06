import os
import sys
import logging
import shutil
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

# Adicionando o caminho do diretório onde está o módulo
sys.path.append(r"C:\Users\Acer Aspire\Documents\MeusProjetos\ORC-RAMON\Funcoes Processamento de Dados")

# Importando as funções do módulo
from data_processing import renomear_planilha, consolidar_dados

def criar_copia(file_path):
    """Cria uma cópia do arquivo Excel na mesma pasta, com um sufixo '_copia'."""
    base, ext = os.path.splitext(file_path)  # Separar o nome do arquivo e a extensão
    copia_path = f"{base}_copia{ext}"  # Novo caminho para a cópia

    # Faz uma cópia do arquivo
    try:
        shutil.copyfile(file_path, copia_path)
        logging.info(f"Cópia criada: {copia_path}")
        return copia_path
    except Exception as e:
        logging.error(f"Erro ao criar cópia do arquivo: {e}")
        return None

def selecionar_colunas(file_path):
    """Mostra uma janela para selecionar quais colunas excluir da planilha."""
    try:
        df = pd.read_excel(file_path, sheet_name='matriz')

        # Verifica se a planilha está vazia
        if df.empty:
            logging.warning("A planilha está vazia. Nenhuma coluna a ser excluída.")
            return []

        # Cria a janela de seleção
        root = tk.Tk()
        root.title("Selecionar Colunas a Excluir")

        # Define o tamanho padrão da janela
        root.geometry('950x700')
        root.config(bg='#f0f0f0')

        # Carregar e definir o ícone da janela (logo)
        img = ImageTk.PhotoImage(Image.open(r"C:\Users\Acer Aspire\Documents\MeusProjetos\ORC-RAMON\icon\icon-ui\orc-icon.png"))
        root.iconphoto(False, img)

        # Estilos de fonte
        fonte_padrao = ("Segoe UI", 12)
        fonte_botoes = ("Segoe UI", 14, "bold")

        # Cria um frame principal para adaptar os elementos ao redimensionamento
        frame = tk.Frame(root, bg='#f0f0f0')
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Cria um rótulo explicativo
        label = tk.Label(frame, text="Selecione as colunas que deseja excluir:", font=("Segoe UI", 14), bg='#f0f0f0')
        label.pack(pady=10)

        # Cria um Frame para a Listbox e a barra de rolagem
        listbox_frame = tk.Frame(frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Adiciona a barra de rolagem
        scrollbar = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Cria a Listbox para exibir as colunas
        listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, width=60, height=20, font=fonte_padrao,
                             selectbackground='#ADD8E6', selectforeground='black', relief="solid", bd=2,
                             yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        # Adiciona as colunas à Listbox
        for col in df.columns:
            listbox.insert(tk.END, col)

        # Função para selecionar todas as colunas
        def selecionar_todas():
            listbox.select_set(0, tk.END)  # Seleciona todas as colunas

        # Função para desmarcar todas as colunas
        def desmarcar_todas():
            listbox.select_clear(0, tk.END)  # Desmarca todas as colunas

        # Função para obter colunas desmarcadas (mantidas)
        def obter_colunas_desmarcadas():
            todas_colunas = listbox.get(0, tk.END)
            desmarcadas = [todas_colunas[i] for i in range(len(todas_colunas)) if i not in listbox.curselection()]
            if not desmarcadas:
                messagebox.showinfo("Nenhuma seleção", "Nenhuma coluna foi desmarcada para manter.")
            root.destroy()  # Fecha a janela
            return desmarcadas

        # Cria um Frame para os botões
        button_frame = tk.Frame(frame, bg='#f0f0f0')
        button_frame.pack(pady=10)

        # Botão para selecionar todas as colunas
        select_all_button = tk.Button(button_frame, text="Selecionar Todas", font=fonte_botoes, bg='#ADD8E6', fg='black',
                                      relief="raised", bd=3, command=selecionar_todas)
        select_all_button.pack(side=tk.LEFT, padx=10)

        # Botão para desmarcar todas as colunas
        deselect_all_button = tk.Button(button_frame, text="Desmarcar Todas", font=fonte_botoes, bg='#ADD8E6', fg='black',
                                        relief="raised", bd=3, command=desmarcar_todas)
        deselect_all_button.pack(side=tk.LEFT, padx=10)

        # Botão para confirmar a seleção
        confirm_button = tk.Button(button_frame, text="Confirmar", font=fonte_botoes, bg='#32CD32', fg='white',
                                   relief="raised", bd=3, command=lambda: obter_colunas_desmarcadas())
        confirm_button.pack(side=tk.LEFT, padx=10)

        # Inicia o loop da interface gráfica
        root.mainloop()

        return obter_colunas_desmarcadas()

    except Exception as e:
        logging.error(f"Erro ao selecionar colunas: {e}")
        return []

def excluir_colunas(file_path, colunas_excluir):
    """Exclui colunas da planilha com base na seleção do usuário."""
    try:
        df = pd.read_excel(file_path, sheet_name='matriz')

        # Verifica se as colunas esperadas existem
        if 'Função' not in df.columns or 'Dotação Inicial' not in df.columns:
            logging.error("As colunas 'Função' e 'Dotação Inicial' não foram encontradas na planilha.")
            return

        if df.empty:
            logging.warning("A planilha está vazia. Nenhuma coluna a ser excluída.")
            return

        # Mantém apenas as colunas que foram desmarcadas pelo usuário
        colunas_a_manter = ['Função', 'Dotação Inicial'] + colunas_excluir
        df = df[colunas_a_manter]
        df.to_excel(file_path, sheet_name='matriz', index=False)
        logging.info("Colunas excluídas conforme seleção do usuário.")

    except Exception as e:
        logging.error(f"Erro ao excluir colunas: {e}")

if __name__ == "__main__":
    file_path = "C:/Users/Acer Aspire/Documents/testes/Balancete de despesa 25 08 24-2.xlsx"

    renomeada_path = renomear_planilha(file_path)

    if renomeada_path:
        copia_path = criar_copia(renomeada_path)
        
        if copia_path:
            colunas_excluir = selecionar_colunas(renomeada_path)
            if colunas_excluir:
                excluir_colunas(renomeada_path, colunas_excluir)
                consolidar_dados(renomeada_path)
            else:
                logging.error("Nenhuma coluna foi selecionada para exclusão.")
        else:
            logging.error("Falha ao criar a cópia do arquivo.")
    else:
        logging.error("Falha ao renomear a planilha.")
