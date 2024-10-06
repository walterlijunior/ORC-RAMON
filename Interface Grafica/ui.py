import os
import sys
import logging
import shutil

# Adicionando o caminho do diretório onde está o módulo
sys.path.append(r"C:\Users\Acer Aspire\Documents\MeusProjetos\ORC-RAMON\Funcoes Processamento de Dados")

# Importando as funções do módulo
from data_processing import renomear_planilha, excluir_colunas, consolidar_dados

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

if __name__ == "__main__":
    # Caminho do arquivo original (modifique conforme necessário)
    file_path = "C:/Users/Acer Aspire/Documents/testes/Balancete de despesa 25 08 24-2.xlsx"

    # Chamando a função para renomear a planilha
    renomeada_path = renomear_planilha(file_path)

    if renomeada_path:  # Verifica se a renomeação foi bem-sucedida
        # Criar uma cópia do arquivo renomeado antes de modificar
        copia_path = criar_copia(renomeada_path)
        
        if copia_path:  # Verifica se a cópia foi criada com sucesso
            excluir_colunas(renomeada_path)  # Exclui colunas na própria planilha
            consolidar_dados(renomeada_path)  # Consolida os dados na própria planilha
        else:
            logging.error("A cópia do arquivo não foi criada. As funções de exclusão e consolidação não serão executadas.")
    else:
        logging.error("A renomeação da planilha falhou. O processo não pode continuar.")
