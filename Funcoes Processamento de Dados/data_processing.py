import pandas as pd
import os
import logging

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def renomear_planilha(file_path):
    """Renomeia a planilha do arquivo Excel para 'matriz' se necessário."""
    if not os.path.exists(file_path) or not file_path.endswith(('.xlsx', '.xls')):
        logging.error(f"Arquivo inválido: {file_path}")
        return None

    try:
        with pd.ExcelFile(file_path) as xls:
            if 'matriz' not in xls.sheet_names:
                df = pd.read_excel(xls, xls.sheet_names[0])  # Lê a primeira planilha
                # Salva diretamente no mesmo arquivo, renomeando a planilha
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='matriz', index=False)
                logging.info(f"Planilha renomeada para 'matriz'.")
                return file_path
            else:
                logging.info("Nenhuma alteração feita; o nome da planilha já é 'matriz'.")
                return file_path

    except Exception as e:
        logging.error(f"Erro ao processar o arquivo: {e}")
        return None

def excluir_colunas(file_path):
    """Exclui colunas da planilha, mantendo apenas 'Função' e 'Dotação Inicial'."""
    try:
        df = pd.read_excel(file_path, sheet_name='matriz')

        # Verifica se as colunas esperadas existem
        if 'Função' not in df.columns or 'Dotação Inicial' not in df.columns:
            logging.error("As colunas 'Função' e 'Dotação Inicial' não foram encontradas na planilha.")
            return

        if df.empty:
            logging.warning("A planilha está vazia. Nenhuma coluna a ser excluída.")
            return

        df = df[['Função', 'Dotação Inicial']]
        df.to_excel(file_path, sheet_name='matriz', index=False)
        logging.info("Colunas excluídas, apenas 'Função' e 'Dotação Inicial' mantidas.")
        
    except Exception as e:
        logging.error(f"Erro ao excluir colunas: {e}")

def consolidar_dados(file_path):
    """Consolida os dados na coluna 'Função' da planilha."""
    try:
        df = pd.read_excel(file_path, sheet_name='matriz')

        if df.empty:
            logging.warning("A planilha está vazia. Nenhum dado a ser consolidado.")
            return

        df_consolidado = df.groupby('Função', as_index=False).agg({'Dotação Inicial': 'sum'})
        df_consolidado.to_excel(file_path, sheet_name='matriz', index=False)
        logging.info("Dados consolidados na coluna 'Função'.")
        
    except Exception as e:
        logging.error(f"Erro ao consolidar dados: {e}")

# Fluxo principal do programa
if __name__ == "__main__":
    # Caminho do arquivo original (modifique conforme necessário)
    file_path = "C:/Users/Acer Aspire/Documents/testes/Balancete de despesa 25 08 24-2.xlsx"
    
    # Chamando a função para renomear a planilha
    renomeada_path = renomear_planilha(file_path)

    if renomeada_path:  # Verifica se a renomeação foi bem-sucedida
        excluir_colunas(renomeada_path)  # Exclui colunas na própria planilha
        consolidar_dados(renomeada_path)  # Consolida os dados na própria planilha
    else:
        logging.error("A renomeação da planilha falhou. O processo não pode continuar.")
