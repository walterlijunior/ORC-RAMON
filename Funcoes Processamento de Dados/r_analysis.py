import os
import pandas as pd
import rpy2.robjects as ro
from rpy2.robjects import pandas2ri
from rpy2.robjects.packages import importr

# Ativa a conversão automática entre pandas DataFrames e R DataFrames
pandas2ri.activate()

# Importa pacotes do R
ggplot2 = importr('ggplot2')
dplyr = importr('dplyr')

def executar_analise(file_path):
    """Executa análises e visualizações em um arquivo Excel."""
    
    # Lê o arquivo Excel
    df = pd.read_excel(file_path, sheet_name='matriz')
    
    if df.empty:
        print("A planilha está vazia. Nenhuma análise a ser feita.")
        return

    # Converte o DataFrame do pandas para o R
    r_df = pandas2ri.py2rpy(df)

    # Define o DataFrame no ambiente R
    ro.globalenv['df'] = r_df

    # Exemplo de análise: resumo estatístico
    ro.r('summary_stats <- summary(df)')
    print("Resumo Estatístico:")
    print(ro.r('summary_stats'))

    # Exemplo de gráfico usando ggplot2
    plot_gráfico(file_path)

def plot_gráfico(file_path):
    """Cria um gráfico e salva como imagem."""
    
    # Define o caminho para salvar a imagem
    img_path = os.path.splitext(file_path)[0] + '_grafico.png'

    # Comando R para criar o gráfico
    r_code = f"""
    library(ggplot2)
    ggplot(df, aes(x = Função, y = `Dotação Inicial`)) +
        geom_bar(stat='identity') +
        labs(title='Gráfico de Dotação Inicial por Função', x='Função', y='Dotação Inicial') +
        theme_minimal()
    ggsave('{img_path}')
    """
    
    # Executa o código R
    ro.r(r_code)

    print(f"Gráfico salvo em: {img_path}")

if __name__ == "__main__":
    file_path = "C:/Users/Acer Aspire/Documents/testes/Balancete de despesa 25 08 24-2.xlsx"
    executar_analise(file_path)
