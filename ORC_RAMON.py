# Importando as bibliotecas
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import filedialog
from tkinter import Tk

# Imagem de texto
print(r"""

                              ___  ____   ____      ____      _    __  __  ___  _   _   ____    ___  
                             / _ \|  _ \ / ___|    |  _ \    / \  |  \/  |/ _ \| \ | | |___ \  / _ \ 
                            | | | | |_) | |   _____| |_) |  / _ \ | |\/| | | | |  \| |   __) || | | |
                            | |_| |  _ <| |__|_____|  _ <  / ___ \| |  | | |_| | |\  |  / __/ | |_| |
                             \___/|_| \_\\____|    |_| \_\/_/   \_\_|  |_|\___/|_| \_| |_____(_)___/             

            

                                    #-.        ..          .::.          ..         :*
                                    -:==-:-==+***#@%+.  :*@@@@@@#:   =#@#***+==-:-==:-
                                      -=-===+==:.  .=#*+@@%=::=#@@*+%=.  .:==+==-=-=  
                                        :=--::::-=:   #@@=      -@@%   .--:::::=-:    
                                           :::++=::=-*%%*        +%%*-=::==+::-       
                                                .=++%@@@@%#+  =#%@@@@%++=.            
                                               -@@@@@@@@@@######@@@@@@@@@@+           
                                               .-------#@@%*#@*%@@*-------:           
                                                -+::==-+%@@%@++@@#.+:+-:-+.           
                                               - :-#*-  ::......::   =*#::.-          
                                             .- ::.     :@@@@@%%%-      .:  -         
                                             = :.        *@@@+  :         -  -        
                                            .=.+         =%%%+..-         .=.+        
                                             *::-        +======+.        =.:=        
                                             .=::-.      .@@@#::-       :-.:=         
                                               --::-=:    *++- ..   .:=-::=-          
                                                .:-: .-:=.+++- :===:= .--:            
                                                   .-*:=- +++- : .:==*:               
                                                    -: :==#**- -=--. :-               
                                                   =: :=  .**= :   =..:-              
                                                   -=.:+   **- :  .+::=:              
                                                    -=: +-:**=::::=:.+:               
                                                     .+*:-. ::.-::=+:                 
                                                     +:.:+=+*-+=:-..=                 
                                                     +.=:  -+--  .+.+                 
                                                     :=--:=-@*-::---:                 
                                                       :====*+++-+.                   
                                                       :--: **. :--:                  



""")
# Solicitando ao usuário se deseja iniciar o trabalho de análise
start_analysis = input("Você deseja iniciar o trabalho de análise? (Sim/Não) ")

if start_analysis.lower() == 'sim':
    # Solicitando o caminho do arquivo ao usuário
    root = Tk()
    root.withdraw() # we don't want a full GUI, so keep the root window from appearing
    file_path = filedialog.askopenfilename() # show an "Open" dialog box and return the path to the selected file

    # Importar o arquivo excel
    df = pd.read_excel(file_path)

    # Solicitando o nome do arquivo de saída ao usuário
    output_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])

    # Criando um escritor de Excel com 'openpyxl'
    writer = pd.ExcelWriter(output_file_path, engine='openpyxl')

    while True:
        # Mostrando ao usuário as colunas que têm valores
        print("As colunas do arquivo são: ")
        for col in df.columns.tolist():
            print(col)

        # Solicitando as colunas para manter
        cols_to_keep = input("Por Favor, insira as colunas para manter (Exemplo: Função,Dotação Inicial): ").split(',')

        # Excluindo todas as outras colunas
        df_temp = df[cols_to_keep]

        # Solicitando a coluna para agrupar 
        group_by_col = input("Por Favor, insira a coluna para Agrupar: ")

        # Agrupar a coluna escolhida e calcular a soma de "Dotação Inicial"
        df_grouped = df_temp.groupby(group_by_col)['Dotação Inicial'].sum().reset_index()

        # Solicitando a ordem de classificação
        sort_order = input("Você deseja classificar em ordem Crescente ou Decrescente? ")

        # Classificar em ordem escolhida
        df_grouped = df_grouped.sort_values('Dotação Inicial', ascending=(sort_order.lower() == 'Crescente'))

        # Calcular o total da 'Dotação Inicial'
        total_dotation = df_grouped['Dotação Inicial'].sum()

        # Criar uma nova coluna de Percentual que é a porcentagem que cada 'Dotação Inicial' representa do total
        df_grouped['Percentual'] = (df_grouped['Dotação Inicial'] / total_dotation) * 100

        # Escrever o DataFrame para uma planilha Excel
        df_grouped.to_excel(writer, sheet_name=group_by_col, index=False)

        # Perguntando ao usuário se deseja continuar
        continue_export = input("Você deseja continuar exportando dados para outra planilha? (Sim/Não) ")

        if continue_export.lower() != 'sim':
            break

    # Salvar o arquivo Excel
    writer.book.save(output_file_path)

