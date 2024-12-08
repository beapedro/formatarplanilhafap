import os
import pandas as pd

origem = r'\aguardando'
destino = r'\atualizados'
os.makedirs(destino, exist_ok=True)  


pathtask = os.listdir(origem)
numero = 0
for file in pathtask:

    numero +=1 
    if file.endswith('.xlsx'):
        nome_arquivo = file
        caminho_arquivo = os.path.join(origem, nome_arquivo)

        try:

            data = pd.read_excel(caminho_arquivo)
            df = pd.DataFrame(data)
            print(f"Processando arquivo: {nome_arquivo}")
            print(f"Colunas encontradas: {df.columns}")

            colunas_desejadas = ['CNAE', 'Número de Ordem', 'Vigencia']
            if not all(coluna in df.columns for coluna in colunas_desejadas):
                print(f"Arquivo {nome_arquivo} não contém todas as colunas esperadas.")
                continue

            if 'Índice de Frequência' in df.columns:
                df['Indice'] = df['Índice de Frequência']
                df['OrdemFapId'] = 1
            elif 'Índice de Gravidade' in df.columns:
                df['Indice'] = df['Índice de Gravidade']
                df['OrdemFapId'] = 2
            elif 'Índice de Custo' in df.columns:
                df['Indice'] = df['Índice de Custo']
                df['OrdemFapId'] = 3
            else:
                print(f"Nenhuma coluna de índice encontrada no arquivo {nome_arquivo}.")
                continue

            colunas_final = ['Número de Ordem', 'Indice', 'OrdemFapId', 'CNAE', 'Vigencia']
            df = df[colunas_final]

            cnae = df = df['CNAE']
            nome = 'ordemFAP' + cnae
            caminho = os.path.join(destino, f'{nome}{numero}.xlsx')
            with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name=nome, index=False)
            print(f"Arquivo salvo em: {caminho}")
            os.remove(caminho_arquivo)

        except Exception as e:
            print(f"Erro ao processar o arquivo {nome_arquivo}: {e}")

