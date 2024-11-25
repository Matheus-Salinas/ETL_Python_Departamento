import os
import pandas as pd

# Caminho da pasta com os arquivos Excel
pasta = "data/raw"
arquivos_excel = [arquivo for arquivo in os.listdir(pasta) if arquivo.endswith(('.xls', '.xlsx'))]

print(f"Arquivos encontrados: {arquivos_excel}")

dados_consolidados = []

for arquivo in arquivos_excel:
    caminho_arquivo = os.path.join(pasta, arquivo)
    print(f"Processando arquivo: {arquivo}")
    try:
        # Lê diretamente a única aba do arquivo Excel
        df = pd.read_excel(caminho_arquivo)
        print(f"Colunas encontradas no arquivo {arquivo}: {df.columns.tolist()}")
        if 'Departamento' in df.columns:
            dados_selecionados = df[['Departamento']]
            dados_consolidados.append(dados_selecionados)
        else:
            print(f"A coluna 'Departamento' não foi encontrada no arquivo {arquivo}")
    except Exception as e:
        print(f"Erro ao processar o arquivo {arquivo}: {e}")

# Verifica o tamanho dos dados consolidados antes de continuar
print(f"Total de registros antes de remover duplicatas: {sum(len(df) for df in dados_consolidados)}")

dados_consolidados = pd.concat(dados_consolidados, ignore_index=True)
dados_unicos = dados_consolidados.drop_duplicates()

print(f"Total de registros únicos após remoção de duplicatas: {len(dados_unicos)}")

dados_unicos.to_excel("resultado_departamento_unico.xlsx", index=False)
print("Planilha consolidada criada: resultado_departamento_unico.xlsx")
