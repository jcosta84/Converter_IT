import pandas as pd
import os
from glob import glob

# caminho do ficheiro (atenção às barras invertidas)
#caminho = r"F:\Nova pasta (2)\Itinerarios\IT a ser removido\Rt14_9_it1_20250905_lt13.csv"

# importar o CSV (ajuste sep e encoding se necessário)
#df = pd.read_csv(caminho, sep=";", encoding="latin1")

# mostrar as primeiras linhas
#print(df.head())

# diretório onde estão os CSVs
pasta = r"F:\Nova pasta (2)\Itinerarios\IT a ser removido"

# encontrar todos os arquivos .csv dentro da pasta
arquivos = glob(os.path.join(pasta, "*.csv"))

print("Arquivos encontrados:", arquivos)

# importar e juntar todos os CSVs
lista_dfs = [pd.read_csv(arq, sep=";", encoding="latin1") for arq in arquivos]

# concatenar todos num único DataFrame
df_final = pd.concat(lista_dfs, ignore_index=True)

# mostrar primeiras linhas
print(df_final)
print("\nTotal de linhas:", len(df_final))