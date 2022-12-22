# -*- coding: cp1252 -*-
# Se n�o rodar abra o cmd e escreva "pip -r requirements.txt"
# Jinja2
# openpyxl

from datetime import datetime
import pandas as pd

try:
    balanco_antes = pd.read_csv("./antes/saldos_estoque.csv", sep=";")
    balanco_depois = pd.read_csv("./depois/saldos_estoque.csv", sep=";")
except FileNotFoundError:
    print("Verifique se os arquivos com saldos de estoque est�o nos diret�rios corretos.")
    exit()

balanco_antes.head()

# Criando DataFrame que vai armazenar informa��es sobre os produtos

nome_colunas = ["C�digo","Nome Produto","Qnt. Antes","Qnt. Depois","Diferen�a"]
df_resultado = pd.DataFrame(columns= nome_colunas)

# Identificando se tem algum arquivo com mais pe�as que o outro

print(f"Quantidade de modelos antes do balan�o = {balanco_antes.shape[0]}\nQuantidade de modelos depois do balan�o = {balanco_depois.shape[0]}\n")

# Mostrando quantidad de produtos ativos

quantidades_antes = balanco_antes["Balan�o"].value_counts().to_dict()
total_itens_lidos_antes = 0
for key in quantidades_antes.keys():
  key_number = int(float(key.replace(",",".")))
  if key_number != 0:
    total_itens_lidos_antes += quantidades_antes[key] * key_number

quantidades_depois = balanco_depois["Balan�o"].value_counts().to_dict()
total_itens_lidos_depois = 0
for key in quantidades_depois.keys():
  key_number = int(float(key.replace(",",".")))
  if key_number != 0:
    total_itens_lidos_depois += quantidades_depois[key] * key_number

print(f"Quantidade de produtos em estoque antes do balan�o = {total_itens_lidos_antes}\nQuantidade de produtos em estoque depois do balan�o = {total_itens_lidos_depois}\n")

# Escrendo infroma��es em arquivo de texto

with open('informacoes.txt','w') as f:
    f.write(f"Quantidade de produtos antes do balan�o = {balanco_antes.shape[0]}\n")
    f.write(balanco_antes["Balan�o"].value_counts().to_string())
    f.write(f"\nQuantidade de produtos depois do balan�o = {balanco_depois.shape[0]}\n")
    f.write(balanco_depois["Balan�o"].value_counts().to_string())
    f.write(f"\n\nQuantidade de produtos em estoque antes do balan�o = {total_itens_lidos_antes}\nQuantidade de produtos em estoque depois do balan�o = {total_itens_lidos_depois}\n")
f.close()

# Vai pegar os produtos que existe em um dataframe e n�o no outro

def pega_produtos_amais(df, id_produto, tempo):
  global df_resultado

  produto = df.loc[df["ID Produto"] == id_produto]

  n_index = df.loc[df["ID Produto"] == id_produto].index.values[0]
  df.drop(index= n_index, inplace= True)

  sinal = "-" if tempo == "antes" else ""
  codigo = produto["Codigo produto"].values[0][:-2] if "\t" not in produto["Codigo produto"].values[0] else produto["Codigo produto"].values[0]
  df_produto = pd.DataFrame({
      "C�digo": codigo,
      "Nome Produto": tempo.upper() + " " + produto["Descri��o Produto"].values[0],  
      "Qnt. Antes": float(produto["Balan�o"].values[0].replace(',', '.')) if tempo == "antes" else 0.0, 
      "Qnt. Depois": 0.0 if tempo == "antes" else float(produto["Balan�o"].values[0].replace(',', '.')),
      "Diferen�a": float(sinal + produto["Balan�o"].values[0].replace(',', '.'))
  }, index=[0])
  
  df_resultado = pd.concat([df_resultado, df_produto], ignore_index=True)

# Verifica se existe produtos em uma dataframe que n�o tem na outra.
# Se existir j� coloca eles nos produtos irregulares e exclui do dataframe original

for id_produto in balanco_antes["ID Produto"]:
  if (id_produto not in balanco_depois["ID Produto"].values):
    pega_produtos_amais(balanco_antes, id_produto, "antes")

for id_produto in balanco_depois["ID Produto"]:
  if (id_produto not in balanco_antes["ID Produto"].values):
    pega_produtos_amais(balanco_depois, id_produto, "depois")

# Pega a quantidade de antes e depois e coompara
# Se os valores forem diferentes � adicionado na tabela de erros

for id_produto in balanco_antes["ID Produto"]:
  qnt_antes = float(balanco_antes.loc[balanco_antes["ID Produto"] == id_produto]["Balan�o"].values[0].replace(',', '.'))
  qnt_depois = float(balanco_depois.loc[balanco_depois["ID Produto"] == id_produto]["Balan�o"].values[0].replace(',', '.'))
  if (qnt_antes != qnt_depois):
    qnt_diferenca = qnt_depois - qnt_antes

    codigo = balanco_antes.loc[balanco_antes["ID Produto"] == id_produto]["Codigo produto"].values[0][:-2] if "\t" not in balanco_antes.loc[balanco_antes["ID Produto"] == id_produto]["Codigo produto"].values[0] else balanco_antes.loc[balanco_antes["ID Produto"] == id_produto]["Codigo produto"].values[0]
    df_produto = pd.DataFrame({
      "C�digo": codigo,
      "Nome Produto": balanco_antes.loc[balanco_antes["ID Produto"] == id_produto]["Descri��o Produto"].values[0], 
      "Qnt. Antes": qnt_antes, 
      "Qnt. Depois": qnt_depois, 
      "Diferen�a": qnt_diferenca
    }, index=[0])
    
    df_resultado = pd.concat([df_resultado, df_produto], ignore_index=True)

# Colocar em ordem crescente
df_resultado = df_resultado.sort_values(by="Diferen�a")

array_total = [
    len(df_resultado),
    "-",
    df_resultado["Qnt. Antes"].sum(),
    df_resultado["Qnt. Depois"].sum(),
    df_resultado["Diferen�a"].sum()
]

df_resultado.loc["Total"] = array_total

def color_number(v, color):
    return f"color: {color};" if v < 0 else "color: green;"

nome_arq = "balanco " + str(datetime.now())[:-7].replace(":", "") + ".xlsx"

df_resultado.style.\
              applymap(color_number, color='red', subset="Diferen�a").\
              format(precision=1).\
              to_excel(nome_arq,  engine='openpyxl', index= False)
