import pandas as pd

# Carrega o DataFrame a partir do Excel
db = pd.read_excel("modelos_jfa.xlsx")

# Itera sobre as linhas do DataFrame
# for index, row in db.iterrows():
    # produto = row["Produto"]
    # if "200a" in produto.lower():
        # print(f"Produto: {row["Produto"]} - {row["Produto2"]} - {row["Qtde"]}")

soma_quantidade = db.groupby('Produto2')['Qtde'].sum()
# 
print("Soma da quantidade por tipo de fonte:")
print(soma_quantidade)

db = pd.read_excel("modelos_usina.xlsx")

# Itera sobre as linhas do DataFrame
# for index, row in db.iterrows():
    # produto = row["Produto"]
    # if "200a" in produto.lower():
        # print(f"Produto: {row["Produto"]} - {row["Produto2"]} - {row["Qtde"]}")

soma_quantidade = db.groupby('Produto2')['Qtde'].sum()
# 
print("Soma da quantidade por tipo de fonte:")
print(soma_quantidade)
