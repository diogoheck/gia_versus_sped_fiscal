# importando a biblioteca pandas
import pandas as pd

# Inicializando um dicionário com arrays
data = {
    '1102': [20, 30, 40, 50],
    '1403': [99, 98, 95, 90],
    'compras': [100, 500, 600, 700]
}

# criando um dataframe com index
df = pd.DataFrame(
    data, index=['01-01-2021', '01-01-2021', '02-01-2021', '04-01-2021'])

# mostrando o dataframe
print(df)
