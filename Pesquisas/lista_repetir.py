import pandas as pd

# df = pd.DataFrame({'points': [25, 12, 15, 14],
#                    'assists': [5, 7, 13, 12]})
df = pd.DataFrame({})

df['bounds'] = pd.Series([3, 3, 7, 4, 5, 5, 6])
df['rebounds'] = pd.Series([3, 3, 7])


df = df.fillna(0)

print(df)

writer = pd.ExcelWriter('spbrenome.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Diogo')

writer.save()
