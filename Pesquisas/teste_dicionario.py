import pandas as pd

dic_compras = {
    '1403':
    {'01-01-2022': 1099,
     '02-01-2022': 1464,
     'total': 2500,
     },
        '1102': {
        '01-01-2022': 1065,
        'total': 2700,

    },
    'total_compras': {
        '01-01-2022': 2000,
        '02-01-2022': 15000
    },


        '': {
        '': ''
    },
        '5405': {
        '01/01/2022': 1566,
        '03/01/2022': 1000,
        'total': 3566,
    },
    '5102': {
        '01-01-2022': 1000,
        '10-01-2022': 500,
        'total': 1000,

    },
    'total_vendas': {
        '01-01-2022': 3566,
        '02-01-2022': 1000,
        'Total': 4000
    }
}


dic_data2 = {
    '1403': [15, 24, 37],

    '1556': [10, 208],
    '1102': [500, 1007, 809],
}

# df = pd.concat({k: pd.Series(v)
#                 for k, v in dic_data.items()}).reset_index()

# print(df)
df = pd.DataFrame()


def add():
    dic_cfop = {}
    lista = []
    dic_data.get('01-01-2022').get('1403').append(45)
    dic_cfop['1407'] = lista
    dic_data['25-03-2022'] = dic_cfop['1407'].append(78)


if __name__ == '__main__':

    for k, v in dic_compras.items():
        df[k] = pd.Series(v)


writer = pd.ExcelWriter('alabama.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Diogo')

writer.save()
