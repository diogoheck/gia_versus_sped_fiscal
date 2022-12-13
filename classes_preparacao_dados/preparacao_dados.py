

def coletar_dados(dicionario_cfops, cfop, valor, data):
    if dicionario_cfops.get(cfop):
        dicionario_cfops.get(cfop).append(valor)
    else:
        lista = []
        lista.append(valor)
        dicionario_cfops.update({cfop: lista})

    return dicionario_cfops


def coletar_dados_v2(dic_global, cfop, valor, data):

    if dic_global.get(data):
        if dic_global.get(data).get(cfop):
            dic_global.get(data).get(cfop).append(valor)
        else:
            lista = []
            dic_cfops = {}
            dic_cfops[cfop] = lista.append(valor)
            dic_global.get(data).update(dic_cfops[cfop])

    else:
        lista = []
        dic_cfops[cfop] = lista.append(valor)
        dic_global[data] = dic_cfops[cfop]

    return dic_global


if __name__ == '__main__':
    dicionario = {}
    dicionario.update({'1102': [15.50]})
    dicionario.update({'1403': [15.50]})
    if dicionario.get('1556'):
        dicionario.get('1156').append(48)
    else:
        lista = []
        lista.append(95)
        dicionario.update({'1556': lista})
    print(dicionario)
