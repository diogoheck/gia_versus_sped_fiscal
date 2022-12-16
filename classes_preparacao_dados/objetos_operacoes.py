import pandas as pd


def gravar_excel(df, dic):
    for k, v in dic.items():
        for x, y in v.items():
            # print(x, y)
            df[x] = pd.Series(y)

    writer = pd.ExcelWriter('sabado.xlsx', engine='xlsxwriter')

    df.to_excel(writer, sheet_name='Diogo')

    writer.save()


def pesquisa_operacao(dic_operacoes, descricao, CFOP, data, valor):
    if not dic_operacoes:
        dic_cfops = {}
        dic_datas = {}
        dic_datas[data] = valor
        dic_cfops[CFOP] = dic_datas
        dic_operacoes[descricao] = dic_cfops

    else:
        if dic_operacoes.get(descricao):  # já existe a operação
            if dic_operacoes.get(descricao).get(CFOP):  # já existe o CFOP
                if dic_operacoes.get(descricao).get(CFOP).get(data):  # já existe a data
                    dic_operacoes.get(descricao).get(CFOP)[data] += valor
                    dic_operacoes.get(descricao).get(CFOP)['Total'] += valor
                else:
                    if dic_operacoes.get(descricao).get(CFOP).get('Total'):
                        dic_operacoes.get(
                            descricao).get(CFOP)['Total'] += valor
                    dic_datas = {}
                    dic_total = {}
                    dic_datas[data] = valor

                    dic_operacoes.get(descricao).get(CFOP).update(dic_datas)

            else:
                dic_cfops = {}
                dic_datas = {}
                dic_total = {}
                dic_cfops[CFOP] = dic_datas
                dic_total['Total'] = valor
                dic_datas[data] = valor
                dic_cfops[CFOP].update(dic_total)
                dic_operacoes.get(descricao).update(dic_cfops)

        else:
            dic_cfops = {}
            dic_datas = {}
            dic_total = {}
            nova_operacao = {}
            dic_datas[data] = valor
            # dic_datas['Total'] = valor
            dic_total['Total'] = valor
            dic_cfops[CFOP] = dic_datas
            dic_cfops[CFOP].update(dic_total)
            nova_operacao[descricao] = dic_cfops
            dic_operacoes.update(nova_operacao)

    return dic_operacoes


if __name__ == '__main__':
    df = pd.DataFrame()
    dic_operacoes = {}

    for d in ['01-01-2022', '02-01-2022', '03-01-2022', '04-01-2022',
              '05-01-2022', '06-01-2022', '07-01-2022', '08-01-2022',
              '09-01-2022', '10-01-2022', 'Total']:
        dic_operacoes = pesquisa_operacao(
            dic_operacoes, 'Neutro', '', d, '')

    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Compras', '1403', '02-01-2022', 20)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Vendas', '5102', '03-01-2022', 1000)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Compras', '1403', '03-01-2022', 30)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Compras', '1102', '10-01-2022', 60)

    # dic_operacoes = pesquisa_operacao(
    #     dic_operacoes, 'Compras', '1102', '02-01-2022', 45)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Vendas', '5102', '03-01-2022', 1000)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Vendas', '5405', '01-01-2022', 1500)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Vendas', '5102', '10-01-2022', 630)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Dev_Vendas', '5411', '02-01-2022', 150)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Dev_Vendas', '5411', '07-01-2022', 140)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Dev_compra', '1411', '05-01-2022', 150)
    dic_operacoes = pesquisa_operacao(
        dic_operacoes, 'Compras', '1403', '07-01-2022', 250)
    gravar_excel(df, dic_operacoes)
    print(dic_operacoes)
