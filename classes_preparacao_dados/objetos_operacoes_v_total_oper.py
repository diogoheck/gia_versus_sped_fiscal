import calendar
import pandas as pd


class PastaTrabalho:

    def __init__(self, df) -> None:
        self.df = df

    def create_arquivo_excel(self, nome_pasta):
        self.writer = self.pd.ExcelWriter(nome_pasta, engine='xlsxwriter')

    def salvar_excel(self):
        self.writer.save()

    def inserir_dados(self, dic_operacoes):
        self.df = pd.DataFrame()

        for k, v in sorted(dic_operacoes.items()):
            for x, y in sorted(v.items()):
                # print(x, y)
                self.df[x] = self.pd.Series(y)

        self.df.sum()
        # self.df = self.df.fillna(0)

    def criar_nova_plan(self, nome_plan):
        self.df.to_excel(self.writer, sheet_name=nome_plan)


def calcula_data(mes, ano):
    if mes == '01':
        return (calendar.monthrange(int(ano), 1)[1], '01')
    elif mes == '02':
        return (calendar.monthrange(int(ano), 2)[1], '02')
    elif mes == '03':
        return (calendar.monthrange(int(ano), 3)[1], '03')
    elif mes == '04':
        return (calendar.monthrange(int(ano), 4)[1], '04')
    elif mes == '05':
        return (calendar.monthrange(int(ano), 5)[1], '05')
    elif mes == '06':
        return (calendar.monthrange(int(ano), 6)[1], '06')
    elif mes == '07':
        return (calendar.monthrange(int(ano), 7)[1], '07')
    elif mes == '08':
        return (calendar.monthrange(int(ano), 8)[1], '08')
    elif mes == '09':
        return (calendar.monthrange(int(ano), 9)[1], '09')
    elif mes == '10':
        return (calendar.monthrange(int(ano), 10)[1], '10')
    elif mes == '11':
        return (calendar.monthrange(int(ano), 11)[1], '11')
    elif mes == '12':
        return (calendar.monthrange(int(ano), 12)[1], '12')


def geracao_datas(dicionario, data, valor):
    mes = data[2:4]
    ano = data[4:8]
    data = calcula_data(data[2:4], data[4:8])

    for i in range(1, data[0] + 1):
        if len(str(i)) == 1:
            i = '0' + str(i)
        dicionario[valor].update({str(i) + str(mes) + str(ano): 0.0})
    dicionario[valor].update({'TOTAL': 0.0})
    return dicionario


def criar_nova_operacao(dic_operacoes, operacao, CFOP, data):

    dicionario_CFOP = {}
    dicionario_total_operacao = {}
    dicionario_total_operacao[operacao.split('-')[1]] = {}
    dicionario_CFOP[CFOP] = {}
    dicionario_CFOP = geracao_datas(dicionario_CFOP, data, CFOP)
    dicionario_total_operacao = geracao_datas(
        dicionario_total_operacao, data, operacao.split('-')[1])
    dic_operacoes[int(operacao.split('-')[0])] = {}
    dic_operacoes[int(operacao.split('-')[0])].update(dicionario_CFOP)
    dic_operacoes[int(operacao.split('-')[0])
                  ].update(dicionario_total_operacao)

    # print(dic_operacoes)
    return dic_operacoes


def criar_novo_CFOP(dic_operacoes, operacao, CFOP, data):

    dicionario_CFOP = {}
    dicionario_CFOP[CFOP] = {}
    dicionario_CFOP = geracao_datas(dicionario_CFOP, data, CFOP)
    dic_operacoes[int(operacao.split('-')[0])].update(dicionario_CFOP)

    return dic_operacoes


def atualizar_dados(dic_operacoes, operacao, CFOP, data, valor):

    # CFOP

    dic_operacoes[int(operacao.split('-')[0])][CFOP][data] += valor
    dic_operacoes[int(operacao.split('-')[0])][CFOP]['TOTAL'] += valor

    # OPERAÇÃO
    dic_operacoes[int(operacao.split(
        '-')[0])][operacao.split('-')[1]][data] += valor
    dic_operacoes[int(operacao.split(
        '-')[0])][operacao.split('-')[1]]['TOTAL'] += valor
    
    # print(dic_operacoes)

    return dic_operacoes


if __name__ == '__main__':
    df = pd.DataFrame()
    dic_operacoes = {}

    dado1 = ['1-Compra', '1403', '01012022', 45.20]
    dado2 = ['1-Compra', '1102', '02012022', 105.25]

    if not dic_operacoes.get('1-Compra'):
        dic_operacoes = criar_nova_operacao(
            dic_operacoes, '1-Compra', '1403', '01012022')
    if not dic_operacoes['1-Compra'].get('1102'):
        dic_operacoes = criar_novo_CFOP(
            dic_operacoes, '1-Compra', '1102', '03012022')
    dic_operacoes = atualizar_dados(
        dic_operacoes, '1-Compra', '1403', '01012022', 45.20)
    dic_operacoes = atualizar_dados(
        dic_operacoes, '1-Compra', '1102', '02012022', 105.25)

    print(dic_operacoes)
    for k, v in sorted(dic_operacoes.items()):
        for x, y in sorted(v.items()):
            df[x] = pd.Series(y)

    writer = pd.ExcelWriter('sabado.xlsx', engine='xlsxwriter')

    df.to_excel(writer, sheet_name='Diogo')

    writer.save()
