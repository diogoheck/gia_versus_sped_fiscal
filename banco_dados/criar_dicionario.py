from openpyxl import load_workbook


def planilha(nome):

    wb = load_workbook(nome)

    ws = wb.active

    return ws


def criar_dicionario(plan):
    dic = {}

    for linha in plan:
        lista_cfops = linha[0].value
        tipo_operacao = linha[1].value
        lista_cfops = linha[0].value.split('|')

        for cfop in lista_cfops:
            if cfop:
                dic[cfop] = tipo_operacao

    return dic


if __name__ == '__main__':
    planilha = planilha('./banco_dados/planilha_operacao_fiscal.xlsx')
    dic = criar_dicionario(planilha)
    print(dic)
