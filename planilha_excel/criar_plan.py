
import pandas as pd


def color_negative_red(val):
    """
    Takes a scalar and returns a string with
    the css property `'color: red'` for negative
    strings, black otherwise.
    """
    color = 'red' if val > 0 else 'black'
    return 'color: %s' % color


class PastaTrabalho:

    def __init__(self, pd) -> None:
        self.pd = pd

    def ler_excel(self, nome_arquivo):
        self.df = self.pd.read_excel(nome_arquivo)

    def create_arquivo_excel(self, nome_pasta):
        self.writer = self.pd.ExcelWriter(nome_pasta, engine='xlsxwriter')

    def salvar_excel(self):
        self.writer.save()

    def inserir_dados(self, dic_operacoes, tipos_inex):
        self.df = self.pd.DataFrame()
        # print(dic_operacoes)
        # print(dic_operacoes.items())
        for k, v in sorted(dic_operacoes.items()):
            for x, y in sorted(v.items()):
                # if k not in tipos_inex:
                self.df[x] = self.pd.Series(y)

        self.df.sum()
        # self.df = self.df.fillna(0)

    def criar_nova_plan(self, nome_plan, linha):
        self.df.to_excel(self.writer, sheet_name=nome_plan, startrow=linha)
        workbook = self.writer.book
        worksheet = self.writer.sheets[nome_plan]
        number_format = workbook.add_format({'num_format': '#,##0.00'})
        negrito = workbook.add_format({'bold': True})
        color = workbook.add_format({'bg_color':  'green'})
        percent_format = workbook.add_format({'num_format': '0%'})
        worksheet.set_column('B:Z', 25, number_format)
        self.df.style.applymap(color_negative_red)
        # self.df.style.apply(highlight_max, subset=['F'])
        # worksheet.set_column('B:B', 20, percent_format)
        # worksheet.set_column('F1:F30', 25, color)
        # worksheet.set_row(4,  1, negrito)

        # worksheet.set_row('30:30', 25, negrito)
        # self.writer.save()


def teste(pasta_trabalho, dados):
    pasta_trabalho.criar_nova_plan('plan1', dados)


if __name__ == '__main__':
    dados = {
        'Data': [10, 20, 30, 20, 15, 30, 45],
        'CFOP': [1, 2,  3,  4,  5,  6,  7],
        'valor': [1, 2,  3,  4,  5,  6,  7],
    }
    dados2 = {
        'ID': [10, 20, 30, 20, 15, 30, 45],
        'R$': [5, 6,  3,  4,  5,  6,  7],
    }

    pasta_trabalho = PastaTrabalho(pd)
    pasta_trabalho.create_arquivo_excel('classes_separadas.xlsx')
    # dados = preencher_com_zeros(dados)
    teste(pasta_trabalho, dados)
    # pasta_trabalho.salvar_excel()
    # pasta_trabalho.criar_nova_plan('plan2', dados2)
    pasta_trabalho.salvar_excel()
