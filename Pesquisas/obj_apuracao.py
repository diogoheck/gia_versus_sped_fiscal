
class TipoOperacao:
    def __init__(self, nome) -> None:
        self.nome = nome
        self.lista = []

    def add_operacao(self, nome, CFOP, data, valor):
        self.lista.append(Operacao(nome, CFOP, data, valor))


class Operacao:
    def __init__(self, nome, CFOP, data, valor) -> None:
        self.nome = nome
        self.CFOP = CFOP
        self.data = data
        self.valor = valor


if __name__ == '__main__':
    lista_operacoes = []
    compras = TipoOperacao('Compras')
    lista_operacoes.append(compras)
    compras.add_operacao('Compra', '1102', '01/01/2022', 92)
    compras.add_operacao('Compra', '1403', '01/02/2022', 20)
    compras.add_operacao('Compra', '1102', '03/02/2022', 40)
    compras.add_operacao('Compra', '1102', '03/02/2022', 40)
    vendas = TipoOperacao('Vendas')
    lista_operacoes.append(vendas)
    vendas.add_operacao('Vendas', '5102', '03/02/2022', 40)
    for op in compras.lista:
        print(f'{op.nome}, {op.CFOP}, {op.data}, {op.valor}')
    for op in vendas.lista:
        print(f'{op.nome}, {op.CFOP}, {op.data}, {op.valor}')
    for li in lista_operacoes:
        print(li.nome)
