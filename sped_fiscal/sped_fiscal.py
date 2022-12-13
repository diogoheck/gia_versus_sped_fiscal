class SpedFiscal:
    def __init__(self, razao_social, CNPJ, IE):
        self.razao_social = razao_social
        self.CNPJ = CNPJ
        self.IE = IE
        self.notas = []

    def add(self, campos_c100):
        self.notas.append(C100(campos_c100))

    def procurar(self, numero):
        for nota in self.notas:
            if nota.campos_c100[1] == 'C100':
                if nota.campos_c100[8] == numero:
                    return nota


class C100:
    def __init__(self, campos_c100):
        self.campos_c100 = campos_c100
        self.itens_c190 = []

    def add_C190(self, campos_c190):
        self.itens_c190.append(C190(campos_c190))


class C190:
    def __init__(self, campos_c190):
        self.campos_c190 = campos_c190
