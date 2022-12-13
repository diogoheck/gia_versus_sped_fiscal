

from planilha_excel import criar_plan
import os
from sped_fiscal import sped_fiscal as sp
import pandas as pd
from c190 import gravar_c190_speds as gravar


def eh_arquivo_txt(arquivo):
    if arquivo.split('.')[-1] == 'txt':
        return True
    return False


def criar_novo_sped(razao_social, CNPJ, IE):
    novo_sped = sp.SpedFiscal(razao_social, CNPJ, IE)
    return novo_sped


def criar_nova_planilha():
    pass


def criar_nova_pasta_trabalho(nome_pasta):
    nova_pasta = criar_plan.PastaTrabalho(nome_pasta)
    nova_pasta.criar_nova_pasta_de_trabalho()
    return nova_pasta


def ler_c190_arquivos_sped(pasta, primeira_passada):
    pasta_trabalho = criar_plan.PastaTrabalho(pd)
    for diretorio, subpastas, arquivos in os.walk(pasta):
        for arquivo in arquivos:
            sped = os.path.join(diretorio, arquivo)
            if eh_arquivo_txt(arquivo.split('.')[-1]):
                with open(sped, encoding='ansi') as arquivo_sped:
                    for linha in arquivo_sped:
                        registro = linha.strip().split('|')
                        if registro[1] == '9999':
                            break
                        if registro[1] == '0000':
                            novo_sped = criar_novo_sped(
                                registro[6], registro[7], registro[10])
                            data_sped_inicial = registro[4]
                            data_sped_final = registro[5]
                            if primeira_passada:
                                pasta_trabalho.create_arquivo_excel(
                                    registro[6] + '.xlsx')
                                primeira_passada = False
                        elif registro[1] == 'C100':
                            # data_nota = registro[11]
                            numero_nota = registro[8]
                            novo_sped.add(registro)
                        elif registro[1] == 'C190':
                            nf = novo_sped.procurar(numero_nota)
                            nf.add_C190(registro)
                    gravar.gravar_informacoes_excel(novo_sped)
    gravar.agrupar_planilha_pandas()
