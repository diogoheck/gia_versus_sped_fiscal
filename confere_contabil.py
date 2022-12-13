from openpyxl import Workbook
from banco_dados import ler_planilha as ler
from banco_dados import criar_dicionario as dic
from sped_fiscal import percorrer_arquivos_sped as percorrer
from c190 import ler_c190_speds as ler_c190
import os

SPED = './SPEDs'
ARQUIVO_GERADO = 'C190.xlsx'

wb = Workbook()
ws = wb.active
ws.title = 'C190'
wb.save(ARQUIVO_GERADO)


planilha = ler.planilha('./banco_dados/planilha_operacao_fiscal.xlsx')
dic = dic.criar_dicionario(planilha)

# gerar planilha com informação de todos speds
ler_c190.ler_c190_arquivos_sped(SPED, True)
# gravar.gravar_informacoes_excel(obj_sped)

percorrer.percorrer_todos_arquivos(SPED, dic)

os.system('pause')
