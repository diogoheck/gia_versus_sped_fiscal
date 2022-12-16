from openpyxl import Workbook
from banco_dados import ler_planilha as ler
from banco_dados import criar_dicionario as dic
from sped_fiscal import percorrer_arquivos_sped as percorrer
from c190 import ler_c190_speds as ler_c190
import os
from time import sleep
import warnings
import sys

SPED = './SPEDs'
ARQUIVO_GERADO = 'C190.xlsx'


def remover_arquivo(arquivo):
    try:
        if os.path.exists(arquivo):
            print('\n')
            
            print(f'removendo arquivo "{arquivo}"')
            
            print('\n')
            sleep(2)
            os.remove(arquivo)
    except PermissionError:
        print('\n')
        
        print(f'o arquivo "{arquivo}" está aberto, fazer fechar o arquivo para executar o programa!')
        
        print('\n')
        os.system('pause')
        sys.exit()
        


if __name__ == '__main__':

    warnings.simplefilter(action='ignore', category=FutureWarning)
    remover_arquivo('C190.xlsx')

    try:

        wb = Workbook()
        ws = wb.active
        ws.title = 'C190'
        wb.save(ARQUIVO_GERADO)

        remover_arquivo('log.txt')
        remover_arquivo('ZILDOMAR MENEZES  CIA LTDA..xlsx')

        planilha = ler.planilha('./banco_dados/planilha_operacao_fiscal.xlsx')
        dic = dic.criar_dicionario(planilha)


        # gerar planilha com informação de todos speds
        ler_c190.ler_c190_arquivos_sped(SPED, True)
        # gravar.gravar_informacoes_excel(obj_sped)


        percorrer.percorrer_todos_arquivos(SPED, dic)

        print('\n')
        print('*' * 50)
        print('Finalizado com sucesso')
        print('*' * 50)
        print('\n')
        os.system('pause')
    except Exception as e:
        print(e)

    

