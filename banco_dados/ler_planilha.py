from openpyxl import load_workbook


def planilha(nome):

    wb = load_workbook(nome)

    ws = wb.active

    return ws
