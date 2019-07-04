# -*- coding: utf-8 -*-
"""
Spyder Editor

Este é um arquivo de script temporário.
"""
"""
  Links consultados:
      * https://www.geeksforgeeks.org/reading-excel-file-using-python/
      * https://classic.scraperwiki.com/docs/python/python_excel_guide/
      * https://xlrd.readthedocs.io/en/latest/index.html
      * https://pythonhosted.org/xlrd3/cell.html
"""

import xlrd
import datetime

pasta = xlrd.open_workbook('josafá-arquivo-morto.xlsx')

for k in range(1, pasta.nsheets):
    planilha = pasta.sheet_by_index(k)    
    for i in range(planilha.nrows):
        sc = sum([len(str(planilha.cell(i,k).value)) for k in range(planilha.ncols)])
        if sc != 0:
            print(f'{i}: ', end='')
            for j in range(planilha.ncols):
                celula = planilha.cell(i,j)
                if celula.ctype == xlrd.XL_CELL_TEXT:
                    valor = celula.value
                else:
                    if celula.ctype == xlrd.XL_CELL_DATE:
                        tupla_data = xlrd.xldate_as_tuple(celula.value, pasta.datemode)
                        valor = f'{tupla_data[2]}/{tupla_data[1]}/{tupla_data[0]}'
                    else:
                        valor = celula.value
                print(valor, end='|')
            print()
    input(f'Fim da Planilha {planilha.name}')
