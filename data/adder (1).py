# Max Shron Data Strategy for BRAC
# Proof of concept: adding cells across sheets

import xlrd, xlwt

INPUT1 = 'input1.xlsx'
INPUT2 = 'input2.xlsx'
OUTPUT = 'output.xlsx'

wb1 = xlrd.open_workbook(INPUT1)
output = xlwt.Workbook()

s1 = output.add_sheet('First type')
sum1, sum2, sum3 = 0, 0, 0
for sheet in wb1.sheets():
    sum1 += sheet.cell(3,2).value
    sum2 += sheet.cell(4,2).value
    sum3 += sheet.cell(5,2).value
s1.write(3, 2, sum1)
s1.write(4, 2, sum2)
s1.write(5, 2, sum3)
    
# Copy headers, etc
# 1,1
# 1,2
# 0,[3-5]

output.save('output.xls')
