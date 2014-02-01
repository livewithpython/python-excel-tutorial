# Max Shron Data Strategy for BRAC
# Proof of concept: adding cells across sheets

import xlrd, xlwt
import sys

INPUT1 = 'input1.xlsx'
INPUT2 = 'input2.xlsx'
OUTPUT = 'output.xlsx'

def main(argv):
    tb = xlrd.open_workbook(INPUT1)
    preg = xlrd.open_workbook(INPUT2)
    outwb = xlwt.Workbook()

    out = {} # we store all the data here
    
    # pull out the upazilla
    out['upazilla'] = tb.sheets()[0].cell(1,2).value

    out['new_births'] = sum_across(preg,3,2)
    out['new_pregnancies'] = sum_across(preg,4,2)
    out['new_immunizations'] = sum_across(preg,5,2)
    out['tb_pos'] = sum_across(tb,3,2)
    out['smear_pos'] = sum_across(tb,4,2)
    out['smear_neg'] = sum_across(tb,5,2)

    write_tb_preg(out, outwb) # write data to output.xls
 
def pull_across(wb, cellY, cellX):
    """Get values at coordinates Y,X across all worksheets in wb
    
    A1 is 0,0
    B1 is 0,1
    C4 is 3,2
    """
    out = []
    for sheet in wb.sheets():
        out.append(sheet.cell(cellY,cellX).value)
    return out

def sum_across(wb, cellY, cellX):
    """Sum across one position in a number of worksheets in one wb"""
    return sum(pull_across(wb, cellY, cellX))

def write_tb_preg(out, outwb):
    s1 = outwb.add_sheet('Yearly sums')
    s1.write(1,1,"Upazilla")
    s1.write(1,2, out['upazilla'])

    s1.write(3,1, "TB Positive")
    s1.write(3,2, out['tb_pos'])

    s1.write(4,1, "Smear positive")
    s1.write(4,2, out['smear_pos'])

    s1.write(5,1, "Smear negative")
    s1.write(5,2, out['smear_neg'])

    s1.write(6,1, "New Births")
    s1.write(6,2, out['new_births'])

    s1.write(7,1, "New Pregnancies")
    s1.write(7,2, out['new_pregnancies'])

    s1.write(8,1, "Immunizations")
    s1.write(8,2, out['new_immunizations'])
    output.save('output.xls')
   


if __name__ == "__main__":
    main(sys.argv)
