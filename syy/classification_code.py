import xlrd
import xlwt
from xlutils.copy import copy

old_excel = xlrd.open_workbook('增值税发票税控开票软件清单信息数据接口.xls', formatting_info=True)
code_data = old_excel.sheets()[0].col_values(15)[1:]
new_excel = copy(old_excel)
row = 1
col = 15
ws = new_excel.get_sheet(0)

style = xlwt.XFStyle()
style.num_format_str = '@'
font = xlwt.Font()
font.name = '宋体'
style.font = font

for x in code_data:
    if len(x) < 19:
        ws.write(row, col, label = x.zfill(19), style = style)
    row+=1

new_excel.save('new_code_data.xls')
