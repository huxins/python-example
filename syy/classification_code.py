import xlrd
import xlwt
from xlutils.copy import copy

style = xlwt.XFStyle()
style.num_format_str = '@'
font = xlwt.Font()
font.name = '宋体'
style.font = font

old_excel = xlrd.open_workbook('test.xls', formatting_info=True)
new_excel = copy(old_excel)
col = 15
count = 0
for sheet in old_excel.sheets():
    row = 1
    code_data = sheet.col_values(15)[1:]
    ws = new_excel.get_sheet(count)
    for x in code_data:
        if type(x) == float and x == int(x):
            x=str(int(x))
        if len(x) < 19:
            ws.write(row, col, label = x.ljust(19,'0'), style = style)
        row+=1
    count+=1

new_excel.save('new_code_data.xls')
