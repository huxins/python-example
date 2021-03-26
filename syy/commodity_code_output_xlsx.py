import xlsxwriter

workbook = xlsxwriter.Workbook('商品编码.xlsx')
worksheet = workbook.add_worksheet()

with open('商品编码.txt', 'rt') as f:
    data = f.read()
data = data.splitlines(True)
count = 1
row = 0

for x in data:
    if count < 3:
        count+=1
        continue
    x = x.split(',')
    column = 0
    for e in x:
        if row==0 and column==0:
            e = e[3:]
        worksheet.write(row,column,e)
        column +=1
    row += 1

workbook.close()
