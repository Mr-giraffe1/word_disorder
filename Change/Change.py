import xlrd
def f(s):
    data=xlrd.open_workbook('x.xlsx')
    table=data.sheets()[0]
    rows=table.nrows
    for i in range(rows):
        if table.cell_value(i,0)==s:
            print("查找结果为"+table.cell_value(i,1))

print("请输入查找值")
s=input()
f(s)
