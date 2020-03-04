import xlwings as xw

wb = xw.Book('work_project.XLS')  # connect to an existing file in the current working directory
sht = wb.sheets['Sheet1']
sht2 = wb.sheets['Sheet2']

sht2.range('A1').value = 'Foo 1'

data = sht.range('A1').value 
sht2.range('A1').value = data
data = sht.range('A2:N173').value
quantity = 0
for row in data:
    quantity += row[10]
    if quantity > 24:
        print("quantity")
        quantity = 0
        