import gkeepapi
import xlwt
import xlsxwriter
from xlutils.copy import copy
from xlrd import open_workbook
from datetime import datetime
from datetime import date

pathToFile = r'C:\Users\Deivid\Desktop\test.xls'

keep = gkeepapi.Keep()
success = keep.login('floodstemp@gmail.com','Floods192!')


#note = keep.createList('ToDo',[('Item 1', False), ('Item 2', False), ('Item 3', True)])
note = keep.get('10QdMvjJ2OD4Ddy_wKUeh5qA6KYzJhyUTUPirzfESa-I182kczz8rF4OG8B06GyagrcLs')

keep.sync()

print (note.title)
print (note.text)


style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')


rb = open_workbook(pathToFile)
r_sheet = rb.sheet_by_index(0)

wb = copy(rb)
ws = wb.get_sheet(0)


#Only works with XLSWriter, need to find a way to set up automatically
workbook = xlsxwriter.Workbook('chart_line.xlsx')
worksheet = workbook.add_worksheet()
chart = workbook.add_chart({'type': 'line'})

#cols, rows, value
ws.write(0, 0, note.title)
ws.write(1, 0, datetime.now(), style1)

i = r_sheet.nrows
j = 0
total = 0
completed = 0

today = date.today()
ws.write(i+1,0,today.strftime("%B %d, %Y"))

for x in note.items:
    if j < len(note.items):
        ws.write(i+2, 0,  note.items[j].text)
        ws.write(i+2, 1, note.items[j].checked)

        if x.checked:
            completed+=1
    j+=1
    i+=1


total =  len(note.items)

ws.write(i+3, 0, 'Total Completed:')
ws.write(i+3, 1,  str(completed) + '/' + str(total))

wb.save('test.xls')