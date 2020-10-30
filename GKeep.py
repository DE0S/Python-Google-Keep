#All imports
import gkeepapi
import xlwt
import xlsxwriter
from xlutils.copy import copy
from xlrd import open_workbook
from datetime import datetime
from datetime import date

pathToFile = r''

keep = gkeepapi.Keep()
#!Important...make sure to delete login credentials....
success = keep.login('','')

#Set variable note equal to note ID

gnotes = keep.all()

style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

rb = open_workbook(pathToFile)
r_sheet = rb.sheet_by_index(0)

#Copy current sheet values to variable
wb = copy(rb)
ws = wb.get_sheet(0)


#Only works with XLSWriter, need to find a way to set up automatically
workbook = xlsxwriter.Workbook('chart_line.xlsx')
worksheet = workbook.add_worksheet()
chart = workbook.add_chart({'type': 'line'})

totalRows = r_sheet.nrows
columns = 0
total = 0
completed = 0

today = date.today()
ws.write(totalRows+1,0,today.strftime("%B %d, %Y"))

#loop sorting list based on when notes where added


#Loops going through all notes and all boxes within the gKeep notes
for y in gnotes:
    #write to excel(cols, rows, value, what style)
    ws.write(0, 0, y.title)
    ws.write(1, 0, datetime.now(), style1)
    #print note values
    print (y.title)
    print (y.text)

    for x in y.items:
            ws.write(totalRows+2, 0,  x.text)
            ws.write(totalRows+2, 1, x.checked)  

            if x.checked:
                completed+=1
            
            x.checked = False  
            totalRows+=1 

    total =  len(y.items)

    #Write values to the worksheet
    ws.write(totalRows+3, columns, 'Total Completed:')
    ws.write(totalRows+3, columns+1,  str(completed) + '/' + str(total))
    columns += 2

    #end of loop

wb.save('test.xls')

#Save Changes
keep.sync()