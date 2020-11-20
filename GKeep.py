#TO Do
#Make sure rows shift down with every update

#All imports
import gkeepapi
import xlwt
import xlsxwriter
from xlutils.copy import copy
from xlrd import open_workbook
from datetime import datetime
from datetime import date

pathToFile = r'D:\GoogleKeepPython\Python-Google-Keep-API-\test.xls'

keep = gkeepapi.Keep()

#!Important...make sure to delete login credentials....
success = keep.login('<Email>','<Password>!')

#Set variable note equal to note ID

gnotes = keep.all()


#---Setting up the Excel Sheet------
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

totalRows = r_sheet.nrows #All existing rows in sheet
#----End Excel Sheet setup--------


#Init varialbes
columns = 0
total = 0 #Total num of tasks to be done
completed = 0 #Actual num of completed tasks today
maxCheckLists = 0

#Get the max number of items in any note
for y in gnotes:
    if maxCheckLists < len(y.items):
        maxCheckLists = len(y.items)

#Shift all existing columns down based on the maximum items
for x in range(totalRows):
    worksheet.set_row(x, 100) 
    
#Loops going through all notes and all boxes within the gKeep notes
for y in gnotes:

    totalRows = 0
    #write to excel(cols, rows, value, what style)
    ws.write(0, columns, y.title)
    ws.write(1, columns, datetime.now(), style1)
    #print note values
    print (y.title)
    print (y.text)

    for x in y.items:
            ws.write(totalRows+2, columns,  x.text)
            ws.write(totalRows+2, columns+1, x.checked)  

            if x.checked:
                completed+=1
            
            x.checked = False  
            totalRows+=1 

    total =  len(y.items)

    #Write values to the worksheet
    ws.write(totalRows+3, columns, 'Total Completed:')
    ws.write(totalRows+3, columns+1,  str(completed) + '/' + str(total))

    #2 values for every note + leave a space between each
    columns += 3

    #end of loop

wb.save('test.xls')

#Save Changes
keep.sync()