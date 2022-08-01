import pyodbc 
from pathlib import Path
import os
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl import load_workbook

def compare(entry, data):
    if len(entry) == len(data):
        diff = []
        for ind,ex in enumerate(entry):
            if ex == data[ind]:
                diff.append(True)
            else:
                diff.append(False)
        return diff
    else:
        print("ERROR in compare entries", entry, data)
            

def get_full_request(cursor, database, data, entry):
    sqlcom = 'SELECT * FROM ['+database+'].[dbo].[' + data[0]+'] WHERE '
    count = 0
    for ind, ex in enumerate(data[1]):
        if count < len(entry)-1:
            sqlcom = sqlcom + str(ex) + ' = ' + str(entry[ind]) + ' AND '
        else:
            sqlcom = sqlcom + str(ex) + ' = ' + str(entry[ind]) 
        count = count + 1 
    cursor.execute(sqlcom)
    output = [list(i) for i in cursor.fetchall()] 
    return output    

def get_sheet_structure(sheet):
    #a database row specified as merged cells
    merged_ranges = sheet.merged_cells.ranges
    database = []
    ranges = []
    for range in merged_ranges:
        (col_low, row_low, col_high, row_high) = range.bounds
        ranges.append(range.bounds)
        database.append(sheet.cell(row_low,col_low).value)
    database_and_ranges = zip(database, ranges)
    database_cols = []
    for data in database_and_ranges:
        (col_low, row_low, col_high, row_high) = data[-1]
        if row_high-row_low != 0:
            print("Error,too many rows for info caption")
            exit()
        else:
            pass
        
        col_id = []
        for row in sheet.iter_rows(min_col=col_low,max_col=col_high,min_row=row_low+1,max_row=row_high+1):
            for cell in row:
                col_id.append(cell.value)
        database_cols.append(col_id)

    database_ranges_columns = zip(database,database_cols, ranges)
    return database, database_cols,ranges
        
def main():
    diff_file = open('diff.txt','w')
    ser_file = open('server_private.md','r')
    info = []
    for i in ser_file:
        info.append(i.split()[-1].replace("'",''))
    [server,database,username,password] = info
    

    #connect to the calst database

    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+'; UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()

    folder = 'C:\Source\Repos\python_tools\Spanish_course_styled\Beginner\Lesson 1\The alphabet'
    folder_path = Path(folder)
    exercise = folder_path/'exercise.xlsx'
   
    if exercise.exists():
        print("ok, exercise exists in this folder")
        wb = load_workbook(filename = exercise)
        for i in wb.sheetnames:
            print(i)
            sheet = wb[i]
            names,columns,ranges = get_sheet_structure(sheet = sheet)
           
            for row in range(2,sheet.max_row):
                for data in zip(names,columns,ranges):
                    (col_low, row_low, col_high, row_high) = data[-1]
                    entry = []
                    for cells in sheet.iter_cols(min_col=col_low,min_row=row_low+row, max_col=col_high, max_row= row_low+row):
                        for cell in cells:
                            entry.append(cell.value)
                    
                    if all(x is None for x in entry):
                        pass
                    else:
                        if data[0] == 'WrapperExercises':
                            #it should be defined by both columns
                            output = get_full_request(database=database,cursor=cursor,data=data, entry=entry)
                            if len(output) == 1:
                                diff = compare(entry=entry,data=output[0] )
                                if False in diff:
                                    print(entry,output[0],"diff!!", data[1][diff.index(False)] )
                                    diff_file.write(sheet.title+ ' '+ entry[diff.index(False)] +' '+ output[0][diff.index(False)]+' ' + data[1][diff.index(False)]+'\n')
                            else:
                                print("ERROR, not unique entry, exiting ", data[0] )  
                                exit()
                        elif data[0] == 'TranscriptionConfusionBoxes':
                            output = get_full_request(database=database,cursor=cursor,data=data, entry=entry)
                            if len(output) == 1:
                                diff = compare(entry=entry,data=output[0] )
                                if False in diff:
                                    print(entry,output[0],"diff!!", data[1][diff.index(False)] )
                                    diff_file.write(sheet.title+ ' '+ entry[diff.index(False)] +' '+ output[0][diff.index(False)]+' ' + data[1][diff.index(False)]+'\n')
                            else:
                                print("ERROR, not unique entry, exiting ", data[0] )  
                                exit()
                        else:
                            sqlcom = 'SELECT * FROM ['+database+'].[dbo].[' + data[0]+'] WHERE '
                            sqlcom = sqlcom + str(data[1][0]) + ' = ' + str(entry[0])
                            cursor.execute(sqlcom)
                            output = [list(i) for i in cursor.fetchall()] 
                            if len(output) == 1:
                                diff = compare(entry=entry,data=output[0] )
                                if False in diff:
                                    print(entry,output[0],"diff!!", data[1][diff.index(False)] )
                                    print(sheet.title)
                                    if not output[0][diff.index(False)]:
                                        print(output[0][diff.index(False)],"is empty")
                                    else:
                                        pass
                                    diff_file.write(sheet.title+ ' '+ str(entry[diff.index(False)]) +' '+ str(output[0][diff.index(False)])+' ' + str(data[1][diff.index(False)])+'\n')

                            else:
                                print('wtf', entry, output)
                                exit()



                        

            
                
                
    else:
        print('No excel file in this folder')
            
                
            


           
if __name__ == "__main__":
    main()