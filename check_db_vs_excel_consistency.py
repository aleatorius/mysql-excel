from site import execsitecustomize
import pyodbc 
import argparse
from itertools import groupby

from pathlib import Path
import os
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl import load_workbook
from openpyxl.utils.cell import cols_from_range, coordinate_to_tuple,get_column_interval

def all_equal(iterable):
    g = groupby(iterable)
    return next(g, True) and not next(g, False)

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

def diff_write(diff,diff_file,output,entry,data,sheet,row):
    indices = [i for i, x in enumerate(diff) if x == False]
    (min_col, min_row, max_col, max_row) = data[-1]
    temp = get_column_interval(min_col, max_col)
    for index in indices:
        if not output[index]:
            print(output[index],"is empty")
        else:
            pass
        diff_file.write('Sheet: "'+sheet.title+ '" Cell:"'+ str(temp[index])+str(min_row+row)+' ' + str(data[1][index]) +'"\n Excel: "'+ str(entry[index]) +'" db: "'+ str(output[index])+'"\n')

def noentry_write(noentry_file,entry,data,sheet,row):
    (min_col, min_row, max_col, max_row) = data[-1]
    temp = get_column_interval(min_col, max_col)
    noentry_file.write('Sheet: "'+sheet.title+ '" Cell:"'+ str(temp)+str(min_row+row)+' ' + str(data[1]) +'"\n Excel: "'+ str(entry) +'\n')

def get_columns_to_compare(sheet, id_set_to_compare):
    db_row = 1
    db_names,db_columns,db_ranges = get_sheet_structure(sheet = sheet)
    output = []
    for data in zip(db_names,db_columns,db_ranges):
        (col_low, row_low, col_high, row_high) = data[-1]
        for cols in sheet.iter_cols(min_col=col_low,min_row=row_low+db_row, max_col=col_high, max_row= row_low+db_row):
            for cell in cols:
                for col in id_set_to_compare:
                    if cell.value == col[0] and data[0] == col[1]:
                        output.append(cell.column)
    return output


def compare_between_values_in_columns(sheet,input_columns,warnings_file):
    db_cols_row = 2

    for row in range(db_cols_row+1,sheet.max_row+1):
        columns_to_compare = []
        compared_cells_cols = []
        for col in input_columns:
            cell = sheet.cell(row=row,column=col)
            columns_to_compare.append(cell.value)
            compared_cells_cols.append(cell.column_letter)
        if all_equal(columns_to_compare) == True:
            return columns_to_compare[0],True
        else:
            warnings_file.write("ERROR, exercise ids do not match! Check these cells and correct them:\n")
            string = ''
            for cell_col in compared_cells_cols:
                string = string + cell_col + str(row)+' '
            warnings_file.write(string + '\n')
            warnings_file.write('Values: '+str(columns_to_compare)+'\n\n')
            return columns_to_compare,False


def check_one_value_columns_vs_values(sheet,input_columns,values, warnings_file):
    db_cols_row = 2
    if len(input_columns) != len(values):
        print("ERROR in check_one_value_column")
        exit()
    for colval in zip(input_columns,values):  
        columns_to_compare = [colval[-1]]            
        for row in range(db_cols_row+1,sheet.max_row+1):
            cell = sheet.cell(row=row,column=colval[0])
            columns_to_compare.append(cell.value)
        print(columns_to_compare)
        if all_equal(columns_to_compare) == True:
            return True
        else:
            warnings_file.write("ERROR, error in the column! Check this column and correct it:\n")
            warnings_file.write('Row: '+ cell.column_letter+'\n\n')
            return False


def check_one_value_columns(sheet,input_columns,warnings_file):
    db_cols_row = 2
    columns_to_compare = [] 
    columns_letter =  []
    for colval in input_columns:              
        for row in range(db_cols_row+1,sheet.max_row+1):
            cell = sheet.cell(row=row,column=colval)
            columns_to_compare.append(cell.value)
        columns_letter.append(cell.column_letter)
    print(columns_to_compare)
    if all_equal(columns_to_compare) == True:
        return True
    else:
        warnings_file.write("ERROR, error in the columns! Check these columns and correct them:\n")
        warnings_file.write('Rows: '+ str(columns_letter)+'\n\n')
        return False




def main(folder):
    
    warnings_file = open('warnings.txt','w')
    noentry_file = open('noentry.txt', 'w')
    ser_file = open('server_private.md','r')
    info = []
    for i in ser_file:
        info.append(i.split()[-1].replace("'",''))
    [server,database,username,password] = info
    
    #connect to the calst database
    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+'; UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()
    
    folder_path = Path(folder)
    if folder_path.exists():
        pass
    else:
        print(' There is no such folder as "'+folder+'"','\n Please enter the correct folder name')
        exit()
    
    exercise = folder_path/'exercise.xlsx'
   
    if exercise.exists():
        print("ok, exercise exists in this folder")
        wb = load_workbook(filename = exercise)
        
        sheet = wb['Exercise']
        if sheet.max_row != 3:
            warnings_file.write('For the sheet: '+sheet.title+' there are more rows than 3: '+ str(sheet.max_row)+'\n')
        if sheet.max_column < 14:
            warnings_file.write('For the sheet: '+sheet.title+' there are less than 14 columns: '+ str(sheet.max_column)+'\n')
            
       
        
        
        #check wrapper for this exercise
        
        #for row in range(1,sheet.max_row):
        id_set_to_compare = [('Exercise_Id','WrapperExercises'),('Id','Exercises'),('ExerciseId','Properties')]
        
        exerciseid = get_columns_to_compare(sheet=sheet, id_set_to_compare=id_set_to_compare)
        compared = compare_between_values_in_columns(sheet=sheet,input_columns=exerciseid,warnings_file=warnings_file)
        if compared[-1] == True:
            print(compared[0])
            ExerciseId = compared[0]
        else:
            print(compared)
            exit()
        



        
        #check exercises in the different sheet
        
        confusionbox_sheet =  []
        matches = ['ConfusionBox']
        for name in wb.sheetnames:
            if any(x in name for x in matches):
                confusionbox_sheet.append(name)
        print(confusionbox_sheet)
        for cb in confusionbox_sheet:
            sheet_cb = wb[cb]
            print(sheet_cb.max_row)
            id_set_to_compare = [('ExerciseId','ConfusionBoxes')]
        
            exerciseid = get_columns_to_compare(sheet=sheet_cb, id_set_to_compare=id_set_to_compare)
            check_one_value_columns_vs_values(sheet=sheet_cb,input_columns=exerciseid,values=[ExerciseId], warnings_file=warnings_file)
            print(exerciseid,"cb")

            id_set_to_compare = [('Id','ConfusionBoxes'),('ConfusionBox_Id','TranscriptionConfusionBoxes')]
        
            confusionboxid = get_columns_to_compare(sheet=sheet_cb, id_set_to_compare=id_set_to_compare)
            check_one_value_columns(sheet=sheet_cb,input_columns=confusionboxid, warnings_file=warnings_file)
            print(confusionboxid,"cb")
            
        #compared = compare_values_in_columns(sheet=sheet,input_column=exerciseid,warnings_file=warnings_file)
            

            
        exit()
        sheet = wb['Exercise']
        if sheet.max_row != 3:
            warnings_file.write('For the sheet: '+sheet.title+' there are more rows than 3: '+ str(sheet.max_row)+'\n')
        if sheet.max_column < 14:
            warnings_file.write('For the sheet: '+sheet.title+' there are less than 14 columns: '+ str(sheet.max_column)+'\n')
            
        names,columns,ranges = get_sheet_structure(sheet = sheet)
        db_row = 1
        db_cols_row = 2 

        exit()
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
                    if len(output) == 0:
                        noentry_write(noentry_file=noentry_file,entry=entry,data=data,sheet=sheet, row=row)
                    elif len(output)>1:
                        print("ERROR, not unique entry, exiting ", data[0] )  
                        exit()
                    else:
                        pass
                elif data[0] == 'TranscriptionConfusionBoxes':
                    output = get_full_request(database=database,cursor=cursor,data=data, entry=entry)
                    if len(output) ==0:
                        noentry_write(noentry_file=noentry_file,entry=entry,data=data,sheet=sheet, row=row)

                    elif len(output)>1:
                        print("ERROR, not unique entry, exiting ", data[0] )  
                        exit()
                    else:
                        pass
                else:
                    sqlcom = 'SELECT * FROM ['+database+'].[dbo].[' + data[0]+'] WHERE '
                    sqlcom = sqlcom + str(data[1][0]) + ' = ' + str(entry[0])
                    cursor.execute(sqlcom)
                    output = [list(i) for i in cursor.fetchall()] 
                    if len(output) == 1:
                        diff = compare(entry=entry,data=output[0] )
                        if False in diff:
                            diff_write(diff=diff,diff_file=diff_file,output=output[0],entry=entry,data=data,sheet=sheet, row=row)
                    elif len(output) == 0:
                        noentry_write(noentry_file=noentry_file,entry=entry,data=data,sheet=sheet, row=row)
                    else:
                        print('wtf', entry, output)
                        exit() 
    else:
        print('No excel file in this folder')
    warnings_file.close
    ser_file.close
    noentry_file.close      

if __name__ == "__main__":
    parser = argparse.ArgumentParser(prog='python check_db_vs_excel_consistency.py -f foldername')
    parser.add_argument('-f',dest='folder')
    args = parser.parse_args()
    if args.folder:
        main(folder = args.folder)
    else:
        #folder='C:\Source\Repos\python_tools\Spanish_course_styled\Beginner\Lesson 1\The alphabet'
        folder = 'C:\Source\Repos\python_tools\Spanish_course_styled\Beginner\Lesson 1\\Numbers 1'
        main(folder=folder)