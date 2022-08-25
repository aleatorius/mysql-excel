from typing import ValuesView
from natsort import natsorted
from itertools import groupby
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_interval
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
import diff_folder_and_mysql
import course_structure_db_excel
from scan_for_missing_sound import missing_in_excel
from diff_folder_and_mysql import get_sheet_structure 
import pyodbc 



def get_column(sheet, row, name):
    col = []
    for cells in sheet.iter_cols(min_col=1,min_row=row, max_col=sheet.max_column, max_row=row):
        for cell in cells:
            if cell.value == name:
                col.append(cell.column)
    if len(col) == 1:
        return col[0]
    else:
        print('Warning! Several cols with name '+name+ ' exiting')
        exit()

def check_exerciseid_in_structure(wb_s,cursor,row):
    #it checks wrapper and wrappereercise for id and wrapper_id and for exercise_id
    sheet = wb_s['Lessons']
    names,columns,ranges = get_sheet_structure(sheet = sheet)
    
    
    min_col, min_row,max_col,max_row=ranges[names.index('WrapperExercises')]
    data = []
    coord = []
    ids = columns[names.index('WrapperExercises')]
    for cells in sheet.iter_cols(min_col=min_col,min_row=row, max_col=max_col, max_row=row):
        for cell in cells:
            data.append(cell.value)
            coord.append((cell.row,cell.column))
    print("excel data: ",data, ids, coord)
    
    #compare wrapper id of an exercise with id of its parent folder from Wrappers columns    
    min_col, min_row,max_col,max_row=ranges[names.index('Wrappers')]
    for cells in sheet.iter_cols(min_col=min_col,min_row=2, max_col=max_col, max_row=2):
        for cell in cells:
            if cell.value == 'Id':
                Wrapper_Id_column = cell.column
    #search for parent wrapper id scrolling upward
    finished = False
    wrapper_id_row = row
    while not finished:
        wrapper_id_row = wrapper_id_row - 1
        if sheet.cell(row=wrapper_id_row,column=Wrapper_Id_column).value:
            finished = True
            Wrapper_Id = sheet.cell(row=wrapper_id_row,column=Wrapper_Id_column).value
    #check where whrapper_id coincides with wrapper id, if not, replace with parents id
    if data[ids.index('Wrapper_Id')] == Wrapper_Id:
        pass
    else:
        sheet.cell(row=coord[ids.index('Wrapper_Id')][0], column=coord[ids.index('Wrapper_Id')][1]).value=Wrapper_Id
    #check if it has an exercise id 
    Create_Entry = False
    if data[ids.index('Exercise_Id')]:
        print('already exists the entry in the excel file, checking for existance in the database')
        sqlcommand = 'SELECT * FROM [CalstContent].[dbo].[WrapperExercises] where Exercise_Id = ' + str(data[ids.index('Exercise_Id')]) + ' AND Wrapper_Id = '+str(Wrapper_Id)
        cursor.execute(sqlcommand)
        list = cursor.fetchall()
        if not list:
            Create_Entry = True
        else:
            print('db entry exists')
            Exercise_Id = data[ids.index('Exercise_Id')]
    else:
        Create_Entry = True

    if Create_Entry == True:
        print('creating an entry')
        cursor.execute('SELECT MAX(Id) AS maximum FROM Exercises')
        Exercise_Id = cursor.fetchall()[0][0]+1
        sqlcommand = 'INSERT INTO [dbo].[WrapperExercises] ([Wrapper_Id],[Exercise_Id]) VALUES '
        list = sqlcommand.split()[3].split(',')
        values = sqlcommand.split()[3].replace('[Wrapper_Id]',str(Wrapper_Id)).replace('[Exercise_Id]',str(Exercise_Id))
        sqlcommand = sqlcommand + values
        cursor.execute(sqlcommand)
        sheet.cell(row=coord[ids.index('Exercise_Id')][0], column=coord[ids.index('Exercise_Id')][1]).value = Exercise_Id
        
    else:
        pass
    return Create_Entry,Wrapper_Id,Exercise_Id



def main(folder,cursor):
    path = Path(folder)
    print(path.parent)
    structure_path  = Path(folder+'\\lessons_structure.xlsx')
    firstrun = True
    if structure_path.exists():
        #check actions column for the command "submit"
        wb_s = load_workbook(str(structure_path))
        sheet = wb_s['Lessons']
        to_submit = []
        print(str(structure_path))
        action_col = get_column(sheet=sheet, row = 1, name='Actions')
        exercise_col = get_column(sheet=sheet,row=2,name='Exercise_Id')
        folder_col = get_column(sheet=sheet,row=1,name='Folders')
        
        for cells in sheet.iter_cols(min_col=action_col,min_row=3, max_col=action_col, max_row=sheet.max_row):
            for cell in cells:
                if cell.value == 'submit':
                    to_submit.append(cell.row)
        print('rows to_submit: ', to_submit)
        
        
        for row in to_submit:
            Create_Entry, Wrapper_Id, Exercise_Id = check_exerciseid_in_structure(wb_s=wb_s, cursor=cursor, row=row)
            if Create_Entry:
                wb_s.save(filename=(str(structure_path))) 
            else:
                pass
            #open an exercise
            exercise_path = Path(sheet.cell(row=row, column=folder_col).value.replace('..',str(path.parent)))
            exercise_file = exercise_path/'exercise.xlsx'
            print(str(exercise_file))
            wb = load_workbook(str(exercise_file))
            sheet = wb['Exercise']

            names, columns,ranges = get_sheet_structure(sheet=sheet)
            print(names, columns,ranges)
            #check WrapperExercises values
            min_col,min_row, max_col, max_row = ranges[names.index('WrapperExercises')]
            ids = columns[names.index('WrapperExercises')]

            values = []
            for cells in sheet.iter_cols(min_col=min_col,min_row=3, max_col=max_col, max_row=3):
                for cell in cells:
                    values.append((cell.value,cell.column))
            print(values)
            Edit_Excel = False
            if values[ids.index('Wrapper_Id')][0] != Wrapper_Id:
                print('wrapper_id', str(Wrapper_Id),values[ids.index('Wrapper_Id')][0])
                sheet.cell(row=3, column=values[ids.index('Wrapper_Id')][1]).value=Wrapper_Id
                Edit_Excel = True
            else:
                pass
            if values[ids.index('Exercise_Id')][0] != Exercise_Id:
                print('exercise_id', str(Exercise_Id),values[ids.index('Exercise_Id')][0])
                sheet.cell(row=3, column=values[ids.index('Exercise_Id')][1]).value = Exercise_Id
                Edit_Excel = True
            else:
                pass
            
            min_col,min_row, max_col, max_row = ranges[names.index('Exercises')]
            ids = columns[names.index('Exercises')]

            values = []
            for cells in sheet.iter_cols(min_col=min_col,min_row=3, max_col=max_col, max_row=3):
                for cell in cells:
                    values.append((cell.value,cell.column))
            print(values)
        
           
            if values[ids.index('Id')][0] != Exercise_Id:
                print('exercise_id', str(Exercise_Id),values[ids.index('Id')][0])
                sheet.cell(row=3, column=values[ids.index('Id')][1]).value = Exercise_Id
                Edit_Excel = True
            else:
                pass
            
            
            
            
            
            
            
            
            
            
            if Edit_Excel:
               wb.save(str(exercise_file))
            else:
                pass
            

            


            
    

          
               



if __name__ == "__main__":
    ser_file = open('server_private.md','r')
    info = []
    for i in ser_file:
        info.append(i.split()[-1].replace("'",''))
    [server,database,username,password] = info
    
    #connect to the calst database
    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+'; UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()
    folder = 'C:\\Source\\Repos\\mysql-excel\\Spanish_course_styled\\'
    #folder = 'G:\\My Drive\\CALST_courses\\Spanish_course_styled\\'
    main(folder=folder, cursor=cursor)
    cnxn.commit()