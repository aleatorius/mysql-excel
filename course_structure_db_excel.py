from ctypes import Structure
from natsort import natsorted
import pyodbc 
import os
import argparse
from itertools import groupby
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_interval
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from diff_folder_and_mysql import get_sheet_structure

def get_wrapper(wrapper):
    level = []
    wb = load_workbook(filename = wrapper)
    sheet = wb['Wrappers']
    names,columns,ranges = get_sheet_structure(sheet = sheet)
    data_row = sheet.max_row
    for data in zip(names,columns,ranges):
        (col_low, row_low, col_high, row_high) = data[-1]
        entry = []
        for cells in sheet.iter_cols(min_col=col_low,min_row=data_row, max_col=col_high, max_row=data_row):
            for cell in cells:
                entry.append(cell.value)
        #print(entry)
    #print(columns)
    index = columns[0].index('Name')
    level.append((entry,entry[index]))
    return level

def get_exercise(wrapper):
    level = []
    wb = load_workbook(filename = wrapper)
    sheet = wb['Exercise']
    names,columns,ranges = get_sheet_structure(sheet = sheet)
    data_row = sheet.max_row
    id_col_row = 2
    exercise_names = []
    entry = []
    for data in zip(names,columns,ranges):
        if data[0] == 'WrapperExercises':
            (col_low, row_low, col_high, row_high) = data[-1]
            for cells in sheet.iter_cols(min_col=col_low,min_row=data_row, max_col=col_high, max_row=data_row):
                for cell in cells:
                    entry.append(cell.value)
            
        if data[0] == 'Properties':
            (col_low, row_low, col_high, row_high) = data[-1]
            
            
            for cells in sheet.iter_cols(min_col=col_low,min_row=id_col_row, max_col=col_high, max_row=id_col_row):
                for cell in cells:
                    if cell.value == 'Key':
                        key = cell.col_idx
                    if cell.value == 'Value':
                        value =  cell.col_idx 

            data_cell = sheet.cell(column=key,row=data_row)
            if data_cell.value == 'ExerciseName':
                name_cell = sheet.cell(column=value,row=data_row)
                exercise_names.append(name_cell.value)
            
        

    if len(exercise_names) == 0 or len(exercise_names)>1:
        print("error, too many names or none, quitting")
        exit()
    else:
        return entry,exercise_names[0] 
    
    return level

def get_wrapper_dirs(folder):
    local_depth = os.path.abspath(folder).count(os.path.sep)
    local_level = []
    for subdir, dirs, files in os.walk(folder, topdown=True):
        for name in files:
            filepath = subdir + os.sep + name
            if filepath.endswith(".xlsx"):
                if name == 'wrapper.xlsx':
                    #print(filepath.split('\\'),name)
                    if subdir.count(os.path.sep) == local_depth:
                        local_level = dirs
    if local_level:
        return local_level
    else:
        return False

def main(folder):
    warnings_file = open('warnings.txt','w')
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
    abspath = os.path.abspath(folder_path)
    
    first_level = abspath.count(os.path.sep)
    wrapper = folder_path/'wrapper.xlsx'
    level=get_wrapper(wrapper=wrapper)
    print('Entry Point wrapper :',level)
    
    
    wbout = Workbook()
    sheet_wbout = wbout.active
    sheet_wbout.title = "Lessons"
    
   
    #level_0
    sturcture=[]
    level = get_wrapper_dirs(folder_path)
    for chapter in level:
        print(chapter)
        
        folder = folder_path/chapter
        lessons = get_wrapper_dirs(folder)
    
        wrapper = folder/'wrapper.xlsx'
        level=get_wrapper(wrapper=wrapper)
        print('Chapter wrapper :',level)
        
        for lesson in natsorted(lessons):
            print(lesson)
            folder = folder_path/chapter/lesson
            wrapper = folder/'wrapper.xlsx'
            level=get_wrapper(wrapper=wrapper)
            print('Lesson wrapper :',level)
            sublessons = get_wrapper_dirs(folder)
            for sublesson in natsorted(sublessons):
                print(sublesson)
                folder=folder_path/chapter/lesson/sublesson
                sub_sublessons = get_wrapper_dirs(folder)
                if sub_sublessons != False:
                    folder=folder_path/chapter/lesson/sublesson
                    wrapper = folder/'wrapper.xlsx'
                    level=get_wrapper(wrapper=wrapper)
                    print('sublesson wrapper :',level)
                    for sub_sublesson in sub_sublessons:
                        #print(chapter,": ",lesson,"-> ", sublesson,"-->>", sub_sublesson)
                        sturcture.append((chapter, lesson,sublesson,sub_sublesson))
                        folder = folder_path/chapter/lesson/sublesson/sub_sublesson
                        wrapper = folder/'exercise.xlsx'
                        level = get_exercise(wrapper=wrapper)
                        print('exercise :',level)

                else:
                    #print(chapter,": ",lesson,"-> ", sublesson)
                    sturcture.append((chapter,lesson,sublesson))
                    folder = folder_path/chapter/lesson/sublesson
                    wrapper = folder/'exercise.xlsx'
                    level = get_exercise(wrapper=wrapper)
                    print('exercise :',level)
  
    
  
if __name__ == "__main__":
    parser = argparse.ArgumentParser(prog='python course_structure_db_excel.py -f foldername')
    parser.add_argument('-f',dest='folder')
    args = parser.parse_args()
    if args.folder:
        main(folder = args.folder)
    else:
        #folder='C:\Source\Repos\python_tools\Spanish_course_styled\Beginner\Lesson 1\The alphabet'
        #folder = 'C:\Source\Repos\python_tools\Spanish_course_styled\Beginner\Lesson 1\\Numbers 1'
        #folder = 'C:\Source\Repos\python_tools\Spanish_course_styled\Beginner\Lesson 2\\Nationalities'
        folder = 'C:\Source\Repos\mysql-excel\Spanish_course'
        main(folder=folder)