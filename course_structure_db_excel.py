from ctypes import Structure
from textwrap import indent
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
    wb = load_workbook(filename = wrapper)
    sheet = wb['Wrappers']
    names,columns,ranges = get_sheet_structure(sheet = sheet)
    data_row = sheet.max_row
    entry = []
    for data in zip(names,columns,ranges):
        (col_low, row_low, col_high, row_high) = data[-1]
        for cells in sheet.iter_cols(min_col=col_low,min_row=data_row, max_col=col_high, max_row=data_row):
            for cell in cells:
                entry.append(cell.value)
    #print(columns)
    #print(entry)
    index = columns[0].index('Name')
    level = (entry,entry[index],names,columns)
    #print(level)
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
    cols=[]
    for data in zip(names,columns,ranges):
        if data[0] == 'WrapperExercises':
            (col_low, row_low, col_high, row_high) = data[-1]
            for cells in sheet.iter_cols(min_col=col_low,min_row=data_row, max_col=col_high, max_row=data_row):
                for cell in cells:
                    entry.append(cell.value)
            header_ex_id = data[1]
            header_ex = len(header_ex_id)*[data[0]]
            
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
        return entry,exercise_names[0],header_ex_id, header_ex
    
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
    inputfolder = folder
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
    

    name_structure=[]
    wrapper_structure = []
    exercise_structure = []
    structure = []
    first_level = abspath.count(os.path.sep)
    wrapper = folder_path/'wrapper.xlsx'
    level=get_wrapper(wrapper=wrapper)
    print(level[2],level[3][0])
    header_0_wrapper = len(level[3][0])*level[2]    
    header_wrapper = level[3][0]
    
    newline = []
    newline.append(level[1])
    print('Entry Point wrapper :',level)

    name_structure.append(newline)
    
    wrapper_structure.append(level[0])

    print(name_structure, wrapper_structure)
    
    
    wbout = Workbook()
    sheet_wbout = wbout.active
    sheet_wbout.title = "Lessons"
    header_basic = ['Level 0','Level 1', 'Level 2', 'Level 3','Level 4','Action']
    header = header_basic
    header_0 = (len(header)-1)*['Level names']
    names_cols = len(header)-1
    header_0.append('')
    
   
    #sheet_wbout.append(newline)

    
    #level_0
    
    first_time = True
    level = get_wrapper_dirs(folder_path)
    for chapter in level:
        print(chapter)
        newline = ['']
        folder = folder_path/chapter
        lessons = get_wrapper_dirs(folder)
    
        wrapper = folder/'wrapper.xlsx'
        level=get_wrapper(wrapper=wrapper)

        #print('Chapter wrapper :',level)
        newline.append(level[1])
        #print(newline)
        #
        
        name_structure.append(newline)
        wrapper_structure.append(level[0])
        #print(name_structure)
        
        
        
        
        for lesson in natsorted(lessons):
            print(lesson)
            folder = folder_path/chapter/lesson
            wrapper = folder/'wrapper.xlsx'
            level=get_wrapper(wrapper=wrapper)

            newline = ['','']
            newline.append(level[1])
            name_structure.append(newline)
            wrapper_structure.append(level[0])
            #
        
            #print(name_structure)
            

            #print('Lesson wrapper :',level)
            sublessons = get_wrapper_dirs(folder)
            for sublesson in natsorted(sublessons):
                #print(sublesson)
                folder=folder_path/chapter/lesson/sublesson
                sub_sublessons = get_wrapper_dirs(folder)
                if sub_sublessons != False:
                    folder=folder_path/chapter/lesson/sublesson
                    wrapper = folder/'wrapper.xlsx'
                    level=get_wrapper(wrapper=wrapper)

                    newline = 3*['']
                    newline.append(level[1])
                    name_structure.append(newline)
                    wrapper_structure.append(level[0])
                    #
        
                    #print(name_structure)
                    

                    #print('sublesson wrapper :',level)
                    for sub_sublesson in sub_sublessons:
                        #print(chapter,": ",lesson,"-> ", sublesson,"-->>", sub_sublesson)
                        structure.append((chapter, lesson,sublesson,sub_sublesson))
                        folder = folder_path/chapter/lesson/sublesson/sub_sublesson
                        wrapper = folder/'exercise.xlsx'
                        level = get_exercise(wrapper=wrapper)
                        #print('exercise :',level)
                        if first_time == True:
                            header_ex_id = level[2]
                            header_ex = level[3] 
                            first_time = False
                        else:
                            pass
                        newline = 4*['']
                        newline.append(level[1])
                        name_structure.append(newline)
                        wrapper_structure.append(level[0])
                        #
                        #print(name_structure)
                        

                else:
                    #print(chapter,": ",lesson,"-> ", sublesson)
                    structure.append((chapter,lesson,sublesson))
                    folder = folder_path/chapter/lesson/sublesson
                    wrapper = folder/'exercise.xlsx'
                    level = get_exercise(wrapper=wrapper)
                    #print('exercise :',level)
                    newline = 3*['']
                    newline.append(level[1])
                    wrapper_structure.append(level[0])
                    #
                    name_structure.append(newline)

    header_0 = header_0 + header_0_wrapper+header_ex
    header = header + header_wrapper+header_ex_id
    print(header_0)
    print(header)
    
    sheet_wbout.append(header_0)
    sheet_wbout.append(header)                
    for line,info in zip(name_structure,wrapper_structure):
        print(line)
        for ind,ex in enumerate(line):
            if ex != 0:
                nonzero_index = ind
        if len(info) > 2:
            linen = line+(len(header_basic)-ind-1)*['']+info
        else:
            linen = line+(len(header_basic)-ind-1+len(header_wrapper))*['']+info
        sheet_wbout.append(linen)
    sheet_wbout.merge_cells(start_row=1, start_column=1, end_row=1, end_column=names_cols)
    sheet_wbout.merge_cells(start_row=1, start_column=names_cols+2, end_row=1, end_column=names_cols+1+len(header_wrapper))
    sheet_wbout.merge_cells(start_row=1, start_column=names_cols+2+len(header_wrapper), end_row=1, end_column=names_cols+1+len(header_wrapper)+len(header_ex))
    wbout.save(inputfolder+'\lessons_structure.xlsx')
    wbout.close()        
    

  
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