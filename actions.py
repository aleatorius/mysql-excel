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
import diff_folder_and_mysql

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
    index = columns[0].index('Name')
    level = (entry,entry[index],names,columns)
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
                    if subdir.count(os.path.sep) == local_depth:
                        local_level = dirs
    if local_level:
        return local_level
    else:
        return False

def get_style(header_cells,sheet):
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter # Get the column name
        for cell in col:
            if cell.row > 1:
                try: # Necessary to avoid error on empty cells
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))+2.7
                except:
                    pass
            else:
                pass
        adjusted_width = (max_length) 
        sheet.column_dimensions[column].width = adjusted_width   
    start=1
    twocolor = ['00CCFFCC','00FFFF99','00C0C0C0','0033CCCC']
    ccolor = twocolor[0]
    twopattern = ['lightDown','darkGray']
    for i in header_cells:
        cell_header = sheet.cell(row=1, column=start)
        double = Side(border_style="double", color="00008000")
        cell_header.border = Border(top=double, left=double, right=double, bottom=double)
        cell_header.font  = Font(b=True, color="00008000", size = 8)
        cell_header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False,  shrink_to_fit=False)
        sheet.merge_cells(start_row=1, start_column=start, end_row=1, end_column=start+i-1)
        for j in range(i):
            cell_h = sheet.cell(row=2, column=start+j)
            cell_h.font  = Font(b=True, color="00008000", size = 10)
            cell_h.alignment = Alignment(horizontal="general", vertical="bottom", wrap_text=False,  shrink_to_fit=False)
            if j == 0:
                cell_h.border = Border(left=double, bottom=double)
            else:
                cell_h.border = Border(bottom=double)

        
        ppattern = twopattern[0]
        for l in range(3,sheet.max_row+1):
            for j in range(i):
                cell_ordinary = sheet.cell(row = l,column=start+j)
                cell_ordinary.fill = PatternFill(ppattern, fgColor=ccolor)
                thin = Side(border_style="thin", color="000000")
                cell_ordinary.border = Border(top=thin, left=thin, right=thin, bottom=thin)
            if ppattern == twopattern[0]:
                ppattern = twopattern[1]
            else:
                ppattern = twopattern[0]
        
        start = start + i
        if ccolor == twocolor[0]:
            ccolor = twocolor[1]
        else:
            ccolor = twocolor[0]
         
    if header_cells[-1] == 1:
        cell_h.border = Border(bottom=double, right=double,left=double)   
    else:
        cell_h.border = Border(bottom=double, right=double)   
          
    



def main():
  

    wb_s = load_workbook(filename='C:\Source\Repos\mysql-excel\Spanish_course\lessons_structure.xlsx')
    sheet = wb_s.active
    print(sheet.title)
    names,columns,ranges = get_sheet_structure(sheet = sheet)
    
    print(names, columns,ranges, ranges[names.index('Actions')])
    (col_low, row_low, col_high, row_high) = ranges[names.index('Actions')]
    to_diff = []
    for cells in sheet.iter_cols(min_col=col_low,min_row=row_low+2, max_col=col_high, max_row=sheet.max_row):
        for cell in cells:
            if cell.value:
                if cell.value.lower() == 'diff':
                    print(cell.row)
                    (col_l, row_l, col_h, row_h) = ranges[names.index('Folders')]
                    cell_f = sheet.cell(column=col_l, row=cell.row)
                    to_diff.append(cell_f.value.replace('..','C:\Source\Repos\mysql-excel'))
    print(to_diff)
    for todiff in to_diff:
        diff_folder_and_mysql.main(folder=todiff, output_diff='test.txt')

    
  
if __name__ == "__main__":
    main()