from textwrap import indent
from natsort import natsorted
from itertools import groupby
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_interval
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from diff_folder_and_mysql import get_sheet_structure
import diff_folder_and_mysql
import course_structure_db_excel





def main(structure_file):
  
    structure_path  = Path(structure_file)
    if structure_path.exists():

        wb_s = load_workbook(filename='C:\Source\Repos\mysql-excel\Spanish_course\lesson_structure.xlsx')
        sheet = wb_s.active
        names,columns,ranges = get_sheet_structure(sheet = sheet)
        (col_low, row_low, col_high, row_high) = ranges[names.index('Actions')]
        to_diff = []
        to_add = []
        for cells in sheet.iter_cols(min_col=col_low,min_row=row_low+2, max_col=col_high, max_row=sheet.max_row):
            for cell in cells:
                if cell.value:
                    if cell.value.lower() == 'diff':
                        (col_l, row_l, col_h, row_h) = ranges[names.index('Folders')]
                        cell_f = sheet.cell(column=col_l, row=cell.row)
                        to_diff.append(cell_f.value.replace('..','C:\Source\Repos\mysql-excel'))

        for todiff in to_diff:
            diff_folder_and_mysql.main(folder=todiff, output_diff='test.txt')
    else:
        print(str(structure_path.parent))
        print(str(structure_path.name))
        
        course_structure_db_excel.main(folder=str(structure_path.parent), output=str(structure_path.name ))
        print("No file with lesson info, have created one. You can ad actions here: "+ str(structure_path))
    
  
if __name__ == "__main__":
    main(structure_file='C:\\Source\\Repos\\mysql-excel\\Spanish_course\\testic.xlsx')