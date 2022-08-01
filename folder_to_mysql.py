from pydoc import ispackage
import pyodbc 
from pathlib import Path
import os
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl import load_workbook

def main():
    folder = 'C:\Source\Repos\python_tools\Spanish_course_styled\Beginner\Lesson 3\The house'
    folder_path = Path(folder)
    exercise = folder_path/'exercise.xlsx'

    if exercise.exists():
        print("ok, exercise exists in this folder")
        wb = load_workbook(filename = exercise)
        for i in wb.sheetnames:
            print(i)
            sheet = wb[i]
            merged_ranges = sheet.merged_cells.ranges
            database = []
            for range in merged_ranges:
                print(range, range.bounds)
                (col_low, row_low, col_high, row_high) = range.bounds
                
                database.append(sheet.cell(row_low,col_low).value)
            print(database)
           
if __name__ == "__main__":
    main()