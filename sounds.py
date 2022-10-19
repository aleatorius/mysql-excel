
import pyodbc 
from pathlib import Path
import os
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from sql_to_folders_and_excel import get_style

def get_header(cursor,table):
    sqlcom = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'"+table+"'"
    cursor.execute(sqlcom)  
    data = list(cursor.fetchall())
    header = [i[3] for i in data]
    header_0 = len(data)*[table]
    return header_0,header


def main():
    ser_file = open('server_private.md','r')
    info = []
    for i in ser_file:
        info.append(i.split()[-1].replace("'",''))
    [server,database,username,password] = info

    #connect to the calst database

    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+'; UID='+username+';PWD='+ password)
    cursor = cnxn.cursor()

    #wrapper corresponds to a second levels structures like lessons 1, 2,3 

      
    sqlcom = "SELECT * FROM [CalstContent].[dbo].[Sounds]"
    cursor.execute(sqlcom)  
    data = [list(i) for i in cursor.fetchall()] 
    

    header_0,header = get_header(cursor=cursor,table='Sounds')
    print(header_0,header)

    wbout = Workbook()
    sheet = wbout.active
    sheet.title = "Sounds"
    header_cells = []
    header_0,header = get_header(cursor=cursor, table='Sounds')
    sheet.append(header_0)
    sheet.append(header)
    header_cells.append(len(header))

    for i in data:
        print(i)
        sheet.append(i)
    get_style(header_cells,sheet)
    wbout.save("sounds.xlsx")
    wbout.close()


if __name__ == "__main__":
    main()