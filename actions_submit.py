
from ast import Index
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


def all_equal(iterable):
    g = groupby(iterable)
    return next(g, True) and not next(g, False)


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

def work_on_exercises_sheet(wb,cursor,cnxn, Wrapper_Id, Exercise_Id):
    sheet = wb['Exercise']
    Edit_Excel = False

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

    #go to exercises column, change exercise id and make a db entry if needed
    min_col,min_row, max_col, max_row = ranges[names.index('Exercises')]
    ids = columns[names.index('Exercises')]

    values = []
    for cells in sheet.iter_cols(min_col=min_col,min_row=3, max_col=max_col, max_row=3):
        for cell in cells:
            values.append((cell.value,cell.column))
    print(values)


    if values[ids.index('Id')][0] != Exercise_Id:
        print('exercise_id', str(Exercise_Id),values[ids.index('Id')][0])
        data_row = 3
        sheet.cell(row=data_row, column=values[ids.index('Id')][1]).value = Exercise_Id
        Edit_Excel = True
        #check whether entry exericse exists
        
    else:
        pass

    RendererId = values[ids.index('RendererId')][0]
    print(RendererId)


    sqlcommand = 'SELECT * FROM [CalstContent].[dbo].[Exercises] where Id = '+str(Exercise_Id)
    cursor.execute(sqlcommand)
    list = cursor.fetchall()
    Create_Entry = False
    if not list:
        Create_Entry = True
        print(Create_Entry)
    else:
        print('db entry exists')
        

    if Create_Entry:
        sqlcommand = 'INSERT INTO [dbo].[Exercises] ([Id],[RendererId]) VALUES '
        list = sqlcommand.split()[3].split(',')
        values = sqlcommand.split()[3].replace('[Id]',str(Exercise_Id)).replace('[RendererId]',str(RendererId))
        sqlcommand = sqlcommand + values
        print(sqlcommand)
        cursor.execute(sqlcommand)
    else:
        pass


    #fill in properties, check for exercise id and existence in the db
    print(names)

    for index,name in enumerate(names):
        if name == 'Properties':
            min_col, min_row,max_col,max_row=ranges[index]
            data_row = 3
            entry = []
            for cells in sheet.iter_cols(min_col=min_col,min_row=data_row, max_col=max_col, max_row=data_row):
                for cell in cells:
                    entry.append(cell.value)
            print("excel data: ",entry)

            if entry[columns[index].index('ExerciseId')]!= Exercise_Id:
                entry[columns[index].index('ExerciseId')] = Exercise_Id
                Edit_Excel = True
            else:
                pass
            #check whether entry exists
            sqlcommand = 'SELECT * FROM [CalstContent].[dbo].[Properties] where Id = ' + str(entry[columns[index].index('Id')]) + ' and ExerciseId = '+ str(entry[columns[index].index('ExerciseId')])
            print(sqlcommand)
            cursor.execute(sqlcommand)
            list = cursor.fetchall()
            print(list)
            if not list:
                Create_Entry = True
                Edit_Excel = True
                print(Create_Entry)
            else:
                print('db entry exists')

            if Create_Entry:
                cursor.execute('SELECT MAX(Id) AS maximum FROM Properties')
                Id = cursor.fetchall()[0][0]+1
                print(Id)
                entry[columns[index].index('Id')] = Id

                Id,WrapperId,ExerciseId,WordId,PronunciationId,Value,Key,Description,DialectId,Language_Id = entry
                sqlcommand ='INSERT INTO [dbo].[Properties] VALUES([Id],[WrapperId],[ExerciseId],[WordId],[PronunciationId],[Value],[Key],[Description],[DialectId],[Language_Id])' 

                for id in columns[index]:
                    sqlcommand = sqlcommand.replace('['+id+']','?')
                print(sqlcommand)
                cursor.execute(sqlcommand,Id,WrapperId,ExerciseId,WordId,PronunciationId,Value,Key,Description,DialectId,Language_Id)
                
                cnxn.commit()

            if Edit_Excel:
                for ind,ex in enumerate(entry):
                    data_row = 3
                    print(ind,ex)
                    sheet.cell(row=data_row, column = min_col+ind).value = ex
            else:
                pass
        else:
            print("No properties?")
    return Edit_Excel


def indices(lst, item):
    return [i for i, x in enumerate(lst) if x == item]

def get_entry_by_name(sheet, table_name, names, columns,ranges, row):
    entries = []
    table_indices = indices(names, table_name)
    for index in table_indices:
        entry = []
        min_col, min_row,max_col,max_row=ranges[index]
        range = (min_col, row, max_col, row) 
        for cells in sheet.iter_cols(min_col=min_col,min_row=row, max_col=max_col, max_row=row):
            for cell in cells:
                entry.append(cell.value)
        entries.append((entry,range,columns[index]))
    return entries


def replace_entry(sheet,entry, entry_range):
    min_col, min_row,max_col,max_row=entry_range
    index = 0
    for col in range(min_col,max_col+1):
        cell = sheet.cell(row=min_row, column=col)
        cell.value = entry[index]
        index = index + 1



def main(folder,cursor, cnxn):
    path = Path(folder)
    print(path.parent)
    structure_path  = Path(folder+'\\lessons_structure.xlsx')
    firstrun = True
    if structure_path.exists():
        #check actions column for the command "submit"
        wb_structure = load_workbook(str(structure_path))
        sheet = wb_structure['Lessons']
        to_submit = []
        print(str(structure_path))
        action_col = get_column(sheet=sheet, row = 1, name='Actions')
        folder_col = get_column(sheet=sheet, row=1, name='Folders')
        
        for cells in sheet.iter_cols(min_col=action_col,min_row=3, max_col=action_col, max_row=sheet.max_row):
            for cell in cells:
                if cell.value == 'submit':
                    to_submit.append(cell.row)
        print('rows to_submit: ', to_submit)
        
        
        for line in to_submit:
            #structure file changes
            Create_Entry, Wrapper_Id, Exercise_Id = check_exerciseid_in_structure(wb_s=wb_structure, cursor=cursor, row=line)
            if Create_Entry:
                wb_structure.save(filename=(str(structure_path))) 
            else:
                pass
            #open an exercise.xlsx

            exercise_path = Path(sheet.cell(row=line, column=folder_col).value.replace('..',str(path.parent)))
            exercise_file = exercise_path/'exercise.xlsx'
            print(str(exercise_file))
            wb_exercise = load_workbook(str(exercise_file))
            Edit_Excel = work_on_exercises_sheet(wb=wb_exercise,cursor=cursor, cnxn=cnxn, Wrapper_Id=Wrapper_Id,Exercise_Id=Exercise_Id)
            if Edit_Excel:
                 wb_exercise.save(str(exercise_file))
            else:
                pass
            
            case_vocab = False
            vocab = []
            case_mp = False
            mp = []
            case_nw = False
            nw = []
            for sheetname in wb_exercise.sheetnames:
                if 'Vocab' in  sheetname:
                    case_vocab = True
                    vocab.append(sheetname)
                elif 'MP' in sheetname:
                    #work on Vocab-Confusion Box
                    case_mp = True
                    mp.append(sheetname)
                elif 'Nonword' in sheetname:
                    case_nw = True
                    nw.append(sheetname)
                else:
                    pass

            if case_vocab:
                first_run = True
                print(vocab)
                max_info = []
                for sheet_name in vocab:
                    sheet = wb_exercise[sheet_name]
                    max_info.append(sheet.max_row)
                isequal = all_equal(max_info)
                if isequal:
                    max_row = max_info[0]
                else:
                    print('something wrong with number of rows, terminating')
                    exit()
                print(max_row)
                data_row = 3
                sheet_name = [s for s in vocab if 'Confusion' in s][0]
                print(sheet_name)
                sheet = wb_exercise[sheet_name]
                names,columns,ranges = get_sheet_structure(sheet = sheet)
                    
                for row in range(data_row,max_row+1):
                    #start with confusionbox    
                    if first_run == True:
                        first_run = False
                        print(row)                
                        
                        
                        entry,entry_range,entry_columns = get_entry_by_name(sheet=sheet,table_name='ConfusionBoxes', names=names,ranges=ranges,row=row,columns=columns)[0]
                        
                        Create_Entry = False
                        Edit_Excel = False
                        if entry[entry_columns.index('ExerciseId')]!= Exercise_Id:
                            entry[entry_columns.index('ExerciseId')] = Exercise_Id
                            Edit_Excel = True
                        else:
                            pass

                        print(entry)
                        if entry[entry_columns.index('Id')]:
                            sqlcommand = 'SELECT * FROM [CalstContent].[dbo].[ConfusionBoxes] WHERE Id = '+ str(entry[entry_columns.index('Id')]) +' AND ExerciseId = '+str(entry[entry_columns.index('ExerciseId')])
                            print(sqlcommand)
                            cursor.execute(sqlcommand)
                            list = cursor.fetchall()
                            print(list)
                            if not list:
                                Create_Entry = True 
                                Edit_Excel = True
                            else:
                                pass
                        else:
                            Create_Entry = True
                            Edit_Excel = True

                        if Create_Entry:
                            cursor.execute('SELECT MAX(Id) AS maximum FROM ConfusionBoxes')
                            Id = cursor.fetchall()[0][0]+1
                            print(Id)
                            entry[entry_columns.index('Id')] = Id
                            sqlcommand = 'INSERT INTO [dbo].[ConfusionBoxes] VALUES([Id],[ExerciseId],[DialectId],[CorrectTranscriptionId],[Bin])'
                            Id,ExerciseId,DialectId,CorrectTranscriptionId,Bin = entry
                            for id in entry_columns:
                                sqlcommand = sqlcommand.replace('['+id+']','?')
                            print(sqlcommand)
                            
                            cursor.execute(sqlcommand,Id,ExerciseId,DialectId,CorrectTranscriptionId,Bin)
                            cnxn.commit()
                        else:
                            print('db entry exists')

                        if Edit_Excel:
                            replace_entry(sheet=sheet,entry=entry,entry_range=entry_range)
                            wb_exercise.save(str(exercise_file))    
                        entry_confusionbox_out = entry
                        
                        
                    else:
                        print("here")
                        entry,entry_range,entry_columns = get_entry_by_name(sheet=sheet,table_name='ConfusionBoxes', names=names,ranges=ranges,row=row,columns=columns)[0]
                        if entry != entry_confusionbox_out:                          
                            replace_entry(sheet=sheet,entry=entry_confusionbox_out,entry_range=entry_range)
                            wb_exercise.save(str(exercise_file))
                    
                        
                    entry,entry_range,entry_columns = get_entry_by_name(sheet=sheet,table_name='Words', names=names,ranges=ranges,row=row,columns=columns)[0]
                    entry_words = entry
                    columns_words = entry_columns
                    Edit_Excel = False
                    print(entry, entry[entry_columns.index('Id')] )
                    replace_words = False
                    Create_Entry = False
                    if replace_words:
                        Create_Entry = True
                    else:
                        if not entry[entry_columns.index('Id')]:
                            Create_Entry = True
                        else:
                            sqlcommand = 'SELECT * FROM [CalstContent].[dbo].[Words] WHERE Id = '+str(entry[entry_columns.index('Id')])+ ' AND Text = \''+str(entry[entry_columns.index('Text')])+'\''
                            print(sqlcommand)
                            cursor.execute(sqlcommand)
                            list = cursor.fetchall()
                            if not list:
                                Create_Entry = True
                            else:
                                print('db entry exists, words')
                    
                    if Create_Entry:
                        cursor.execute('SELECT MAX(Id) AS maximum FROM Words')
                        Id = cursor.fetchall()[0][0]+1
                        Edit_Excel = True
                        print(Id)
                        entry[entry_columns.index('Id')] = Id
                        print(entry)
                        sqlcommand = 'INSERT INTO [dbo].[Words] VALUES([Id],[Text],[LanguageId])'
                        Id,Text,LanguageId = entry
                        for id in entry_columns:
                            sqlcommand = sqlcommand.replace('['+id+']','?')
                        print(sqlcommand)
                        entry_words = entry
                        
                        cursor.execute(sqlcommand,Id,Text,LanguageId)
                        cnxn.commit()
                    if Edit_Excel:
                            replace_entry(sheet=sheet,entry=entry_words,entry_range=entry_range)
                            wb_exercise.save(str(exercise_file)) 

                    #transcriptions
                    Create_Entry = False
                    print(entry_words)
                    
                    entry,entry_range,entry_columns = get_entry_by_name(sheet=sheet,table_name='Transcriptions', names=names,ranges=ranges,row=row,columns=columns)[0]                    
                    print(entry)
                    
                    print(entry[entry_columns.index('WordId')],entry_words[columns_words.index('Id')])
                    
                    if entry[entry_columns.index('WordId')]!= entry_words[columns_words.index('Id')]:
                        entry[entry_columns.index('WordId')]= entry_words[columns_words.index('Id')]
                        Create_Entry = True
                    else:
                        sqlcommand = 'SELECT * FROM [CalstContent].[dbo].[Transcriptions] WHERE Id = '+str(entry[entry_columns.index('Id')])+ ' AND WordId = \''+str(entry[entry_columns.index('WordId')])+'\''
                        print(sqlcommand)
                        cursor.execute(sqlcommand)
                        list = cursor.fetchall()
                        if not list:
                            Create_Entry = True
                        else:
                            print('db entry exists, words')
                    print(entry)
                    print(Create_Entry)
                    



                    #entry_TranscriptionConfusionBoxes = get_entry_by_name(sheet=sheet,table_name='TranscriptionConfusionBoxes', names=names,ranges=ranges,row=row,columns=columns)
                    #print(entry_TranscriptionConfusionBoxes)

                    #entry_Words = get_entry_by_name(sheet=sheet,table_name='Words', names=names,ranges=ranges,row=row,columns=columns)
                    #print(entry_Words)

                    #entry_Transcriptions = get_entry_by_name(sheet=sheet,table_name='Transcriptions', names=names,ranges=ranges,row=row, columns=columns)
                    #print(entry_Transcriptions[0][2])

                    
                    

                    
                    

                
                
                
                
                
            elif case_mp:
                print(mp)
            elif case_nw:
                print(nw)
            else:
                pass

    else:
        print("No exercise file here. quitting")
        exit()                




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
    main(folder=folder, cursor=cursor, cnxn=cnxn)
    cnxn.commit()