- ``pip install -r requirements.txt  (do it preferably in venv env)``
  -  **NB**: do not push server.md with your pass to the git repo, I use server_private.md which is ignored in [.gitignore](.gitignore)
- [sql_to_folders_and_excel.py](sql_to_folders_and_excel.py) -- this script reads from the CALST database (the path and the password is specified in the [server.md](server.md) where a password should be inserted) and interprets this database as the list of folders with excel files like wrapper.xlsx and exercise.xlsx, following lesson structures
- [diff_folder_and_mysql.py](diff_folder_and_mysql.py)  -- this script compares a modified excel file with the calst database:
  - usage sample: ``python.exe .\vocabulary_case\diff_folder_and_mysql.py -f 'C:\Source\Repos\python_tools\Spanish_course_styled\Beginner\Lesson 1\Numbers 1'``
    - can be used without the option ``-f``, in this case should be in 
        
         ```
         if __name__ == "__main__":
            parser = argparse.ArgumentParser(prog='python diff_folder_and_mysql.py -f foldername')
            parser.add_argument('-f',dest='folder')
            args = parser.parse_args()
            if args.folder:
                main(folder = args.folder)
            else:
                folder = 'C:\Source\Repos\python_tools\Spanish_course_styled\Beginner\Lesson 1\The alphabet'
                main(folder=folder) 
 [check_db_vs_excel_consistency.py](check_db_vs_excel_consistency.py) -- checks some columns to be consistent between each other    

