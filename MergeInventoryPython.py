
setx PATH "%PATH%;C:\Users\c00298579\AppData\Local\Programs\Python\Python39\Scripts"

# install python pip
# https://stackoverflow.com/questions/23708898/pip-is-not-recognized-as-an-internal-or-external-command
pip install pyodbc

#How to Create a Table in SQL Server using Python
#https://datatofish.com/create-table-sql-server-python/

#How to Make Inserts Into SQL Server 100x faster with Pyodbc
#https://towardsdatascience.com/how-i-made-inserts-into-sql-server-100x-faster-with-pyodbc-5a0b5afdba5

#Connecting to MS SQL Server with Windows Authentication using Python?
#https://stackoverflow.com/questions/16515420/connecting-to-ms-sql-server-with-windows-authentication-using-python



# Inventory:
import pandas as pd
import numpy as np
#import seaborn as sns
#import matplotlib.pyplot as plt
#import xml.etree.ElementTree as ET
#import openpyxl
#import pandas as pd
import glob
import os
#import xlsxwriter


def append_csv(data_path, export_path, filtelocation, ColumnName):
   # For loop to read dataframe
   # https://stackoverflow.com/questions/28669482/appending-pandas-dataframes-generated-in-a-for-loop#comment45637397_28670223
   # https://sparkbyexamples.com/pandas/pandas-concat-dataframes-explained/#:~:text=Use%20pandas.,append%20one%20DataFrame%20with%20another.
   #data_path = r'D:\PANDA\20231215\Inventory_Board' # use your path
   dir_list = os.listdir(data_path)
   appended_data = []
   for file in dir_list:
      df_temp = pd.read_csv(os.path.join(data_path , file), index_col=None, header=0)
      appended_data.append(df_temp)
   # https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html
   # https://datalore-forum.jetbrains.com/t/pandas-to-excel-where-does-it-go/360/2
   appended_data = pd.concat(appended_data)
   # appended_data.to_excel('appended.xlsx')
   #export_path = r'D:\PANDA\20231215' # use your path
   appended_data.to_excel(os.path.join(export_path, 'appended.xlsx'))
   # Filter Dataframe
   if not filtelocation == "":
      filter_df = pd.read_excel(filtelocation, header=None)
      filter_list = filter_df[0].tolist()
      df_filtered = appended_data[appended_data[ColumnName].isin(filter_list)]
      df_filtered.to_excel(os.path.join(filter_file_path, 'appended_Filtered.xlsx'))

def append_xlsx_specific_sheet(data_path, export_path, specific_Sheet, export_name, filtelocation, ColumnName):
   files = os.listdir(data_path)
   files_xlsx = [f for f in files if f[-4:] == 'xlsx']
   df_list = []
   for f in files_xlsx:
      data = pd.read_excel(os.path.join(data_path, f), sheet_name=specific_Sheet,header=1)
      df_list.append(data)
      print('file: ' + f + 'loaded')
   df = pd.concat(df_list)
   #df.to_excel(os.path.join(export_path, 'append_xlsx_specific_sheet.xlsx'))
   df.to_excel(os.path.join(export_path, export_name))
   if not filtelocation == "":
      # Filter Dataframe
      #filter_file_path = r'D:\PANDA\20231215' # use your path
      #filter_file = 'filter_List.xlsx'
      #filter_df = pd.read_excel(os.path.join(filter_file_path, filter_file), header=None)
      #df_filtered = df.isin(filter_df)
      filter_df = pd.read_excel(filtelocation, header=None)
      filter_list = filter_df[0].tolist()
      df_filtered = df[df[ColumnName].isin(filter_list)]
      df_filtered.to_excel(os.path.join(filter_file_path, specific_Sheet+'_df_Filtered.xlsx'))


# call function to concatenate csv files and export in xlsx format 
#path2 = r'D:\PANDA\20231215\Inventory_Board' # use your path
path2 = r'D:\PANDA\20231215\Inventory_Antenna' # use your path
path3 = r'D:\PANDA\20231215' # use your path
#append_csv(path2, path3, filter)
filter_file_path = r'D:\PANDA\20231215' # use your path
filter_file = 'filter_List.xlsx'
filter_file_path_and_file = ''
filter_file_path_and_file = os.path.join(filter_file_path, filter_file)
ColumnName = 'NEName'
append_csv(path2, path3, filter_file_path_and_file, ColumnName)



# call function to concatenate xlsx specific sheet "SECTOREQM" in files and export in xlsx format 
path2 = r'D:\PANDA\20231215\Cell' # use your path
path3 = r'D:\PANDA\20231215' # use your path
specific_Sheet = "SECTOREQM"
export_name = 'append_xlsx_SECTOREQM_sheet.xlsx'
filter_file_path = r'D:\PANDA\20231215' # use your path
filter_file = 'filter_List.xlsx'
filter_file_path_and_file = ''
filter_file_path_and_file = os.path.join(filter_file_path, filter_file)
ColumnName = '*Name'
append_xlsx_specific_sheet(path2, path3, specific_Sheet, export_name, filter_file_path_and_file, ColumnName)
# call function to concatenate xlsx specific sheet "CELL" in files and export in xlsx format 
path2 = r'D:\PANDA\20231215\Cell' # use your path
path3 = r'D:\PANDA\20231215' # use your path
specific_Sheet = "CELL"
export_name = 'append_xlsx_CELL_sheet.xlsx'
filter_file_path = r'D:\PANDA\20231215' # use your path
filter_file = 'filter_List.xlsx'
filter_file_path_and_file = ''
filter_file_path_and_file = os.path.join(filter_file_path, filter_file)
ColumnName = '*eNodeB Name'
append_xlsx_specific_sheet(path2, path3, specific_Sheet, export_name, filter_file_path_and_file, ColumnName)
# call function to concatenate xlsx specific sheet "EUCELLSECTOREQM" in files and export in xlsx format 
path2 = r'D:\PANDA\20231215\Cell' # use your path
path3 = r'D:\PANDA\20231215' # use your path
specific_Sheet = "EUCELLSECTOREQM"
export_name = 'append_xlsx_EUCELLSECTOREQM_sheet.xlsx'
filter_file_path = r'D:\PANDA\20231215' # use your path
filter_file = 'filter_List.xlsx'
filter_file_path_and_file = ''
filter_file_path_and_file = os.path.join(filter_file_path, filter_file)
ColumnName = '*eNodeB Name'
append_xlsx_specific_sheet(path2, path3, specific_Sheet, export_name, filter_file_path_and_file, ColumnName)



# File:cell, Sheet:SECTOREQM, 	Column:"*Name", "Sector Equipment ID",		Sector Equipment Antenna, Field:0,130,0,R0A,RXTX_MODE,MASTER;0,130,0,R0B,RXTX_MODE,MASTER




table = pd.pivot_table(appended_data,
                       index='NEName',
                       columns='Board Name',
                       values='NEType',
                       aggfunc=[np.count_nonzero],
                       fill_value=0
                       )
table.to_csv (r'D:\PANDA\20231215\Board_Name_Pivot.csv', index = True) # place 'r' before the path name

table = pd.pivot_table(appended_data,
                       index='NEName',
                       columns='Board Type',
                       values='NEType',
                       aggfunc=[np.count_nonzero],
                       fill_value=0
                       )
table.to_csv (r'D:\PANDA\20230502\Board_Type_Pivot.csv', index = True) # place 'r' before the path name

table = pd.pivot_table(appended_data,
                       index='NEName',
                       columns='Manufacturer Data',
                       values='NEType',
                       aggfunc=[np.count_nonzero],
                       fill_value=0
                       )
table.to_csv (r'D:\PANDA\20230502\Board_Manufacturer_Data_Pivot.csv', index = True) # place 'r' before the path name
