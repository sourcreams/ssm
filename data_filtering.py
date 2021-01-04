import openpyxl
import pyodbc

conn_String = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\6408284\Desktop\새 폴더 (2)\moduleBook.accdb;'

pai_location = 2
vehicle_location = 1
module_location = 3
value_location = 4

def search_duplicate(sht_name):
    for i in range(2, sht_name.max_row+1):

