# Imported Modules
import os
import csv
import xlwt
import datetime as DT
from xlrd import open_workbook

Working_Directory = 'c:/Projects/SV0002 - EDSA Report Generator/Test Directory'
os.chdir(Working_Directory)

w = xlwt.Workbook()
ws = w.add_sheet('Image')
ws.insert_bitmap('logo.bmp', 0, 0)
w.save('images.xls')
