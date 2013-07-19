"""EDSA Report Generator"""

import os
import csv

"""Temporary Working Directory"""

os.chdir('c:/Projects/EDSA Report Generator/Test Directory')

with open('ARCHEAT.csv') as csvfile:
     iFile = csv.reader(csvfile)
     for row in iFile:
         print(', '.join(row))
