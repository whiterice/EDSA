"""EDSA Report Generator"""

import os
import csv

"""Temporary Working Directory"""

os.chdir('c:/Projects/EDSA Report Generator/Test Directory')

with open('ARCHEAT.csv') as csvfile:
    iFile = csv.reader(csvfile)

    i=0
    
    for row in iFile:
        if i == 0:
            print('# '.join(row))

        else:
            print(', '.join(row))

        i = i+1
