"""EDSA Report Generator"""

import os
import csv

"""Home Working Directory"""
os.chdir('c:/Projects/EDSA Report Generator/Test Directory')

"""PowerCore Working Directory"""
os.chdir('c:/Documents and Settings/Scott Vermeire/My Documents/Dropbox/EDSA Report Generator/Python Code/Test Directory')

with open('ARCHEAT.csv') as csvfile:
    iFile = csv.reader(csvfile)

    i=0
    
    for row in iFile:
        if i == 0:
            print('# '.join(row))

        else:
            print(', '.join(row))

        i = i+1
