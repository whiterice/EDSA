"""EDSA Report Generator"""

import os
import csv

"""Home Working Directory"""
"""os.chdir('c:/Projects/EDSA Report Generator/Test Directory')"""

"""PowerCore Working Directory"""
os.chdir('c:\Documents and Settings\Scott Vermeire\My Documents\Dropbox\EDSA Report Generator\Test Directory')

"""Array Initialization"""
Data = [0 for x in range(14)]

with open('ARCHEAT.csv') as csvfile:
    iFile = csv.reader(csvfile)

    i=0
    
    for row in iFile:
        if i == 0:
            print('# '.join(row))
            Heading = row
            print(Heading)
        else:
            print(', '.join(row))
            Data[i] = row
            print(Data[i])
            
        i = i+1
