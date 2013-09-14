# EDSA Report Generator

import os
import csv

# Home Working Directory
#os.chdir('c:/Projects/SV0002 - EDSA Report Generator/Test Directory')

# PowerCore Working Directory
# os.chdir('c:\Documents and Settings\Scott Vermeire\My Documents\Dropbox\EDSA Report Generator\Test Directory')"""

#Hard Drive Directory
os.chdir('f:\Personal Projects\SV0002 - EDSA Report Generator/Test Directory')

# Array Initialization for Reading File
Data = [0 for x in range(14)]

# Split Data from Headings
with open('ARCHEAT.csv') as csvfile:
    iFile = csv.reader(csvfile)

    i=0
    for row in iFile:
        if i == 0:
            Heading = row
            # print(Heading)
        else:
            Data[i] = row
            print(Data[i])
            # print(Data[i][2])
        i = i+1   

"""
count=1
while count < i:
    if Data[count][2] == '0.208':
        print(Data[count])
    else:
        pass
    count=count+1
"""

# Initialization for Sorting Lists
Equipment208V = [0 for x in range(i)]
Equipment240V = [0 for x in range(i)]
Equipment600V = [0 for x in range(i)]
Equipment480V = [0 for x in range(i)]
Equipment4160V = [0 for x in range(i)]
EquipmentUNSORTED = [0 for x in range(i)]
count=1
j=0
k=0
l=0
m=0
n=0
p=0

#List Sorting
while count < i:
    if Data[count][2]=='0.208':
        Equipment208V[j]= Data[count]
        j=j+1
    elif Data[count][2]=='0.240':
        Equipment240V[k]= Data[count]
        k=k+1
    elif Data[count][2]=='0.600':
        Equipment600V[l]= Data[count]
        l=l+1
    elif Data[count][2]=='0.480':
        Equipment480V[m]= Data[count]
        m=m+1
    elif Data[count][2]=='4.160':
        Equipment4160V[n]= Data[count]
        n=n+1
    else:
        EquipmentUNSORTED[p]= Data[count]
        p=p+1
    count=count+1

#Resulting List
if len(Equipment208V)!=0:
    print('\n\nEquipment208V: \n')
    for EachObject in Equipment208V:
        print('\n', EachObject)
if len(Equipment240V)!=0:
    print('\n\nEquipment240V: \n')
    for EachObject in Equipment240V:
        print('\n', EachObject)
if len(Equipment600V)!=0:
    print('\n\nEquipment600V: \n')
    for EachObject in Equipment600V:
        print('\n', EachObject)
if len(Equipment480V)!=0:
    print('\n\nEquipment480V: \n')
    for EachObject in Equipment480V:
        print('\n', EachObject)
if len(Equipment4160V)!=0:
    print('\n\nEquipment4160V: \n')
    for EachObject in Equipment4160V:
        print('\n', EachObject)
if len(EquipmentUNSORTED)!=0:
    print('\n\nEquipmentUNSORTED: \n')
    for EachObject in EquipmentUNSORTED:
        print('\n', EachObject)


