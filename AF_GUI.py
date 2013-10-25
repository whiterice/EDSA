import EDSA_Report_Generator
import os

print "****************************\n*Arc Flash Report Generator*\n****************************"

GUI_dir=os.getcwd()
os.chdir('../Test Directory/S2777-City Of London Arc Flash 2013')
Available_Options=os.listdir(os.getcwd())

Full_Path = os.getcwd()
Full_Path_List = Full_Path.split('\\')

Main_dir = Full_Path_List[len(Full_Path_List)-1]

print '\n{!s}:'.format(Main_dir)

Available_dir =[]
for dir_i in Available_Options:
    if os.path.isdir(dir_i):
        Available_dir.append(dir_i)

for i,v in enumerate(Available_dir):
    print "{:d}) {!s}".format(i,v)

print '\nSelect directory using the associated integer:'
Available_dir_index = int(raw_input())

w_dir = Available_dir[Available_dir_index]

print '\nYou Have Selected: {!s}'.format(w_dir)
os.chdir(w_dir)



if w_dir.find('-')!=-1:
    temp = w_dir.split('-')
    Job_Number = temp[0]
else:
    Job_Number = '<Unknown>'

if Main_dir.find('-')!=-1:
    temp = Main_dir.split('-')
    Customer_Company = temp[1]
else:
    Customer_Company = '<Unknown>'

if w_dir.find('-')!=-1:
    temp = w_dir.split('-')
    Customer_Building = temp[len(temp)-1]
else:
    Customer_Building = '<Unknown>'        
 

Customer_Address = '<Unknown>'

Working_Directory = os.getcwd()

resp = 'n'
while resp!='y':
    Arguments_Required = ['Job_Number: {!s}'.format(Job_Number),
                          'Customer_Company: {!s}'.format(Customer_Company),
                          'Customer_Building: {!s}'.format(Customer_Building),
                          'Customer_Address: {!s}'.format(Customer_Address),
                          'Working_Directory: {!s}'.format(Working_Directory)]

    print '\nDetected Info:'
    for i,v in enumerate(Arguments_Required):
        print "{:d}) {!s}".format(i,v)

    print '\nIs This Correct?'
    resp = str(raw_input())

#os.chdir(GUI_dir)
EDSA_Report_Generator.ArcheatTable(Job_Number, Customer_Company, Customer_Building, Customer_Address,
                                    Working_Directory)

exit()
