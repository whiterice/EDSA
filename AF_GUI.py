import EDSA_Report_Generator
import os

GUI_dir=os.getcwd()

print "****************************\n*Arc Flash Report Generator*\n****************************"
print "Instructions:\nOnly Archeat Tables Available. Arc Flash Ressults must be saved under 'ARCHEAT.csv'"

cont = 'y'
while (cont != 'quit'):
    try:
        print "Please Specify where you Archeat Data is:"
        Working_Directory = str(raw_input())

        print "Job Number:"
        Job_Number = str(raw_input())

        print "Company Owning Site Under Investigation:"
        Customer_Company = str(raw_input())

        print "Name Of Site Under Investigation:"
        Customer_Building = str(raw_input())

        print "Address Of Site Under Investigation:"
        Customer_Address = str(raw_input())

        print '\n'

        """Runs the Tap Changer Simulation Through TapChanger.py"""
        ArcheatTable(Job_Number, Customer_Company, Customer_Building, Customer_Address,
                     Working_Directory)

    except:
        print "\nInavlid Input\n"
    
    print "\nWant to Run another Scneario?(y/n)"
    cont = raw_input()

    print "\n"

exit()
