# EDSA Report Generator

# Imported Modules
import sys
import os
import csv
import xlwt
import datetime as DT
from xlrd import open_workbook
import argparse
import subprocess

global Workbook_FileName

def ArcheatTable(Job_Number, Customer_Company, Customer_Building, Customer_Address, Working_Directory, Report_Directory):

    """
    Job_Number = 'S2913',
    Customer_Company = 'Sifto'
    Customer_Building = 'Sifto Evaporator Plant'
    Customer_Address = 'Goderich'
    Working_Directory = 'd:\Scott\Scripts\Python\Report Generator\Test\'
    """

	

    Logo_Directory = 'z:\Source-Code\ArcFlash\AF-Report-full\Images'

    os.chdir(Working_Directory)
	
    EquipmentList=[]
		            
    Workbook_FileName = '{!s}-AF-Archeat_Tables[{:%Y-%m-%d_%H%M%S}].xls'.format(Job_Number, DT.datetime.now())
	
	# Styles for Excel Report
    Main_Title_Style1 = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid_fill, fore_colour pale_blue; font: height 400, name HandelGothic BT, color-index dark_blue; border: left 2, left_colour black, right 2, right_colour black, top 2, top_colour black, bottom 0, bottom_colour black', num_format_str='#,##0')
    Main_Title_Style2 = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid_fill, fore_colour pale_blue; font: height 320, name HandelGothic BT, color-index dark_blue; border: left 2, left_colour black, right 2, right_colour black, top 0, top_colour black, bottom 0, bottom_colour black', num_format_str='#,##0')
    Main_Title_Style3 = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid_fill, fore_colour pale_blue; font: height 400, name HandelGothic BT, color-index dark_red; border: left 2, left_colour black, right 2, right_colour black, top 0, top_colour black, bottom 2, bottom_colour black', num_format_str='#,##0')
    Main_Title_Style4 = xlwt.easyxf('alignment: horizontal center; font: height 400, name HandelGothic BT; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0')

    Headings_Style = xlwt.easyxf('alignment: rotation +90, horizontal center; pattern: pattern solid_fill, fore_colour gray50; font: height 200, name Arial Black, color-index dark_blue; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0')

    Gap_Style = xlwt.easyxf('alignment: horizontal center, vertical center; pattern: pattern solid, fore_colour white; font: height 10, name Arial, color-index black; border: left 0, right 0, top 0, bottom 0', num_format_str='#,##0.000')

	
	
    TableText_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour white; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')
    TableText_StyleL = xlwt.easyxf('alignment: horizontal left; pattern: pattern solid, fore_colour white; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')

	
	
    ArcheatNM_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour turquoise; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')
    Archeat0_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour green; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')
    Archeat1_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour yellow; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')
    Archeat2_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour orange; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')
    Archeat3_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour pink; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')
    Archeat4_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour red; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')
    ArcheatNA_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour brown; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')

    GeneralNotes_Title_Style = xlwt.easyxf('alignment: horizontal left; pattern: pattern solid, fore_colour white; font: height 360, name Calibri, color-index black;', num_format_str='#,##0')
    Explanations_Title_Style = xlwt.easyxf('alignment: horizontal left; pattern: pattern solid, fore_colour white; font: italic True, height 360, name Calibri, color-index black;', num_format_str='#,##0')
    GeneralNotesL_Style = xlwt.easyxf('alignment: horizontal left; pattern: pattern solid, fore_colour white; font: bold True, height 220, name Calibri, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0')
    GeneralNotesR_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour white; font: bold True, height 220, name Calibri, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0')
    Explanations_Style = xlwt.easyxf('alignment: wrap True, vertical top, horizontal left; pattern: pattern solid, fore_colour white; font: italic True, bold True, height 200, name Calibri, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0')


    #Equipment Class Declaration
    class Equipment:
        'Common Base Class for all Equipment'

        def __init__(self, BusName, ProtectiveDeviceName, BusVoltage, BoltedFaultCurrent,
                     BranchCurrent, CriticalCase, ArcingCurrent, TripDelayTime, FaultDuration,
                     Configuration, ArcFlashBoundary, WorkingDistance, AvailableEnergy, PPEClass):

            self.BusName = BusName
            self.ProtectiveDeviceName = ProtectiveDeviceName
            self.BusVoltage = BusVoltage
            self.BoltedFaultCurrent = BoltedFaultCurrent
            self.BranchCurrent = BranchCurrent
            self.CriticalCase = CriticalCase
            self.ArcingCurrent = ArcingCurrent
            self.TripDelayTime = TripDelayTime
            self.FaultDuration = FaultDuration
            self.Configuration = Configuration
            self.ArcFlashBoundary = ArcFlashBoundary
            self.WorkingDistance = WorkingDistance
            self.AvailableEnergy = AvailableEnergy
            if PPEClass == '0':
                    self.PPEClass = 'NM'
            else:
                    self.PPEClass = PPEClass

            #PDC Missmatch Sanitization
            if self.ProtectiveDeviceName.find('!')!=-1:
                (a, b) = self.ProtectiveDeviceName.split('!')
                self.ProtectiveDeviceName = str(a+b)
            elif self.ProtectiveDeviceName.find('#')!=-1:
                (a, b) = self.ProtectiveDeviceName.split('#')
                self.ProtectiveDeviceName = str(a+b)
            else:
                pass

            #Direct csv file Sanitization
            names = ('BusName',
                   'ProtectiveDeviceName',
                   'BusVoltage',
                   'BoltedFaultCurrent',
                   'BranchCurrent',
                   'CriticalCase',
                   'ArcingCurrent',
                   'TripDelayTime',
                   'FaultDuration',
                   'Configuration',
                   'ArcFlashBoundary',
                   'WorkingDistance',
                   'AvailableEnergy',
                   'PPEClass',)
            
            for n in names:
                v = getattr(self, n)
                v=str(v)
                if (v.find('"')!=-1):
                    (a, info, c) = v.split('"')
                    setattr(self, n, info)
                    z=getattr(self, n) 
                else:
                    pass


            #Set Archeat Colours
            if self.PPEClass=='0' :
                self.Archeat_Style = Archeat0_Style
            elif self.PPEClass=='NM' :
                self.Archeat_Style = ArcheatNM_Style
            elif self.PPEClass=='1' :
                self.Archeat_Style = Archeat1_Style
            elif self.PPEClass=='2' :
                self.Archeat_Style = Archeat2_Style
            elif self.PPEClass=='3' :
                self.Archeat_Style = Archeat3_Style
            elif self.PPEClass=='4' :
                self.Archeat_Style = Archeat4_Style
            elif self.PPEClass=='Danger' :
                self.Archeat_Style = ArcheatNA_Style
            elif self.PPEClass=='N/A' :
                self.Archeat_Style = ArcheatNA_Style
            else:
                self.Archeat_Style = Archeat0_Style

            # Set Calc Factor
            self.CalcFactor = '1.5'

            #Setup Arc Flash Boundaries
            if self.BusVoltage <= '0.050':
                self.LimitedAB = 'Not Specified'
                self.RestrictedAB = 'Not Specified'
            elif (self.BusVoltage > '0.050')&(self.BusVoltage <= '0.150'):
                self.LimitedAB = '42'
                self.RestrictedAB = 'Avoid Contact'
            elif (self.BusVoltage > '0.150')&(self.BusVoltage <= '0.750'):
                self.LimitedAB = '42'
                self.RestrictedAB = '12'
            elif (self.BusVoltage > '0.750')&(self.BusVoltage <= '15.000'):
                self.LimitedAB = '60'
                self.RestrictedAB = '26'
            elif (self.BusVoltage > '15.000')&(self.BusVoltage <= '36.000'):
                self.LimitedAB = '72'
                self.RestrictedAB = '31'
            elif (self.BusVoltage > '36.000')&(self.BusVoltage <= '46.000'):
                self.LimitedAB = '96'
                self.RestrictedAB = '33'
            else:
                self.LimitedAB = 'Equipment_Voltage_Error'
                self.RestrictedAB = 'Equipment_Voltage_Error'
                print ('Voltage Out of Range for {!s}.  Please Update Voltage Range').format(self.BusName)

            #Voltage Sanitize
            #(kV, V) = self.BusVoltage.split('.')
            #self.BusVoltageGroup = ((int(kV)*1000)+(int(V)))
            self.BusVoltageGroup = int(float(self.BusVoltage)*1000)

            #Analysis Date
            self.AnalysisDate = '{:%Y-%m-%d}'.format(DT.datetime.now())

            #PPE Requirements as per Table 5 of CSAZ462
            if self.PPEClass == '0' or self.PPEClass == 'NM':
                self.InsulatedTools = 'RQ'
                self.FaceShield = 'AN'
                self.SuitHood = 'NR'
            elif self.PPEClass == '1':
                self.InsulatedTools = '1'
                self.FaceShield = 'RQ'
                self.SuitHood = 'NR'
            elif self.PPEClass == '2':
                self.InsulatedTools = 'RQ'
                self.FaceShield = 'RQ'
                self.SuitHood = 'NR'
            elif self.PPEClass == '3':
                self.InsulatedTools = 'RQ'
                self.FaceShield = 'NR'
                self.SuitHood = 'RQ'
            elif self.PPEClass == '4':
                self.InsulatedTools = 'RQ'
                self.FaceShield = 'NR'
                self.SuitHood = 'RQ'
            elif self.PPEClass == 'N/A' or self.PPEClass == 'Danger':
                self.InsulatedTools = 'RQ'
                self.FaceShield = 'NR'
                self.SuitHood = 'RQ' 
            else:
                self.InsulatedTools = 'ERROR'
                self.FaceShield = 'ERROR'
                self.SuitHood = 'ERROR'

            #Glove Requirements
            #Setup Arc Flash Boundaries
            if self.BusVoltage <= '0.500':
                self.GloveClass = '00'
            elif (self.BusVoltage > '0.500')&(self.BusVoltage <= '1.000'):
                self.GloveClass = '0'
            elif (self.BusVoltage > '1.000')&(self.BusVoltage <= '7.500'):
                self.GloveClass = '1'
            elif (self.BusVoltage > '7.500')&(self.BusVoltage <= '17.000'):
                self.GloveClass = '2'
            elif (self.BusVoltage > '17.000')&(self.BusVoltage <= '26.500'):
                self.GloveClass = '3'
            elif (self.BusVoltage > '26.500')&(self.BusVoltage <= '36.000'):
                self.GloveClass = '4'
            else:
                self.GloveClass = 'ERROR'

        def __str__(self):
            names = ('BusName',
                   'ProtectiveDeviceName',
                   'BusVoltage',
                   'BoltedFaultCurrent',
                   'BranchCurrent',
                   'CriticalCase',
                   'ArcingCurrent',
                   'TripDelayTime',
                   'FaultDuration',
                   'Configuration',
                   'ArcFlashBoundary',
                   'WorkingDistance',
                   'AvailableEnergy',
                   'PPEClass',
                   'CalcFactor',
                   'LimitedAB',
                   'RestrictedAB',
                   'ProhibitedAB',
                   'BusVoltageGroup',)
            out = []
            for n in names:
                v = getattr(self, n)
                out.append("{name:<30} : {value:>30}\n".format(name=n, value=v))
            out.append('{}\n'.format('*' * 63))

            return ''.join(out)
            
        def DisplayEquipment(self):
            print str(self)
           
        def PrintArcheatTableRow(self, BusIteration):
            ws.write(BusIteration, 0, self.BusName, TableText_StyleL)
            ws.write(BusIteration, 1, self.ProtectiveDeviceName, TableText_StyleL)
            ws.write(BusIteration, 2, self.BusVoltage, TableText_Style)
            ws.write(BusIteration, 3, self.BoltedFaultCurrent, TableText_Style)
            ws.write(BusIteration, 4, self.BranchCurrent, TableText_Style)
            ws.write(BusIteration, 5, self.CriticalCase, TableText_Style)
            ws.write(BusIteration, 6, self.ArcingCurrent, TableText_Style)
            ws.write(BusIteration, 7, self.TripDelayTime, TableText_Style)
            ws.write(BusIteration, 8, self.FaultDuration, TableText_Style)
            ws.write(BusIteration, 9, self.Configuration, TableText_Style)
            ws.write(BusIteration, 10, self.ArcFlashBoundary, TableText_Style)
            ws.write(BusIteration, 11, self.WorkingDistance, TableText_Style)
            ws.write(BusIteration, 12, self.AvailableEnergy, TableText_Style)
            ws.write(BusIteration, 13, self.PPEClass, self.Archeat_Style)

    # Split Data from Headings and organize into Equipment Class
    i=0
	
    
    #try:
    with open('ARCHEAT.csv') as csvfile:
        FileReader = csv.reader(csvfile, delimiter=',', quotechar='|')
        for row in FileReader:
            if i == 0:
                Heading = row
            else:
                if (len(row[0]) > 2):
                    EquipmentList.append(Equipment(row[0], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[10], row[11], row[12], row[13], row[14], row[15]))
                else:
                    pass
                

            i = i+1


    #Remove Unwanted Table Columns
    a=Heading[1]
    b=Heading[9]
    Heading.remove(a)
    Heading.remove(b)

    #Sanitize Headings
    Headings_Pure = []
    for h in Heading:
        if (h.find('"')!=-1):
            (a, info, c) = h.split('"')
            Headings_Pure.append(info)
        else:
            Headings_Pure.append(h)

    index = []
    q = 0
    for h in Headings_Pure:
        if (h == ''):
            index.append(q)
        else:
            pass
        q=q+1

    for h in index:
        Headings_Pure.remove(Headings_Pure[h])
    
    #EquipmentList[0].DisplayEquipment()
    #print(EquipmentList[0].BusVoltage)

    #Generates The List of Voltages in the System
    VoltagesPresent = []
    VoltagesList = []

    """
    for eachclass in EquipmentList:
        eachclass.DisplayEquipment()
    """

    def remove_values_from_list(List, Value):
        while Value in List:
            List.remove(Value)

    def DuplicateSearch(List, Val):
        Double = 0
        for i in List:
            if i==Val:
                Double=1
            else:
                pass
        return(Double)
        
    for eachbus in EquipmentList:
        VoltagesPresent.append(eachbus.BusVoltageGroup)

    for V1 in VoltagesPresent:
        Repeat_flag = DuplicateSearch(VoltagesList, V1)
        if Repeat_flag == 0:
            VoltagesList.append(V1)
        else:
            pass

    #Sort Equipment by PPEClass
    ClassList=['NM', '0', '1', '2', '3', '4', 'Danger', 'N/A']
    Temp1=[]
    Temp2=[]
    concerns=[]
    for eachclass in ClassList:
        for EachItem in EquipmentList:
            if EachItem.PPEClass==eachclass:
                Temp1.append(Equipment(EachItem.BusName, EachItem.ProtectiveDeviceName, EachItem.BusVoltage, EachItem.BoltedFaultCurrent, EachItem.BranchCurrent, EachItem.CriticalCase, EachItem.ArcingCurrent, EachItem.TripDelayTime, EachItem.FaultDuration, EachItem.Configuration, EachItem.ArcFlashBoundary, EachItem.WorkingDistance, EachItem.AvailableEnergy, EachItem.PPEClass))
            else:
                pass

        for eachobject in Temp1:
            Temp2.append(eachobject)
        Temp1=[]

    #Sanitize PPE Classes
    for eachobject in Temp2:
        if eachobject.PPEClass=='N/A':
            eachobject.PPEClass='Danger'
	
	txt_FileName = 'Concerns.txt'
	with open(txt_FileName, 'w') as txt_file:
			#for eachobject in Temp2:
				#eachobject.BusName.replace('_', '-')
				#print("{!s}\n".format(eachobject.BusName))
			#txt_file.write("\begin{enumerate}\n")
                        Classes = ['Danger', 'N/A', '4', '3']
                        for EachClass in Classes:
                            for eachobject in Temp2:
                                    if eachobject.PPEClass==EachClass:
                                            #print '{!s} is Arc Hazard Class {!s}\n'.format(eachobject.BusName, eachobject.PPEClass)
                                            txt_file.write("\item {!s} is Arc Hazard Class {!s}\n".format(eachobject.BusName.replace('_', '-'), eachobject.PPEClass))


   
    


##			for eachobject in Temp2:
##				if eachobject.PPEClass=='Danger':
##					#print '{!s} is Arc Hazard Class {!s}\n'.format(eachobject.BusName, eachobject.PPEClass)
##					txt_file.write("\item {!s} is Arc Hazard Class {!s}\n".format(eachobject.BusName.replace('_', '-'), eachobject.PPEClass))
##				elif eachobject.PPEClass=='N/A':
##					#print '{!s} is Arc Hazard Class {!s}\n'.format(eachobject.BusName, eachobject.PPEClass)
##					txt_file.write('\item {!s} is Arc Hazard Class {!s}\n'.format(eachobject.BusName.replace('_', '-'), eachobject.PPEClass))
##				elif eachobject.PPEClass=='4':
##					#print '{!s} is Arc Hazard Class {!s}\n'.format(eachobject.BusName, eachobject.PPEClass)
##					txt_file.write('\item {!s} is Arc Hazard Class {!s}\n'.format(eachobject.BusName.replace('_', '-'), eachobject.PPEClass))
##				elif eachobject.PPEClass=='3':
##					#print '{!s} is Arc Hazard Class {!s}\n'.format(eachobject.BusName, eachobject.PPEClass)
##					txt_file.write('\item {!s} is Arc Hazard Class {!s}\n'.format(eachobject.BusName.replace('_', '-'), eachobject.PPEClass))
##			#txt_file.write("\end{enumerate}\n")



    # Sort Equipment by Voltages
    Temp=[]
    SortedEquipmentLists=[]
    BusesPerVoltage=[]
    UnsortedCount =0
    print 'Voltages Detected: ', VoltagesList

    for eachvoltage in VoltagesList:
        for EachItem in Temp2:
            if EachItem.BusVoltageGroup==eachvoltage:
                Temp.append(Equipment(EachItem.BusName, EachItem.ProtectiveDeviceName, EachItem.BusVoltage, EachItem.BoltedFaultCurrent, EachItem.BranchCurrent, EachItem.CriticalCase, EachItem.ArcingCurrent, EachItem.TripDelayTime, EachItem.FaultDuration, EachItem.Configuration, EachItem.ArcFlashBoundary, EachItem.WorkingDistance, EachItem.AvailableEnergy, EachItem.PPEClass))
            else:
                pass
        for eachobject in Temp:
            SortedEquipmentLists.append(eachobject)

        UnsortedCount = UnsortedCount + len(Temp)
        print 'Sorted ', len(Temp), '/', len(EquipmentList), '...'
        BusesPerVoltage.append(len(Temp))
        Temp=[]

    UnsortedCount = len(EquipmentList) - UnsortedCount
    print UnsortedCount, ' pieces of Equipment were left unsorted'                                         

    #Label Paths
    Label_Script_Path = "z:\Source-Code\ArcFlash\AF-Report-full\Python\CreateLabel\CreateAFLabel.py"
    print "\n\n", Label_Script_Path, "\n\n"


    #Write to Excel
    wb = xlwt.Workbook()

    for eachvoltage in VoltagesList:

        #Sheet Name
        ws = wb.add_sheet('{!s}V Equipment'.format(eachvoltage))    

        #Header and Footer
        FOOTER = str(u"&L[{:%Y/%m/%d}]" u"&RPowerCore Engineering www.PowerCore.ca".format(DT.datetime.now()))
        HEADER = ' '

        ws.footer_str = (FOOTER)
        ws.header_str = (HEADER)

        #Title Block

        line=0

        ws.write_merge(line, line, 1, 13, ('{!s} - {!s}').format(Customer_Company, Customer_Building), Main_Title_Style1)

        line = line + 1
     
        ws.write_merge(line, line, 1, 13, ('{!s}').format(Customer_Address), Main_Title_Style2)

        line = line + 1

        ws.write_merge(line, line, 1, 13, ('Arc Flash Analysis - {!s}V Equipment').format(eachvoltage), Main_Title_Style3)

        line=line+1

        
        
        for q, eachcol in enumerate(Headings_Pure):
            ws.write(line, q, eachcol, Headings_Style)

        line=line+1

        Label_Dir = os.getcwd()

        #Print Arc Heat Info
        for eachclass in SortedEquipmentLists:
            if eachclass.BusVoltageGroup==eachvoltage:   
                eachclass.PrintArcheatTableRow(line)
                SheetCalcFactor = eachclass.CalcFactor
                SheetLimitedAB = eachclass.LimitedAB
                SheetRestrictedAB = eachclass.RestrictedAB
                line=line+1

                subprocess.call(["python", Label_Script_Path,
                                 "Warning", eachclass.BusName, eachclass.BusVoltage, eachclass.LimitedAB,
                                 eachclass.RestrictedAB, eachclass.GloveClass, eachclass.InsulatedTools, eachclass.AvailableEnergy,
                                 eachclass.ArcFlashBoundary, eachclass.WorkingDistance, eachclass.PPEClass, eachclass.FaceShield,
                                 eachclass.SuitHood, eachclass.AnalysisDate, Report_Directory])
                print "{!s} Label Created".format(eachclass.BusName)
                
                Label_Check = 0
                while(Label_Check == 0):
                    for file in os.listdir(os.getcwd()):
                        if file.find(eachclass.BusName):
                            Label_Check = 1
        #LabelsDirectory = '{!s}-Labels[{:%Y-%m-%d_%H%M%S}].xls'.format(Job_Number, DT.datetime.now())
        #os.rename('Labels', LabelsDirectory)

        #Space Before Notes and General Explanation 
        line=line+1

        #Column Width Adjustments
        ws.col(0).width=256*32
        ws.col(1).width=256*24
        ws.col(2).width=256*13
        ws.col(3).width=256*7
        ws.col(4).width=256*7
        ws.col(5).width=256*7
        ws.col(6).width=256*7
        ws.col(7).width=256*7
        ws.col(8).width=256*7
        ws.col(9).width=256*7
        ws.col(10).width=256*7
        ws.col(11).width=256*8
        ws.col(12).width=256*6

        #LOGO from Templates
        os.chdir(Logo_Directory)
        ws.write_merge(0, 2, 0, 0, ' ', Main_Title_Style4)
        ws.insert_bitmap('logo6.bmp', 0, 0)
     

        #General Notes & Explanation
        ws.write(line, 0, 'General Notes:', GeneralNotes_Title_Style)
        ws.write(line, 4, 'Explanations:', Explanations_Title_Style)

        line=line+1

        #Line + 1
        ws.write_merge(line, line, 0, 1, 'All equipment Voltage', GeneralNotesL_Style)
        ws.write(line, 2, eachvoltage, GeneralNotesR_Style)
        ws.write_merge(line, line+2, 4, 7, 'Arc Flash Boundary', Explanations_Style)
        ws.write_merge(line, line+2, 8, 13, 'Minimum distance from the arc within which a second degree burn could occur if no protective clothing is worn.', Explanations_Style)

        line=line+1


        #Line + 2
        ws.write_merge(line, line, 0, 1, 'IEEE Calculation Factor', GeneralNotesL_Style)
        ws.write(line, 2, SheetCalcFactor, GeneralNotesR_Style)

        line=line+1


        #Line + 3
        ws.write_merge(line, line, 0, 1, 'Limited Approach Distance (inch)', GeneralNotesL_Style)
        ws.write(line, 2, SheetLimitedAB, GeneralNotesR_Style)

        line=line+1

        #Line + 4
        ws.write_merge(line, line, 0, 1, 'Restricted Shock Distance (inch)', GeneralNotesL_Style)
        ws.write(line, 2, SheetRestrictedAB, GeneralNotesR_Style)
        ws.write_merge(line, line+1, 4, 7, 'Working Distance', Explanations_Style)
        ws.write_merge(line, line+1, 8, 13, "Closest distance a worker's body, excluding arms and hands, would be exposed to the arc.", Explanations_Style)

        line=line+1


        #Line + 5
        #ws.write_merge(line, line, 0, 1, 'Prohibited Approach Distance (inch)', GeneralNotesL_Style)
        #ws.write(line, 2, SheetProhibitedAB, GeneralNotesR_Style)

        line=line+1

        #Line + 6
        ws.write_merge(line, line+1, 4, 7, 'Incident Energy', Explanations_Style)
        ws.write_merge(line, line+1, 8, 13, 'Energy released at the specified working distance expressed in cal/cm^2', Explanations_Style)

        line=line+2

        

        #Line + 8
        ws.write_merge(line, line+1, 4, 7, 'Clothing Class', Explanations_Style)
        ws.write_merge(line, line+1, 8, 13, 'Minimum clothing class designed to protect worker from second degree burns', Explanations_Style)
        line=line+5

        #Set Print Area
        ws.paper_size_code = 1
        ws.horz_page_breaks = [(line, 0, 14)]
        ws.vert_page_breaks = [(14, 0, line)]

        #Set Page Witdh to 1 Page
        ws.fit_height_to_pages = 0
        ws.fit_width_to_pages = 1
        ws.fit_num_pages = 1
        
    os.chdir(Working_Directory)


    wb.save(Workbook_FileName)

    print '\n', Workbook_FileName, ' Generated', '\n'

    #except:
        #print "\n\nNo Valide ARCHEAT.csv File Located!"
##
##def main():
##
##    parser = argparse.ArgumentParser(description='Creates Archeat Tables from EDSA Output')
##    parser.add_argument('Job_Num', help='Job Number (ie. S2876)')
##    parser.add_argument('Customer_Comp', help='Customer')
##    parser.add_argument('Customer_Build', help='Location')
##    parser.add_argument('Customer_Add', help='Address')
##    parser.add_argument('Working_Dir', help='working directory')
##    args = parser.parse_args()
##
##    ArcheatTable(args.Job_Num, args.Customer_Comp, args.Customer_Build, args.Customer_Add, args.Working_Dir)
##
##
##    os.chdir(args.Working_Dir)
##
##   
##    print '\n', Workbook_FileName, ' Generated', '\n'
##    
##
if __name__ == '__main__':
    main()
    sys.exit()

