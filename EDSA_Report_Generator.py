# EDSA Report Generator

# Imported Modules
import os
import csv
import xlwt
import datetime as DT
from xlrd import open_workbook

#Variable List
Job_Number = 'S1234'
Customer_Company = 'PowerCore'
Customer_Building = 'Main Office'
Customer_Address = '4096 Meadowbrook Drive'

Working_Directory = 'e:\Personal Projects\SV0002 - EDSA Report Generator/Test Directory'
#Working_Directory = 'c:/Projects/SV0002 - EDSA Report Generator/Test Directory'
os.chdir(Working_Directory)

EquipmentList=[]
        
# Styles for Excel Report
Main_Title_Style1 = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid_fill, fore_colour pale_blue; font: height 400, name HandelGothic BT, color-index dark_blue; border: left 2, left_colour black, right 2, right_colour black, top 2, top_colour black, bottom 0, bottom_colour black', num_format_str='#,##0')
Main_Title_Style2 = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid_fill, fore_colour pale_blue; font: height 320, name HandelGothic BT, color-index dark_blue; border: left 2, left_colour black, right 2, right_colour black, top 0, top_colour black, bottom 0, bottom_colour black', num_format_str='#,##0')
Main_Title_Style3 = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid_fill, fore_colour pale_blue; font: height 400, name HandelGothic BT, color-index dark_red; border: left 2, left_colour black, right 2, right_colour black, top 0, top_colour black, bottom 2, bottom_colour black', num_format_str='#,##0')

Headings_Style = xlwt.easyxf('alignment: rotation +90, horizontal center; pattern: pattern solid_fill, fore_colour gray50; font: height 200, name Arial Black, color-index dark_blue; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0')

Gap_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour white; font: height 10, name Arial, color-index black; border: left 0, right 0, top 0, bottom 0', num_format_str='#,##0.000')

TableText_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour white; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.000')
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
Explanations_Style = xlwt.easyxf('alignment: wrap True, vertical top, horizontal left; pattern: pattern solid, fore_colour white; font: italic True, bold True, height 220, name Calibri, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0')


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
        self.PPEClass = PPEClass
        if self.PPEClass=='0' :
            self.Archeat_Style = Archeat0_Style
        elif self.PPEClass=='1' :
            self.Archeat_Style = Archeat1_Style
        elif self.PPEClass=='2' :
            self.Archeat_Style = Archeat2_Style
        elif self.PPEClass=='3' :
            self.Archeat_Style = Archeat3_Style
        elif self.PPEClass=='4' :
            self.Archeat_Style = Archeat4_Style
        elif self.PPEClass=='NA' :
            self.Archeat_Style = ArcheatNA_Style
        else:
            self.Archeat_Style = Archeat0_Style           
        self.CalcFactor = '1.5'
        if self.BusVoltage < '0.050':
            self.LimitedAB = 'Not Specified'
            self.RestrictedAB = 'Not Specified'
            self.ProhibitedAB = 'Not Specified'
        elif (self.BusVoltage >= '0.050')&(self.BusVoltage <= '0.300'):
            self.LimitedAB = '42'
            self.RestrictedAB = 'Avoid Contact'
            self.ProhibitedAB = 'Avoid Contact'
        elif (self.BusVoltage >= '0.301')&(self.BusVoltage <= '0.750'):
            self.LimitedAB = '42'
            self.RestrictedAB = '1'
            self.ProhibitedAB = '1'
        elif (self.BusVoltage >= '0.751')&(self.BusVoltage <= '15.000'):
            self.LimitedAB = '60'
            self.RestrictedAB = '26'
            self.ProhibitedAB = '7'
        elif (self.BusVoltage >= '15.100')&(self.BusVoltage <= '36.000'):
            self.LimitedAB = '72'
            self.RestrictedAB = '31'
            self.ProhibitedAB = '10'
        elif (self.BusVoltage >= '36.100')&(self.BusVoltage <= '46.000'):
            self.LimitedAB = '96'
            self.RestrictedAB = '33'
            self.ProhibitedAB = '17'
        else:
            self.LimitedAB = 'Equipment_Voltage_Error'
            self.RestrictedAB = 'Equipment_Voltage_Error'
            self.ProhibitedAB = 'Equipment_Voltage_Error'
            print ('Voltage Out of Range for {!s}.  Please Update Voltage Range').format(self.BusName)
        (kV, V) = self.BusVoltage.split('.')
        self.BusVoltageGroup = ((int(kV)*1000)+(int(V)))

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
        ws.write(BusIteration, 0, self.BusName, TableText_Style)
        ws.write(BusIteration, 1, self.ProtectiveDeviceName, TableText_Style)
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

with open('ARCHEAT.csv') as csvfile:
    FileReader = csv.reader(csvfile, delimiter=',', quotechar='|')
    for row in FileReader:
        if i == 0:
            Heading = row
        else:
            EquipmentList.append(Equipment(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13]))
        i = i+1

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

print 'Voltages Detected: ', VoltagesList


# Sort Equipment by Voltages
Temp=[]
SortedEquipmentLists=[]
BusesPerVoltage=[]
UnsortedCount =0

for eachvoltage in VoltagesList:
    for EachItem in EquipmentList:
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
                                          

#Write to Excel
wb = xlwt.Workbook()

for eachvoltage in VoltagesList:
    ws = wb.add_sheet('{!s}V Equipment'.format(eachvoltage))

    line=0

    ws.write_merge(line, line, 1, 13, ('{!s} - {!s}').format(Customer_Company, Customer_Building), Main_Title_Style1)

    line = line + 1
 
    ws.write_merge(line, line, 1, 13, ('{!s}').format(Customer_Address), Main_Title_Style2)

    line = line + 1

    ws.write_merge(line, line, 1, 13, ('Arc Flash Analysis - {!s}V Equipment').format(eachvoltage), Main_Title_Style3)

    line=line+1

    q=0
    for eachcol in Heading:
        ws.write(line, q, eachcol, Headings_Style)
        q = q + 1

    line=line+1

    for eachclass in SortedEquipmentLists:
        if eachclass.BusVoltageGroup==eachvoltage:   
            eachclass.PrintArcheatTableRow(line)
            SheetCalcFactor = eachclass.CalcFactor
            SheetLimitedAB = eachclass.LimitedAB
            SheetRestrictedAB = eachclass.RestrictedAB
            SheetProhibitedAB = eachclass.ProhibitedAB
            line=line+1

    #Space Before Notes and General Explanation
    line=line+1

    #Column Width Adjustments
    ws.col(0).width=256*24
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

    #General Notes & Explanation
    #Line + 0
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
    ws.write_merge(line, line, 0, 1, 'Prohibited Approach Distance (inch)', GeneralNotesL_Style)
    ws.write(line, 2, SheetProhibitedAB, GeneralNotesR_Style)

    line=line+1

    #Line + 6
    ws.write_merge(line, line+1, 4, 7, 'Incident Energy', Explanations_Style)
    ws.write_merge(line, line+1, 8, 13, 'Energy released at the specified working distance expressed in cal/cm^2', Explanations_Style)
    line=line+2

    #Line + 8
    ws.write_merge(line, line+1, 4, 7, 'Clothing Class', Explanations_Style)
    ws.write_merge(line, line+1, 8, 13, 'Minimum clothing class designed to protect worker from second degree burns', Explanations_Style)
    line=line+1


Workbook_FileName = '{!s}-AF_Archeat_Tables[{:%Y-%m-%d_%H%M%S}].xls'.format(Job_Number, DT.datetime.now())
wb.save(Workbook_FileName)

print '\n', Workbook_FileName, ' Generated', '\n'
