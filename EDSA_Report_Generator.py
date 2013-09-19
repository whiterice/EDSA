# EDSA Report Generator

# Imported Modules
import os
import csv
import xlwt
import datetime as DT

#Variable List
Job_Number = 'S1234'
Cutomer_Company = 'PowerCore'
Customer_Buidling = 'Main Office'
Customer_Address = '4096 Meadowbrook Drive'
Working_Directory = 'f:\Personal Projects\SV0002 - EDSA Report Generator/Test Directory'
#Working_Directory = 'c:/Projects/SV0002 - EDSA Report Generator/Test Directory'
os.chdir(Working_Directory)

        
# Styles for Excel Report
TableText_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour white; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.00')
Archeat0_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour green; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.00')
Archeat1_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour yellow; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.00')
Archeat2_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour orange; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.00')
Archeat3_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour pink; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.00')
Archeat4_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour red; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.00')
ArcheatNA_Style = xlwt.easyxf('alignment: horizontal center; pattern: pattern solid, fore_colour brown; font: height 200, name Arial, color-index black; border: left 1, right 1, top 1, bottom 1', num_format_str='#,##0.00')

# List Declaration
EquipmentList = []
Equipment208V = []
Equipment240V = []
Equipment600V = []
Equipment480V = []
Equipment4160V = []
EquipmentUNSORTED = []

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
        self.CalcFactor = 1.5
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
               'ProhibitedAB',)
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
        ws.write(BusIteration, 8, self.Configuration, TableText_Style)
        ws.write(BusIteration, 9, self.ArcFlashBoundary, TableText_Style)
        ws.write(BusIteration, 10, self.WorkingDistance, TableText_Style)
        ws.write(BusIteration, 11, self.AvailableEnergy, TableText_Style)
        ws.write(BusIteration, 12, self.PPEClass, self.Archeat_Style)

    def Sanitize(self):
        (kV, V) = self.BusVoltage.split('.')
        self.BusVoltage = ((int(kV)*1000)+(int(V)))

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

for eachclass in EquipmentList:
    eachclass.Sanitize()
"""
for eachclass in EquipmentList:
    eachclass.DisplayEquipment()
"""

def remove_values_from_list(List, Value):
    while Value in List:
        List.remove(Value)

for eachbus in EquipmentList:
    VoltagesPresent.append(eachbus.BusVoltage)

for eachvoltage in VoltagesPresent:
    VoltagesList.append(eachvoltage)
    remove_values_from_list(VoltagesPresent, eachvoltage)

# Sort Equipment by Voltages
Temp=[]
SortedEquipmentLists=[]
BusesPerVoltage=[]
UnsortedCount =0

for eachvoltage in VoltagesList:
    for EachItem in EquipmentList:
        if EachItem.BusVoltage==eachvoltage:
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

"""
for w in SortedEquipmentLists:
    w.DisplayEquipment()
"""                                          
"""
#Write to Excel
wb = xlwt.Workbook()

for eachlist in SortedEqupmentLists:
    Voltage = eachlist[0].BusVoltage
    ws = wb.add_sheet('{!s}kV Equipment'.format(Voltage))

    line=0

    for eachclass in eachlist:
        eachclass.PrintArcheatTableRow(line)
        line=line+1

    #Space Before Notes and General Explanation
    line=line+1

    
    ws.col(0).width=256*24
    ws.col(1).width=256*24
    ws.col(2).width=256*7
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


Workbook_FileName = '{!s}-AF_Archeat_Tables[{:%Y-%m-%d_%H%M%S}].xls'.format(Job_Number, DT.datetime.now())
wb.save(Workbook_FileName)

print '\n', Workbook_FileName, ' Generated', '\n'
"""
