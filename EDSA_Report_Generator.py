# EDSA Report Generator

# Imported Modules
import os
import csv
import xlwt
from datetime import datetime

# Home Working Directory
#os.chdir('c:/Projects/SV0002 - EDSA Report Generator/Test Directory')

#Hard Drive Directory
os.chdir('f:\Personal Projects\SV0002 - EDSA Report Generator/Test Directory')

# List Declaration
EquipmentList = []
Equipment208V = []
Equipment240V = []
Equipment600V = []
Equipment480V = []
Equipment4160V = []
EquipmentUNSORTED = []

#Class Declaration
class Equipment:
    'Common Base Class for all Equipment'

    def __init__(self, BusName, ProtectiveDeviceName, BusVoltage, BoltedFaultCurrent, BranchCurrent, CriticalCase, ArcingCurrent, TripDelayTime, FaultDuration, Configuration, ArcFlashBoundary, WorkingDistance, AvailableEnergy, PPEClass):
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
        
    def DisplayEquipment(self):
        print '\nBusName :              ', self.BusName, '\nProtectiveDeviceName : ', self.ProtectiveDeviceName, '\nBusVoltage :           ', self.BusVoltage, '\nBoltedFaultCurrent :   ', self.BoltedFaultCurrent, '\nBranchCurrent :        ', self.BranchCurrent, '\nCriticalCase :         ', self.CriticalCase, '\nArcingCurrent :        ', self.ArcingCurrent, '\nTripDelayTime :        ', self.TripDelayTime, '\nFaultDuration :        ', self.FaultDuration, '\nConfiguration :        ', self.Configuration, '\nArcFlashBoundary :     ', self.ArcFlashBoundary, '\nWorkingDistance :      ', self.WorkingDistance, '\nAvailableEnergy :      ', self.AvailableEnergy, '\nPPEClass :             ', self.PPEClass, '\n****************************************\n'

# Styles for Excel Report
TableText_Style = xlwt.easyxf('pattern: pattern solid, fore_colour white; font: height 200, name Arial, color-index black; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0.00')
Archeat0_Style = xlwt.easyxf('pattern: pattern solid, fore_colour green; font: height 200, name Arial, color-index black; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0.00')
Archeat1_Style = xlwt.easyxf('pattern: pattern solid, fore_colour yellow; font: height 200, name Arial, color-index black; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0.00')
Archeat2_Style = xlwt.easyxf('pattern: pattern solid, fore_colour orange; font: height 200, name Arial, color-index black; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0.00')
Archeat3_Style = xlwt.easyxf('pattern: pattern solid, fore_colour pink; font: height 200, name Arial, color-index black; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0.00')
Archeat4_Style = xlwt.easyxf('pattern: pattern solid, fore_colour red; font: height 200, name Arial, color-index black; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0.00')
ArcheatNA_Style = xlwt.easyxf('pattern: pattern solid, fore_colour brown; font: height 200, name Arial, color-index black; border: left 2, right 2, top 2, bottom 2', num_format_str='#,##0.00')


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



# Sort Equipment by Voltages
for EachItem in EquipmentList:
    if EachItem.BusVoltage=='0.208':
        Equipment208V.append(Equipment(EachItem.BusName, EachItem.ProtectiveDeviceName, EachItem.BusVoltage, EachItem.BoltedFaultCurrent, EachItem.BranchCurrent, EachItem.CriticalCase, EachItem.ArcingCurrent, EachItem.TripDelayTime, EachItem.FaultDuration, EachItem.Configuration, EachItem.ArcFlashBoundary, EachItem.WorkingDistance, EachItem.AvailableEnergy, EachItem.PPEClass))
    elif EachItem.BusVoltage=='0.240':
        Equipment240V.append(Equipment(EachItem.BusName, EachItem.ProtectiveDeviceName, EachItem.BusVoltage, EachItem.BoltedFaultCurrent, EachItem.BranchCurrent, EachItem.CriticalCase, EachItem.ArcingCurrent, EachItem.TripDelayTime, EachItem.FaultDuration, EachItem.Configuration, EachItem.ArcFlashBoundary, EachItem.WorkingDistance, EachItem.AvailableEnergy, EachItem.PPEClass))
    elif EachItem.BusVoltage=='0.480':
        Equipment480V.append(Equipment(EachItem.BusName, EachItem.ProtectiveDeviceName, EachItem.BusVoltage, EachItem.BoltedFaultCurrent, EachItem.BranchCurrent, EachItem.CriticalCase, EachItem.ArcingCurrent, EachItem.TripDelayTime, EachItem.FaultDuration, EachItem.Configuration, EachItem.ArcFlashBoundary, EachItem.WorkingDistance, EachItem.AvailableEnergy, EachItem.PPEClass))
    elif EachItem.BusVoltage=='0.600':
        Equipment600V.append(Equipment(EachItem.BusName, EachItem.ProtectiveDeviceName, EachItem.BusVoltage, EachItem.BoltedFaultCurrent, EachItem.BranchCurrent, EachItem.CriticalCase, EachItem.ArcingCurrent, EachItem.TripDelayTime, EachItem.FaultDuration, EachItem.Configuration, EachItem.ArcFlashBoundary, EachItem.WorkingDistance, EachItem.AvailableEnergy, EachItem.PPEClass))
    elif EachItem.BusVoltage=='4.160':
        Equipment4160V.append(Equipment(EachItem.BusName, EachItem.ProtectiveDeviceName, EachItem.BusVoltage, EachItem.BoltedFaultCurrent, EachItem.BranchCurrent, EachItem.CriticalCase, EachItem.ArcingCurrent, EachItem.TripDelayTime, EachItem.FaultDuration, EachItem.Configuration, EachItem.ArcFlashBoundary, EachItem.WorkingDistance, EachItem.AvailableEnergy, EachItem.PPEClass))
    else:
        EquipmentUNSORTED.append(Equipment(EachItem.BusName, EachItem.ProtectiveDeviceName, EachItem.BusVoltage, EachItem.BoltedFaultCurrent, EachItem.BranchCurrent, EachItem.CriticalCase, EachItem.ArcingCurrent, EachItem.TripDelayTime, EachItem.FaultDuration, EachItem.Configuration, EachItem.ArcFlashBoundary, EachItem.WorkingDistance, EachItem.AvailableEnergy, EachItem.PPEClass))


#Export Sorting Results
for EachItem in Equipment208V:
    EachItem.DisplayEquipment()

for EachItem in Equipment240V:
    EachItem.DisplayEquipment()

for EachItem in Equipment480V:
    EachItem.DisplayEquipment()

for EachItem in Equipment600V:
    EachItem.DisplayEquipment()

for EachItem in Equipment4160V:
    EachItem.DisplayEquipment()

for EachItem in EquipmentUNSORTED:
    EachItem.DisplayEquipment()

