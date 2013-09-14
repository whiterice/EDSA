# EDSA Report Generator

import os
import csv

# Home Working Directory
#os.chdir('c:/Projects/SV0002 - EDSA Report Generator/Test Directory')

# PowerCore Working Directory
# os.chdir('c:\Documents and Settings\Scott Vermeire\My Documents\Dropbox\EDSA Report Generator\Test Directory')"""

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
    ECount=0

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
        Equipment.ECount += 1
        
    def DisplayEquipment(self):
        print("\nBusName :              ", self.BusName, "\nProtectiveDeviceName : ", self.ProtectiveDeviceName, "\nBusVoltage :           ", self.BusVoltage, "\nBoltedFaultCurrent :   ", self.BoltedFaultCurrent, "\nBranchCurrent :        ", self.BranchCurrent, "\nCriticalCase :         ", self.CriticalCase, "\nArcingCurrent :        ", self.ArcingCurrent, "\nTripDelayTime :        ", self.TripDelayTime, "\nFaultDuration :        ", self.FaultDuration, "\nConfiguration :        ", self.Configuration, "\nArcFlashBoundary :     ", self.ArcFlashBoundary, "\nWorkingDistance :      ", self.WorkingDistance, "\nAvailableEnergy :      ", self.AvailableEnergy, "\nPPEClass :             ", self.PPEClass, "\n****************************************\n")

    def DisplayEquipmentCount():
        print("\nTotal Amount of Equipment Imported : ", Equipment.ECount, "\n****************************************\n")

# Split Data from Headings and organize into Equipment Class
i=0

with open('ARCHEAT.csv', newline='') as csvfile:
    FileReader = csv.reader(csvfile, delimiter=',', quotechar='|')
    for row in FileReader:
        if i == 0:
            Heading = row
        else:
            EquipmentList.append(Equipment(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13]))
        i = i+1

#EquipmentList[0].DisplayEquipment()
Equipment.DisplayEquipmentCount()
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


