import clr
import sys
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("System")
from System.Collections.Generic import List

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

from itertools import groupby

doc = DocumentManager.Instance.CurrentDBDocument

#VIEW FOR GETTING BOUNDINGBOX
allViews = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
view = None
for i in range(len(allViews)):
    if allViews[i].Name == "Containers":
        view = allViews[i]
        break

def roomNumber(room):
    return room.Number
allRooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
rooms = sorted(allRooms, key = roomNumber)

def midPointBox(maxPoint, minPoint):
    pointMid1 = maxPoint.Subtract(minPoint)
    pointMid2 = pointMid1.Divide(2)
    pointMid = maxPoint.Subtract(pointMid2)
    return pointMid

def stringForNamesList(listToSplit):
    ordering = sorted(listToSplit)
    elementsGroup = []
    for key, group in groupby(ordering):
        elementsGroup.append(list(group))
    stringList = []
    for i in range(len(elementsGroup)):
        numberElements = str(len(elementsGroup[i])) + " x " + elementsGroup[i][0]
        stringList.append(numberElements)
    completeString = " / ".join(stringList)
    return completeString

linkInstances = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()

ARCHcategories = [
    BuiltInCategory.OST_Casework,
    BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_PlumbingFixtures
]
ARCHRDSparameters = [
    "RDS_5_Built in furniture",
    "RDS_5_Loose furniture",
    "RDS_5_Custom furniture"
]

MEPbuiltInCategories = [
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_FireAlarmDevices,
    BuiltInCategory.OST_LightingDevices,
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_SecurityDevices,
    BuiltInCategory.OST_Sprinklers,
    BuiltInCategory.OST_TelephoneDevices
]
MEPRDSparameters = [
    "RDS_9_Communication device",
    "RDS_9_Data device",
    "RDS_9_Duct fitting",
    "RDS_9_Electrical equipment",
    "RDS_9_Electrical fixture",
    "RDS_9_Fire alarm device",
    "RDS_9_Lighting device",
    "RDS_9_Lighting fixture",
    "RDS_9_Mechanical equipment",
    "RDS_9_Pipe accessory",
    "RDS_9_Pipe fitting",
    "RDS_9_Plumbing fixture",
    "RDS_9_Security device",
    "RDS_9_Sprinkler",
    "RDS_9_Telephone device"
]

test = []

TransactionManager.Instance.EnsureInTransaction(doc)

for room in rooms:
    ceilingHeights = []
    ceilingHeightFeets = []
#
    for linkInstance in linkInstances:
        linkDoc = linkInstance.GetLinkDocument()
        roomLevel = room.Level.Name
        roomLocation = room.Location.Point
        roomContainer = room.LookupParameter("AR_ContainerTest").AsString()
        linkName = linkInstance.Name
        if roomContainer in linkName:
            try:
                allWalls = FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
                for wall in allWalls:
                    wallName = wall.Name
                    if "LN." not in wallName and "PT." not in wallName:
                        try:
                            room.LookupParameter("RDS_2_Wall type").Set(wallName)
                        except:
                            pass
            except:
                pass
        #RDS1
        for i in range(len(ARCHcategories)):
            try:
                elementsByCategory = FilteredElementCollector(linkDoc).OfCategory(ARCHcategories[i]) \
                                                                      .WhereElementIsNotElementType() \
                                                                      .ToElements()
                theseElements = []
                for element in elementsByCategory:
                    transform = linkInstance.GetTotalTransform()
                    transformedPoint = transform.OfPoint(element.Location.Point)
                    elementsInRoom = []
                    if room.IsPointInRoom(transformedPoint):
                        elementName = element.Name
                        if ARCHcategories[i] == BuiltInCategory.OST_Casework and elementName.startswith("J") \
                        or ARCHcategories[i] == BuiltInCategory.OST_PlumbingFixtures and elementName.startswith("SAN") \
                        or ARCHcategories[i] == BuiltInCategory.OST_Furniture:
                            theseElements.append(elementName)
                            finalString = stringForNamesList(theseElements)
                            try:
                                room.LookupParameter(ARCHRDSparameters[i]).Set(finalString)
                            except:
                                pass
            except:
                pass
    #RDS1
    #RDS_1_Capacity, RDS_1_Occupants, RDS_1_Occupancy FT
    occupancy = room.LookupParameter("Occupancy").AsString()
    room.LookupParameter("RDS_1_Capacity").Set(occupancy)
    room.LookupParameter("RDS_1_Occupants").Set(occupancy)
    room.LookupParameter("RDS_1_Occupancy FT").Set(occupancy)
    #RDS_2_Windows, RDS_2_Blinds, RDS_2_Equipment
    room.LookupParameter("RDS_2_Windows").Set("N/A")
    room.LookupParameter("RDS_2_Blinds").Set("N/A")
    room.LookupParameter("RDS_2_Equipment").Set("N/A")
    #RDS_6_Accessories, RDS_6_Art Program
    room.LookupParameter("RDS_6_Accessories").Set("N/A")
    room.LookupParameter("RDS_6_Art Program").Set("N/A")
    #RDS_7_Special requirements
    room.LookupParameter("RDS_7_Special requirements").Set("N/A")
    #RDS_1_Department
    functionality = room.LookupParameter("Department").AsString()
    room.LookupParameter("RDS_1_Department").Set(functionality)
    #RDS_1_Adjacencies
    i = rooms.index(room)
    currentRoomLevel = room.Level.Name
    try:
        nextRoomLevel = rooms[i+1].Level.Name
        nameNextRoom = rooms[i+1].LookupParameter("Name").AsString()
        if i == 0:
            nameNextNextRoom = rooms[i+2].LookupParameter("Name").AsString()
            stringBoth = nameNextRoom + " / " + nameNextNextRoom
            room.LookupParameter("RDS_1_Adjacencies").Set(stringBoth)
        elif currentRoomLevel == nextRoomLevel:
            namePreviousRoom = rooms[i-1].LookupParameter("Name").AsString()
            stringBoth = nameNextRoom + " / " + namePreviousRoom
            room.LookupParameter("RDS_1_Adjacencies").Set(stringBoth)
        else:
            namePreviousRoom = rooms[i-1].LookupParameter("Name").AsString()
            namePreviousPreviousRoom = rooms[i-2].LookupParameter("Name").AsString()
            stringBoth = namePreviousRoom + " / " + namePreviousPreviousRoom
            room.LookupParameter("RDS_1_Adjacencies").Set(stringBoth)
    except:
        namePreviousRoom = rooms[i-1].LookupParameter("Name").AsString()
        namePreviousPreviousRoom = rooms[i-2].LookupParameter("Name").AsString()
        stringBoth = namePreviousRoom + " / " + namePreviousPreviousRoom
        room.LookupParameter("RDS_1_Adjacencies").Set(stringBoth)
        #RDS_2_Doors
        try:
            elementsByCategory = FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Doors) \
                                                                  .WhereElementIsNotElementType() \
                                                                  .ToElements()
            theseDoors = []
            hardwareDoors = []
            accessDoors = []
            for door in elementsByCategory:
                transform = linkInstance.GetTotalTransform()
                maxPoint = transform.OfPoint(door.get_BoundingBox(view).Max)
                minPoint = transform.OfPoint(door.get_BoundingBox(view).Min)
                midPoint = midPointBox(maxPoint, minPoint)
                if room.IsPointInRoom(midPoint):
                    theseDoors.append(door.Name)
                    finalString = stringForNamesList(theseDoors)
                    try:
                        #LIST DOORS IN ROOM
                        room.LookupParameter("RDS_2_Doors").Set(finalString)
                        #HARDWARE
                        hardware = door.LookupParameter("AR-Doors-Hardware").AsString()
                        if hardware:
                            hardwareDoors.append(hardware)
                            hardwareString = stringForNamesList(hardwareDoors)
                            room.LookupParameter("RDS_2_Doors hardware").Set(hardwareString)
                        #ACCESS CONTROL
                        accessControl = door.LookupParameter("AR-Doors-Door entry").AsString()
                        if accessControl:
                            accessDoors.append(accessControl)
                            accessString = stringForNamesList(accessDoors)
                            room.LookupParameter("RDS_2_Doors access").Set(accessString)
                    except:
                        pass
        except:
            pass

TransactionManager.Instance.TransactionTaskDone()

OUT = test
