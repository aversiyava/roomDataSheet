import clr
import sys
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)

# Import ToDSType(bool) extension method
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

# Import DocumentManager and TransactionManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Import RevitAPI
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
#The inputs to this node will be stored as a list in the IN variables.
dataEnteringNode = IN

revitLinkInstances = UnwrapElement(IN[0])
view = UnwrapElement(IN[1])
boundingBoxes = []

for container in revitLinkInstances:
	bbxs = container.get_BoundingBox(view).ToProtoType()
	boundingBoxes.append(bbxs)
#Assign your output to the OUT variable.
OUT = boundingBoxes
