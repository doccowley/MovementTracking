# ---------------------------------------------------------------------------
title = "ExportAttTableAsCSV.py"
msg   = "Copyright (c) 2009 - Andrew Cowley, University of Exeter \n\n"
# ---------------------------------------------------------------------------
# Created on: Wed Apr 29 2009 15:46
# Description: 
# Export Attribute Table from ahapefile to csv file.
msg += "This script will export the 'Attribute Table' from a shapefile "
msg += "to a csv file. You need to supply and input shapefile and an "
msg += "output .CSV file name/location."
# ---------------------------------------------------------------------------

# Import system modules
import os, math, arcgisscripting

from easygui import *
msg += "\n\nDo you want to continue?"
if ccbox(msg, title):     # show a Continue/Cancel dialog
    pass  # user chose Continue
else:  # user chose Cancel
    raise SystemExit

myShapeFile = fileopenbox(msg=None, title=None, default='X:\Geography\LNaylor\Glamorgan\GIS\L24\Blocks\*.shp', filetypes=["*.shp"])
if myShapeFile:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit

(dirName, fileName) = os.path.split(myShapeFile)
(fileBaseName, fileExtension)=os.path.splitext(fileName)
defaultFile = dirName + '/' + fileBaseName + '.csv'

myOutputFile = filesavebox(msg=None, title=None, default=defaultFile, filetypes=["*.csv"])
if myOutputFile:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit

print myOutputFile

# Create the Geoprocessor object
gp = arcgisscripting.create()

# Set the Geoprocessing environment...
CurrentWorkingDirectory = os.getcwd()
print "myShapeFile = " + myShapeFile

#Get field names from shapefile and remove Geometry fields (e.g. Shape)
myFields = []
fields = gp.ListFields(myShapeFile)
field = fields.Next()
print "\nField\tType\tScale\tPrecision"
while field:
    if field.type <> 'Geometry':
        myFields.append(field.name)
    print field.Name + "\t" + field.Type + "\t" + str(field.Scale) + "\t" + str(field.Precision)
    field = fields.Next()
print "\n" + ",".join(myFields)

rows = gp.SearchCursor(myShapeFile)
row = rows.Next()

fileObj = open(myOutputFile,"w") # open file for write

fileObj.write(",".join(myFields) + "\n")

while row:
    myRow =[]
    for field in myFields:
        myRow.append(str(row.GetValue(field)))
    print ",".join(myRow)
    fileObj.write(",".join(myRow)  + "\n")

    row = rows.Next()

fileObj.close()

del gp, row, rows, field, fields, myFields

print "\nFINISHED :-)"
msgbox("Finished!", title, ok_button="OK")
