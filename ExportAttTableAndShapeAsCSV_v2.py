# ---------------------------------------------------------------------------
title = "ExportAttTableAndShapeAsCSV.py"
msg   = "Copyright (c) 2009 - Andrew Cowley, University of Exeter \n\n"
# ---------------------------------------------------------------------------
# Created on: Fri Aug 07 2009 15:46
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

#InputSHP1 = fileopenbox(msg=None, title=None, default='X:\Geography\LNaylor\Glamorgan\GIS\L24\Blocks\*.shp', filetypes=["*.shp"])
InputSHP1 = fileopenbox(msg=None, title=None, default='W:\Scripts\PythonScripts\Larissa\*.shp', filetypes=["*.shp"])
if InputSHP1:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit

(dirName, fileName) = os.path.split(InputSHP1)
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
print "InputSHP1 = " + InputSHP1

#Get field names from shapefile and remove Geometry fields (e.g. Shape)
myFields = []
fields = gp.ListFields(InputSHP1)
field = fields.Next()
print "\nField\tType\tScale\tPrecision"
while field:
    if field.type <> 'Geometry':
        myFields.append(field.name)
    print field.Name + "\t" + field.Type + "\t" + str(field.Scale) + "\t" + str(field.Precision)
    field = fields.Next()
print "\n" + ",".join(myFields)

rows = gp.SearchCursor(InputSHP1)
row = rows.Next()

fileObj = open(myOutputFile,"w") # open file for write

line = ",".join(myFields) + ",PartNum,X,Y,Z,M\n"
print line
fileObj.write(line)

desc = gp.Describe(InputSHP1)
shapefieldname = desc.ShapeFieldName

rowcount = 0

while row:
    myRow =[]
    for field in myFields:
        myRow.append(str(row.GetValue(field)))
    print ",".join(myRow)

    feat = row.GetValue(shapefieldname)
    # Print the current multipoint's ID
    #
    partnum = 0
    # Count the number of points in the current multipart feature
    #
    partcount = feat.PartCount
    print "Feature " + str(row.getvalue(desc.OIDFieldName)) + ": Partcount = " + str(partcount) + ":"
    # Enter while loop for each point in the multipoint feature
    #
    while partnum < partcount:
        # Get the point based on the current part number
        #
        pnt = feat.GetPart(partnum)
        # Print x,y coordinates of current point
        #
        print pnt.x, pnt.y
        partnum += 1    
#    part = feat.GetPart(0)
#    pnt = gp.createobject("Point")
#    pnt = part.next()
#    pnt_count = 0

#    while pnt:
        data = [str(partnum), str(pnt.x), str(pnt.y), str(pnt.z), str(pnt.m)]
        line = ",".join(myRow) + "," + ",".join(data)  + "\n"
        print line
        fileObj.write(line)

#        pnt = part.next()
#        pnt_count += 1

#    print "\tPoint Count = " + str(pnt_count)

    rowcount += 1
    row = rows.Next()

fileObj.close()

del gp, row, rows, field, fields, myFields

print "\nFINISHED :-)"
msgbox("Finished!", title, ok_button="OK")
