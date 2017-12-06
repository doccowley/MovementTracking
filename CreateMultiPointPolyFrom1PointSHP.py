# ---------------------------------------------------------------------------
title = "CreatePolyFrom1PointSHP.py"
msg   = "Copyright (c) 2009 - Andrew Cowley, University of Exeter \n\n"
# ---------------------------------------------------------------------------
# Created on: Thu Apr 30 2009 14:42
# Description: 
# Create Point ShapeFile from Excel.
msg += "This script will create a Polyline ShapeFile for an edge or path "
msg += "from a point shapefile. You need to supply input and output"
msg += "shapefile names and locations."
# ---------------------------------------------------------------------------

# Import system modules
import os, math, arcgisscripting #win32com.client

from easygui import *
msg += "\n\nDo you want to continue?"
if ccbox(msg, title):     # show a Continue/Cancel dialog
    pass  # user chose Continue
else:  # user chose Cancel
    raise SystemExit

myPointDefault = 'X:\Geography\LNaylor\Glamorgan\GIS\L24\*.shp'
myShapeFile1 = fileopenbox(msg="Select 1st Point Shapefile", title=None, default=myPointDefault, filetypes=["*.shp"])
if myShapeFile1:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit
#myShapeFile1 = "X:\Geography\LNaylor\Glamorgan\GIS\L24\XLS2SHP_L24_09Mar08.shp"
(dirName1, fileName1) = os.path.split(myShapeFile1)
#(fileBaseName1, fileExtension1)=os.path.splitext(fileName1)

#dirName = dirName1
dirName = 'X:\Geography\LNaylor\Glamorgan\GIS\L24\Test'
#defaultFile = dirName + '\PNT2PLY_' + fileBaseName + '.shp'
defaultFile = dirName + '\PNT2PLY_TEST.shp'
myOutputFile = filesavebox(msg=None, title=None, default=defaultFile, filetypes=["*.shp"])
if myOutputFile:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit
print myOutputFile

# Create the Geoprocessor object
gp = arcgisscripting.create()
#gp = win32com.client.Dispatch('esriGeoprocessing.GPDispatch')

# Set the Geoprocessing environment...
CurrentWorkingDirectory = os.getcwd()
print "cwd    = " + CurrentWorkingDirectory
print "myShapeFile1 = " + myShapeFile1
print "myOutputFile = " + myOutputFile

# Read 1st input shapefile
print "\nCreating 1st array"
myBlocks1 = {}
i = 0
rows1 = gp.SearchCursor(myShapeFile1)
row1 = rows1.Next()
while row1:
    pnt1 = row1.shape.GetPart()
    xx = [pnt1.x, pnt1.y, pnt1.z]
    #myBlocks1[xx] = row1.Block_No
    myBlocks1[i] = xx
    #print "\t", xx, "(", pnt1.x, pnt1.y, pnt1.z, ")"
    row1 = rows1.Next()
    i += 1
# Delete the row and cursor
del row1, rows1
#print myBlocks1

# Create a spatial reference object
sr = gp.CreateObject("spatialreference")

# Use a projection file to define the spatial reference's properties
sr.CreateFromFile(r'C:\Program Files\ArcGIS\Coordinate Systems\Projected Coordinate Systems\National Grids\British National Grid.prj')

# Create the output feature class using the spatial reference object
gp.CreateFeatureClass(os.path.dirname(myOutputFile),os.path.basename(myOutputFile), "POLYLINE","","","ENABLED", sr)

# Open an insert cursor for the new feature class
cur = gp.InsertCursor(myOutputFile)

# Create array to hold polyline's point data
lineArray = gp.createobject("Array")

rows1 = gp.SearchCursor(myShapeFile1)
row1 = rows1.Next()
while row1:
    #print str(round(float(myBlock),1))
#    print "myBlocks1 key => " + str(myBlock1)

#   print "\t", fileName1, myBlocks1[myBlock1][0], myBlocks1[myBlock1][1], myBlocks1[myBlock1][2]

#    # Create start point of line
#    pnt1 = gp.CreateObject("Point")
#    pnt1.x = myBlocks1[myBlock1][0]
#    pnt1.y = myBlocks1[myBlock1][1]
#    pnt1.z = myBlocks1[myBlock1][2]

#    pnt1 = row1.shape.GetPart()
#    xx = [pnt1.x, pnt1.y, pnt1.z]

    # Add point to lineArray
    lineArray.add(row1.shape.GetPart())
#    del pnt1

    row1 = rows1.Next()

    
row = cur.newrow()
row.shape = lineArray
cur.InsertRow(row)

del lineArray, pnt1, row, cur

print "\nFINISHED :-)"
msgbox("Finished!", title, ok_button="OK")

