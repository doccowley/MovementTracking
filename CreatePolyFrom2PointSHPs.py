# ---------------------------------------------------------------------------
title = "CreatePolyFrom2PointSHPs.py"
msg   = "Copyright (c) 2009 - Andrew Cowley, University of Exeter \n\n"
# ---------------------------------------------------------------------------
# Created on: Thu Apr 30 2009 14:42
# Description: 
# Create Point ShapeFile from Excel.
msg += "This script will create a Polyline ShapeFile for start and "
msg += "finish points taken from two source shapefiles. You need to "
msg += "supply two input and one output shapefile names and locations."
msg += "\n\n The input shapefiles must contain a point geometry column "
msg += "named 'Shape' and a column named 'Block_No'."
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
(fileBaseName1, fileExtension1)=os.path.splitext(fileName1)

print myShapeFile1
print dirName1
dirName = dirName1 + '\*.shp'
myShapeFile2 = fileopenbox(msg="Select 2nd Point Shapefile", title=None, default=dirName, filetypes=["*.shp"])
if myShapeFile2:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit
#myShapeFile2 = "X:\Geography\LNaylor\Glamorgan\GIS\L24\XLS2SHP_L24_12Mar08.shp"
(dirName2, fileName2) = os.path.split(myShapeFile2)
(fileBaseName2, fileExtension2)=os.path.splitext(fileName2)

dirName = dirName1
#dirName = 'X:\Geography\LNaylor\Glamorgan\GIS\L24\Blocks'
defaultFile = dirName + '\PT2PL_' + fileBaseName1 + '_' + fileBaseName2 + '.shp'
#defaultFile = dirName + '\PNT2PLY_TEST.shp'
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
print "myShapeFile2 = " + myShapeFile2
print "myOutputFile = " + myOutputFile

# Read 1st input shapefile
print "\nCreating 1st array"
myBlocks1 = {}
rows1 = gp.SearchCursor(myShapeFile1)
row1 = rows1.Next()
while row1:
    pnt1 = row1.shape.GetPart()
    xx = round(float(row1.Block_No),1)
    print xx
    xx2 = [pnt1.x, pnt1.y, pnt1.z]
    #myBlocks1[xx] = row1.Block_No
    myBlocks1[xx] = xx2
    #print "\t", xx, "(", pnt1.x, pnt1.y, pnt1.z, ")"
    row1 = rows1.Next()
# Delete the row and cursor
del row1, rows1
#print myBlocks1

# Read 2nd input shapefile
print "\nCreating 2nd array"
myBlocks2 = {}
rows2 = gp.SearchCursor(myShapeFile2)
row2 = rows2.Next()
while row2:
    pnt2 = row2.shape.GetPart()
    yy = round(float(row2.Block_No),1)
    print yy
    yy2 = [pnt2.x, pnt2.y, pnt2.z]
    #myBlocks2[yy] = row2.Block_No
    myBlocks2[yy] = yy2
    #print "\t", yy, "(", pnt2.x, pnt2.y, pnt2.z, ")"
    row2 = rows2.Next()
del row2, rows2
#print myBlocks2

# Create a spatial reference object
sr = gp.CreateObject("spatialreference")

# Use a projection file to define the spatial reference's properties
sr.CreateFromFile(r'C:\Program Files\ArcGIS\Coordinate Systems\Projected Coordinate Systems\National Grids\British National Grid.prj')

# Create the output feature class using the spatial reference object
gp.CreateFeatureClass(os.path.dirname(myOutputFile),os.path.basename(myOutputFile), "POLYLINE","","","ENABLED", sr)

gp.AddField(myOutputFile,'Block','float')
print "\nCreated field 'Block'"

# Open an insert cursor for the new feature class
cur = gp.InsertCursor(myOutputFile)

matchedBlocks = 0

# Checking for matches in both arrays
print "\nChecking for matches in both arrays"
for myBlock1 in myBlocks1.keys():
    #print str(round(float(myBlock),1))
    print "myBlocks1 key => " + str(myBlock1)

    for myBlock2 in myBlocks2.keys():
        #print "\tmyBlocks2 key => " + str(myBlock2)
        if myBlock2 == myBlock1:
            print "\t", str(myBlock2), "found in both arrays"
            matchedBlocks += 1
            print "\t", fileName1, " : X =", myBlocks1[myBlock1][0], " Y =", myBlocks1[myBlock1][1], " Z =", myBlocks1[myBlock1][2]
            print "\t", fileName2, " : X =", myBlocks2[myBlock2][0], " Y =", myBlocks2[myBlock2][1], " Z =", myBlocks2[myBlock2][2]

            # Create array to hold polyline's point data
            lineArray = gp.createobject("Array")

            # Create start point of line
            pnt1 = gp.CreateObject("Point")
            pnt1.x = myBlocks1[myBlock1][0]
            pnt1.y = myBlocks1[myBlock1][1]
            pnt1.z = myBlocks1[myBlock1][2]

            # Create end point of line
            pnt2 = gp.CreateObject("Point")
            pnt2.x = myBlocks2[myBlock2][0]
            pnt2.y = myBlocks2[myBlock2][1]
            pnt2.z = myBlocks2[myBlock2][2]

            if (pnt1.x == pnt2.x) and (pnt1.y == pnt2.y):
                print "Co-ordinates are IDENTICAL"
                print pnt1.x
                pnt1.x -= 0.001
                print pnt1.x
    
            # Add points to lineArray
            lineArray.add(pnt1)
            lineArray.add(pnt2)

            #print lineArray

            row = cur.newrow()
            row.shape = lineArray
            row.block = myBlock1
            cur.InsertRow(row)

            del lineArray, pnt1, pnt2, row

            break
#        else:
#            print str(myBlock2) + "not found in both"

del cur, myBlock1, myBlocks1, myBlock2, myBlocks2

print "Found", matchedBlocks, " matching block(s)"

print "\nFINISHED :-)"

msgbox("Finished! ... (" + str(matchedBlocks) + " matching records)", title, ok_button="OK")
