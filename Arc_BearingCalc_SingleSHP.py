# ---------------------------------------------------------------------------
title = "Arc_BearingCalc_Glamorgan.py"
msg   = "Copyright (c) 2009 - Andrew Cowley, University of Exeter \n\n"
# ---------------------------------------------------------------------------
# Created on: Fri Apr 24 2009 14:42
# Description: 
msg = "This script will open a specified polyline shapfile and calculate"
msg += " bearing and distance from the firstpoint to the lastpoint of each"
msg += " feature ... "
msg += "\n\nThe Excel worksheet must contain column headings of 'Easting', "
msg += "'Northing', 'Elevation' and 'PID' in any order."

msg += " feature it contains"
#L24_BlockSizeData.xls
# ---------------------------------------------------------------------------

# Import system modules
import os, math, string, arcgisscripting

from easygui import *
msg += "\n\nDo you want to continue?"
if ccbox(msg, title):     # show a Continue/Cancel dialog
    pass  # user chose Continue
else:  # user chose Cancel
    raise SystemExit

myShapeFile = fileopenbox(msg='Select Block Shape File', title=None, default='X:\Geography\LNaylor\Glamorgan\GIS\L24\Blocks\*.shp', filetypes=["*.shp"])
if myShapeFile:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit

myExcelFile = fileopenbox(msg='Select Block Size File', title=None, default='X:\Geography\LNaylor\Glamorgan\GIS\L24\*.xls', filetypes=["*.xls"])
if myShapeFile:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit


# Create the Geoprocessor object
gp = arcgisscripting.create()

# Set the Geoprocessing environment...
CurrentWorkingDirectory = os.getcwd()
print "cwd    = " + CurrentWorkingDirectory
print "myShapeFile = " + myShapeFile

def getBlockSizes(ExcelFile):
    from win32com.client import Dispatch

    xlApp = Dispatch("Excel.Application")
    xlApp.Visible = 0
    #xlApp.Workbooks.Add()
    xlApp.Workbooks.Open(ExcelFile)
    myRows = xlApp.ActiveWorkbook.ActiveSheet.UsedRange.Rows.Count
    array1 = xlApp.ActiveWorkbook.ActiveSheet.Range("a1:a" + str(myRows)).Value
    array2 = xlApp.ActiveWorkbook.ActiveSheet.Range("b1:b" + str(myRows)).Value
    #xlApp.ActiveSheet.Cells(1,1).Value = 'Python Rules!'
    #xlApp.ActiveWorkbook.ActiveSheet.Cells(1,2).Value = 'Python Rules 2!'
    #xlApp.ActiveWorkbook.Close(SaveChanges=0) # see note 1
    xlApp.Quit()
    xlApp.Visible = 0 # see note 2
    del xlApp

    i=0
    myVals = {}

    myVals['None'] = 'None'
    for i in range(1, myRows):
        xx = str(array1[i][0]).rstrip('0').rstrip('.')
        yy = array2[i]
        if i > 0:
            myVals[xx] = yy[0]
        #print xx + " " + str(yy[0])
#    else:
#            print 'The for loop is over\n'
    del myRows, array1, array2, xx, yy
    return myVals

#myExcelFile = 'X:\Geography\LNaylor\Glamorgan\GIS\L24\L24_BlockSizeData.xls'
print "Reading block sizes from ", myExcelFile
blockArray = getBlockSizes(myExcelFile)
print blockArray

#################################
if gp.ListFields(myShapeFile,'Distance').Next():
    print "Field 'Distance' already exists -> Deleting ..."
    gp.deletefield (myShapeFile, 'Distance')
if gp.ListFields(myShapeFile,'Direction').Next():
    print "Field 'Direction' already exists -> Deleting ..."
    gp.deletefield (myShapeFile, 'Direction')
if gp.ListFields(myShapeFile,'A_Axis').Next():
    print "Field 'A_Axis' already exists -> Deleting ..."
    gp.deletefield (myShapeFile, 'A_Axis')
if gp.ListFields(myShapeFile,'Moved').Next():
    print "Field 'Moved' already exists -> Deleting ..."
    gp.deletefield (myShapeFile, 'Moved')

gp.AddField(myShapeFile,'Distance','float')
print "Created field 'Distance'"
gp.AddField(myShapeFile,'Direction','float')
print "Created field 'Direction'"
gp.AddField(myShapeFile,'A_Axis','float')
print "Created field 'A_Axis'"
gp.AddField(myShapeFile,'Moved','text')
print "Created field 'Moved'"
################################

#rows = gp.SearchCursor(myShapeFile)
rows = gp.UpdateCursor(myShapeFile)
row = rows.Next()

block1_flag = 0

#print "Block, E1, N1, E2, N2, Bearing, Distance, A_Axis, Moved"
while row:
    if block1_flag == 0:
        x1, y1 = row.shape.FirstPoint.split(" ")
        block1_flag = 1
    else:     
    #    myBlock = string.rstrip(str(round(row.block,1)),'.0')
        myBlock = str(round(row.block,1)).rstrip('0').rstrip('.')
        x2, y2 = row.shape.LastPoint.split(" ")

        deltaE = float(x2) - float(x1)
        deltaN = float(y2) - float(y1)

        print deltaE
        print deltaN

        dist = math.sqrt(math.pow(deltaE,2) + math.pow(deltaN,2))
        angle = (90 - (math.atan2(deltaN, deltaE) / math.pi * 180) + 360 ) % 360
        A_Axis = str(blockArray[myBlock])
        if A_Axis <> 'None':
            A_Axis = blockArray[myBlock] 
            if dist > (0.5 * A_Axis):
                moved = 1
            else:
                moved = 0

        #print row.block + ", " +x1 + ", " + y1 + ", " + x2 + ", " + y2 + ", " + str(angle) + ", " + str(dist) + ", " + str(A_Axis) + ", " + str(moved)

        row.Direction = angle
        row.Distance = dist
        if A_Axis <> 'None':
            row.A_Axis = A_Axis
            row.Moved = str(moved)
        rows.UpdateRow(row)

        x1 = x2
        y1 = y2
        del myBlock, x2, y2, deltaE, deltaN, dist, angle, A_Axis
   
    row = rows.Next()

del x1, y1, gp, myExcelFile, blockArray, myShapeFile, row, rows

print "\nFINISHED :-)"
msgbox("Finished!", title, ok_button="OK")

