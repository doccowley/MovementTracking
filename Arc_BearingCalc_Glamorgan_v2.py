# ---------------------------------------------------------------------------
title = "Arc_BearingCalc_Glamorgan.py"
msg   = "Copyright (c) 2009-11 - Andrew Cowley, University of Exeter \n\n"
# ---------------------------------------------------------------------------
# Created on: Fri Apr 24 2009 14:42
# Description: 
msg = "This script will open a specified polyline shapfile and calculate"
msg += " bearing and distance from the firstpoint to the lastpoint of each"
msg += " feature ... "
msg += "\n\nThe Excel worksheet must contain column headings of 'Easting', "
msg += "'Northing', 'Elevation' and 'PID' in any order."
# ---------------------------------------------------------------------------

# Import system modules
import os, math, string, arcgisscripting
from easygui import *
from decimal import *

'''Michael Lange <klappnase (at) freakmail (dot) de>
The Meter class provides a simple progress bar widget for Tkinter.

INITIALIZATION OPTIONS:
The widget accepts all options of a Tkinter.Frame plus the following:

    fillcolor -- the color that is used to indicate the progress of the
                 corresponding process; default is "orchid1".
    value -- a float value between 0.0 and 1.0 (corresponding to 0% - 100%)
             that represents the current status of the process; values higher
             than 1.0 (lower than 0.0) are automagically set to 1.0 (0.0); default is 0.0 .
    text -- the text that is displayed inside the widget; if set to None the widget
            displays its value as percentage; if you don't want any text, use text="";
            default is None.
    font -- the font to use for the widget's text; the default is system specific.
    textcolor -- the color to use for the widget's text; default is "black".

WIDGET METHODS:
All methods of a Tkinter.Frame can be used; additionally there are two widget specific methods:

    get() -- returns a tuple of the form (value, text)
    set(value, text) -- updates the widget's value and the displayed text;
                        if value is omitted it defaults to 0.0 , text defaults to None .
'''

import Tkinter

class Meter(Tkinter.Frame):
    def __init__(self, master, width=300, height=20, bg='white', fillcolor='orchid1',\
                 value=0.0, text=None, font=None, textcolor='black', *args, **kw):
        Tkinter.Frame.__init__(self, master, bg=bg, width=width, height=height, *args, **kw)
        self._value = value

        self._canv = Tkinter.Canvas(self, bg=self['bg'], width=self['width'], height=self['height'],\
                                    highlightthickness=0, relief='flat', bd=0)
        self._canv.pack(fill='both', expand=1)
        self._rect = self._canv.create_rectangle(0, 0, 0, self._canv.winfo_reqheight(), fill=fillcolor,\
                                                 width=0)
        self._text = self._canv.create_text(self._canv.winfo_reqwidth()/2, self._canv.winfo_reqheight()/2,\
                                            text='', fill=textcolor)
        if font:
            self._canv.itemconfigure(self._text, font=font)

        self.set(value, text)
        self.bind('<Configure>', self._update_coords)

    def _update_coords(self, event):
        '''Updates the position of the text and rectangle inside the canvas when the size of
        the widget gets changed.'''
        # looks like we have to call update_idletasks() twice to make sure
        # to get the results we expect
        self._canv.update_idletasks()
        self._canv.coords(self._text, self._canv.winfo_width()/2, self._canv.winfo_height()/2)
        self._canv.coords(self._rect, 0, 0, self._canv.winfo_width()*self._value, self._canv.winfo_height())
        self._canv.update_idletasks()

    def get(self):
        return self._value, self._canv.itemcget(self._text, 'text')

    def set(self, value=0.0, text=None):
        #make the value failsafe:
        if value < 0.0:
            value = 0.0
        elif value > 1.0:
            value = 1.0
        self._value = value
        if text == None:
            #if no text is specified use the default percentage string:
            text = str(int(round(100 * value))) + ' %'
        self._canv.coords(self._rect, 0, 0, self._canv.winfo_width()*value, self._canv.winfo_height())
        self._canv.itemconfigure(self._text, text=text)
        self._canv.update_idletasks()

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
        #print yy[0]
        #print 'String version Block ' + str(xx) + ' = ' + str(yy[0])
        if i > 0:
            myVals[xx] = str(yy[0])
    del myRows, array1, array2, xx, yy
    return myVals

##---------------------------------------------------------##

#def _demo(meter, value):
#    meter.set(value)
#    if value < 1.0:
#        value = value + 0.005
#        meter.after(10, lambda: _demo(meter, value))
#    else:
#        meter.set(value, 'Demo successfully finished')


# Define some variables/constants
yes = set(['yes','y', 'ye', ''])
no = set(['no','n'])

#Set precision for decimal module
getcontext()
getcontext().prec = 2

try:
    from easygui import *
    easygui_loaded = True
except ImportError:
    easygui_loaded = False
    print "WARNING!!! Could not load EasyGUI module\n"
    #raise SystemExit

try:
    import arcgisscripting
except ImportError:
    print "FAIL!!! Could not load ArcGIS Scripting module\n"
    raise SystemExit

##---------------------------------------------------------##

tkProgress = Tkinter.Tk(className='Script Progress')
m = Meter(tkProgress, relief='ridge', bd=3, fillcolor='cornflower blue')
m.pack(fill='x')
m.set(0.0, 'Waiting for input ...')
#m.after(1000, lambda: _demo(m, 0.0))
#tkProgress.mainloop()

if easygui_loaded == True: # show a Continue/Cancel dialog
    if not ccbox(msg + "\n\nDo you want to continue?", title):     # user chose Cancel
        raise SystemExit
else:   # command line choice
    print msg + "\n\nDo you want to continue?"
    choice = raw_input().lower()
    if choice in yes:
        print "OK, we shall continue ;-) ... "
    elif choice in no:
        raise SystemExit
    else:
        raise SystemExit

#Get input Shapefile
m.set(0.0, 'Waiting for input ... source shape file')
if easygui_loaded == True:
    defDir = 'X:\Geography\LNaylor\Glamorgan\GIS\L24\Blocks'
    defDir = 'W:\Scripts\PythonScripts\Larissa'
    myShapeFile = fileopenbox(msg='Select Block Shape File', title=None, default = defDir + '\*.shp', filetypes=["*.shp"])
    if not myShapeFile:     # file chosen
        raise SystemExit
else:
    print "Please enter soure Shapefile location including path (e.g. X:\A-Folder\An-Excel-File.xls)"
    myShapeFile = raw_input().lower()
    if not os.path.exists( myShapeFile ):
        print "File does not exist ... bye!"
        raise SystemExit
print "Source file : '" + myShapeFile + "' exists :-)"

#Create default directory
(dirName, fileName) = os.path.split(myShapeFile)

m.set(0.0, 'Waiting for input ... Excel Block Size file')
dirName1 = dirName + '\*.xls'
myExcelFile = fileopenbox(msg='Select Block Size File', title=None, default=dirName1, filetypes=["*.xls"])
if myShapeFile:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit

m.set(0.0, 'Creating ArcGIS objects ...')
# Create the Geoprocessor object
gp = arcgisscripting.create()

# Set the Geoprocessing environment...
CurrentWorkingDirectory = os.getcwd()
print "cwd    = " + CurrentWorkingDirectory
print "myShapeFile = " + myShapeFile

m.set(0.0, 'Reading block sizes ...')
print "Reading block sizes from ", myExcelFile
blockArray = getBlockSizes(myExcelFile)
#print blockArray

m.set(0.0, 'Checking Shapefile for Schema Lock ...')
lockTest = gp.TestSchemaLock(myShapeFile)
# If the lock can be applied, continue ... else QUIT
print 'Schema Lock test result : ' + str(lockTest)
if lockTest == 'FALSE':
    print "Unable to acquire the schema lock to add/delete/modify fields in the shapefile ... try closing ArcMap/ArcCatalog etc"
    msgbox("Unable to acquire the schema lock! Cannot add/delete/modify fields in the shapefile ... try closing ArcMap/ArcCatalog", title, ok_button="OK")
    tkProgress.destroy()
    raise SystemExit

m.set(0.0, 'Checking Dist/Dir/A_Axis/Moved fields ...')
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
if gp.ListFields(myShapeFile,'ScrNote').Next():
    print "Field 'ScrNote' already exists -> Deleting ..."
    gp.deletefield (myShapeFile, 'ScrNote')

gp.AddField(myShapeFile,'Distance','float')
print "Created field 'Distance'"
gp.AddField(myShapeFile,'Direction','float')
print "Created field 'Direction'"
gp.AddField(myShapeFile,'A_Axis','float')
print "Created field 'A_Axis'"
gp.AddField(myShapeFile,'Moved','short')
print "Created field 'Moved'"
gp.AddField(myShapeFile,'ScrNote','text','50')
print "Created field 'ScrNote'"
################################

myRows = gp.GetCount_management(myShapeFile)
rows = gp.UpdateCursor(myShapeFile)
row = rows.Next()
progress = 0

while row:
    #print 'Orig : ' + str(row.block)
    myBlock = str(round(row.block,1)).rstrip('0').rstrip('.')
    #print 'Rstrip : ' + str(myBlock)
    print 'Block : ' + str(myBlock)
    
    A_Axis = 'Empty'
    A_Axis  = str(blockArray.get(myBlock, A_Axis))
    #print A_Axis
    
    x1, y1 = row.shape.FirstPoint.split(" ")
    x2, y2 = row.shape.LastPoint.split(" ")
    #print x1, y1, x2, y2

    deltaE = float(x2) - float(x1)
    deltaN = float(y2) - float(y1)

    dist = math.sqrt(math.pow(deltaE,2) + math.pow(deltaN,2))
    angle = (90 - (math.atan2(deltaN, deltaE) / math.pi * 180) + 360 ) % 360
    row.Direction = angle
    row.Distance = dist
    #print angle, dist

    if A_Axis <> 'Empty':
        if dist > (0.5 * float(A_Axis)):
            moved = 1
        else:
            moved = 0
        row.A_Axis = A_Axis
        row.Moved = str(moved)
        #print row.A_Axis, row.Moved
    else:
        row.ScrNote = 'Not found in Excel blocks file'

    del myBlock, x1, x2, y1, y2, deltaE, deltaN, dist, angle, A_Axis

    #Update row
    rows.UpdateRow(row)

    #Update progress
    progress = Decimal(progress + 1)
    myCompleted = Decimal(progress+1)/Decimal(myRows)
    m.set(float(myCompleted))

    row = rows.Next()

del gp, myExcelFile, blockArray, myShapeFile, row, rows

print "\nFINISHED :-)"
msgbox("Finished! ... (" + str(myRows) + " records)", title, ok_button="OK")

tkProgress.destroy()

