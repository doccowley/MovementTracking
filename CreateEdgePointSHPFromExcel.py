# ---------------------------------------------------------------------------
title = "CreatePointSHPFromExcel.py"
msg   = "Copyright (c) 2009 - Dr Andrew Cowley, University of Exeter \n\n"
# ---------------------------------------------------------------------------
# Created on: Thu Apr 30 2009 14:42
# Description: 
msg += "This script will create a Point ShapeFile for an edge or path from "
msg += "Excel. You need to supply an input Excel file (the last active "
msg += "worksheet will be used) and an output shapefile name/location. "
msg += "\n\nThe Excel worksheet must contain column headings of 'Easting',"
msg += "'Northing' and 'Elevation' in any order."
msg += "\n\nYou will be given the option to read in other columns that this "
msg += "script recognises."
# ---------------------------------------------------------------------------

# Import system modules
import array, os, math, time
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

def getExcelArray(ExcelFile):
    mySheets = []

    xl = Dispatch("Excel.Application")
    xl.Visible = 0
    #xl.Workbooks.Add()
    wb = xl.Workbooks.Open(ExcelFile)
    sheetCount = wb.Worksheets.Count
    if sheetCount > 1:
        for sheetnum in range (1, sheetCount + 1):
            print "Sheet " + str(sheetnum) + ") " + wb.Worksheets(sheetnum).Name
            mySheets.append(wb.Worksheets(sheetnum).Name)           
        print mySheets

        msg ="Which worksheet are you interested in?"
        title = "Create Edge Point SHP From Excel"
        #choices = ["Vanilla", "Chocolate", "Strawberry", "Rocky Road"]
        sheetChoice = choicebox(msg, title, mySheets)
        print "You chose : " + str(sheetChoice)
        sh = wb.Worksheets(sheetChoice)
    else:
        print "Only one sheet"
        sh = xl.ActiveWorkbook.ActiveSheet

    myRows = sh.UsedRange.Rows.Count
    myCols = sh.UsedRange.Columns.Count
    array = sh.Range(sh.Cells(1,1),sh.Cells(myRows,myCols)).Value
    #xlApp.ActiveSheet.Cells(1,1).Value = 'Python Rules!'
    #xlApp.ActiveWorkbook.ActiveSheet.Cells(1,2).Value = 'Python Rules 2!'
    #xlApp.ActiveWorkbook.Close(SaveChanges=0) # see note 1
    xl.Quit()
    xl.Visible = 0 # see note 2
    del xl

    return array, myCols, myRows

def closeProgressBar():
    tkProgress.destroy()

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

try:
    from win32com.client import Dispatch
except ImportError:
    print "FAIL!!! Could not load win32com.client Dispatch\n"
    raise SystemExit

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

# Show progress bar
tkProgress = Tkinter.Tk(className='Script Progress')
m = Meter(tkProgress, relief='ridge', bd=3, fillcolor='cornflower blue')
m.pack(fill='x')
m.set(0.0, 'Waiting for input ...')
#m.after(1000, lambda: _demo(m, 0.0))
#tkProgress.mainloop()

#Get input file
m.set(0.0, 'Waiting for input ... Excel source file')
if easygui_loaded == True:
    #myExcelFile = fileopenbox(msg="Select 1st Excel File", title=None, default='X:\Geography\LNaylor\Glamorgan\GIS\*.xls', filetypes=["*.xls"])
    myExcelFile = fileopenbox(msg="Select 1st Excel File", title=None, default='C:\Documents and Settings\ac278\Desktop\*.xls', filetypes=["*.xls"])
    if not myExcelFile:     # user chose 'Cancel'
        raise SystemExit
else:
    print "Please enter 1st Excel File including path (e.g. X:\A-Folder\An-Excel-File.xls)"
    myExcelFile = raw_input().lower()
    if not os.path.exists( myExcelFile ):
        print "File does not exist ... bye!"
        raise SystemExit
print "Source file : '" + myExcelFile + "' exists :-)"

#Create default output filename
(dirName, fileName) = os.path.split(myExcelFile)
(fileBaseName, fileExtension)=os.path.splitext(fileName)
if not dirName[-1:] in ['\\','/']: #check dirName ends in forward or back slash, if not add
    dirName += '/'
defaultFile = dirName + 'XLS2SHP_' + fileBaseName.replace(" ", "_") + '.shp' #Also replace spaces with underscores
print "Default output file : " + defaultFile

#Choose output file
m.set(0.0, 'Waiting for input ... Output shape file')
if easygui_loaded == True: # show a file save dialog
    myOutputFile = filesavebox(msg=None, title=None, default=defaultFile, filetypes=["*.shp"])
    if not myOutputFile:     # user chose Cancel
        raise SystemExit
else: #command line
    print "Please enter output filename including path (e.g. X:\A-Folder\An-OutputFile.shp)"
    myOutputFile = raw_input().lower()
    (OutdirName, OutfileName) = os.path.split(myOutputFile)
    if not os.path.exists(OutdirName):
        print "Path '" + OutdirName + "' does not exist ... bye!"
        raise SystemExit
print myOutputFile
print "Output file : '" + myOutputFile
#raise SystemExit

# Set the Geoprocessing environment...
CurrentWorkingDirectory = os.getcwd()
print "cwd    = " + CurrentWorkingDirectory
print "myExcelFile = " + myExcelFile
print "myOutputFile = " + myOutputFile

m.set(0.0, 'Reading input Excel file ...')
(my1stArray, my1stCols, my1stRows) = getExcelArray(myExcelFile)

myExtraCols = []
myOtherCols = {}

print str(my1stCols) + ' Cols : ' + str(my1stRows) + ' Rows\n'

#print my1stArray

print "\nAdding fields ..."
for i in xrange(my1stCols):
    print my1stArray[0][i][0:8]
    if my1stArray[0][i] == "Easting":
        myEasting = i
    elif my1stArray[0][i] == "Northing":
        myNorthing = i
    elif my1stArray[0][i] == "Elevation":
        myElevation = i
    elif str(my1stArray[0][i]) in ('PID', 'CSM_Code','Layer','Descript'):
        myOtherCols[i] = my1stArray[0][i][0:8] #Dict
    else:
        myExtraCols.append(str(i) + ") " + str(my1stArray[0][i][0:8]))

if 'myEasting' in locals():
    print "myEasting exists in column" + str(myEasting)
if 'myNorthing' in locals():
    print "myNorthing exists in column" + str(myNorthing)
if 'myElevation' in locals():
    print "myElevation exists in column" + str(myElevation)

m.set(0.0, 'Waiting for input ... additional fields')
print "Do you wish to add any of these fields?"
print myOtherCols
myExtraCols = multchoicebox(msg='Pick as many items as you like.', title=' ', choices=(myExtraCols)) #EasyGUI multchoicebox
for ExtraCol in myExtraCols:
    index, item = ExtraCol.split(") ")    
    print index + " => " + item
    myOtherCols[int(index)] = item
print myOtherCols

#Check if want Easting, Northing and Elevation in Attribute table as well as shape column
m.set(0.0, 'Waiting for input ... E,N,Z as fields?')
bENZAttribs = ynbox(msg='Do you want Easting, Northing and Elevation data in Attribute table as well as in shape column?', title=' ', choices=('Yes', 'No'), image=None)
if bENZAttribs == 1:
    print 'You answered yes to adding ENZ as attributes'
else:
    print 'You answered no to adding ENZ as attributes!'

m.set(0.0, 'Creating ArcGIS objects ...')
#raise SystemExit

# Create the Geoprocessor object
gp = arcgisscripting.create()

# Create a spatial reference object
sr = gp.CreateObject("spatialreference")

# Use a projection file to define the spatial reference's properties
sr.CreateFromFile(r'C:\Program Files\ArcGIS\Coordinate Systems\Projected Coordinate Systems\National Grids\British National Grid.prj')

# Create the output feature class using the spatial reference object
gp.CreateFeatureClass(os.path.dirname(myOutputFile),os.path.basename(myOutputFile), "Point","","","ENABLED", sr)

print "\nAdding Fields ..."
if bENZAttribs == 1:
    print "\t" + str(myEasting) + " => Easting"
    gp.AddField(myOutputFile,'Easting','float')

    print "\t" + str(myNorthing) + " => Northing"
    gp.AddField(myOutputFile,'Northing','float')

    print "\t" + str(myElevation) + " => Elevation"
    gp.AddField(myOutputFile,'Elevation','float')
    
for key in myOtherCols:
    if myOtherCols[key] == 'Date':
        print "Use of 'Date' as column header ... changing to 'Date_'!!"
        myOtherCols[key] = 'Date_'
    if myOtherCols[key] == 'Date ':
        print "Use of 'Date ' as column header ... changing to 'Date_'!!"
        myOtherCols[key] = 'Date_'
        
    fieldname = gp.ValidateFieldName(myOtherCols[key], os.path.dirname(myOutputFile))

    if fieldname in ('Block_No'):
        print 1
        #fieldtype = 'float'
        fieldtype = 'text' # Changed 08/07/2011 to work around block labelling 18.1 18a etc
    elif fieldname in ('Date_', 'DateTime', 'Date', 'Date '):
        print 2
        fieldtype = 'text'
    else:
        print 3
        fieldtype = 'text'

    print "\t" + str(key) + " => " + fieldname + " (" + fieldtype + ")"
    gp.AddField(myOutputFile,fieldname,fieldtype)

print myOtherCols

# Open an insert cursor for the new feature class
cur = gp.InsertCursor(myOutputFile)

# Create an array and point object needed to create features
#lineArray = gp.CreateObject("Array")
pnt = gp.CreateObject("Point")

print "\nAdding " + str(my1stRows - 1) + " rows ..."
        
for j in xrange(my1stRows - 1):
    #print str(my1stArray[j+1][1]) + " " + str(my1stArray[j+1][myEasting]) + " " + str(my1stArray[j+1][myNorthing]) + " " + str(my1stArray[j+1][myElevation])
    if my1stArray[j+1][myEasting] <> None:
        pnt.x = my1stArray[j+1][myEasting]
        pnt.y = my1stArray[j+1][myNorthing]
        pnt.z = my1stArray[j+1][myElevation]
        row = cur.NewRow()
        row.shape = pnt
        if bENZAttribs == 1:
            #print 'Adding ENZ as attributes'
            row.SetValue('Easting', my1stArray[j+1][myEasting])
            row.SetValue('Northing', my1stArray[j+1][myNorthing])
            row.SetValue('Elevation', my1stArray[j+1][myElevation])
        for myField in myOtherCols:
            if my1stArray[j+1][myField] <> None:
                if myOtherCols[myField] in ('Date_', 'DateTime', 'Date ', 'Date'):
                    myDate = str(my1stArray[j+1][myField])
                    myDate = myDate.replace(':', "_")
                    row.SetValue(myOtherCols[myField], myDate)
                else:
                    print myField, myOtherCols[myField], j, my1stArray[j+1][myField]
                    row.SetValue(myOtherCols[myField], my1stArray[j+1][myField])
        cur.InsertRow(row)

        #Update progress
        myCompleted = Decimal(j+1)/Decimal(my1stRows-1)
        #print float(myCompleted)
        m.set(float(myCompleted))

del gp, sr, i, j

print "\nFINISHED :-)"
msgbox("Finished! ... (" + str(my1stRows - 1) + " records)", title, ok_button="OK")

closeProgressBar()

