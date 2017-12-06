# ---------------------------------------------------------------------------
title = "CreatePolygonFromPolylineSHP.py"
msg   = "Copyright (c) 2009 - Andrew Cowley, University of Exeter \n\n"
# ---------------------------------------------------------------------------
# Created on: Thu Apr 30 2009 14:42
# Description: 
# Create Point ShapeFile from Excel.
msg += "This script will create a polygon shapefile from an input "
msg += "polyline shapefile. You need to supply input and output "
msg += "shapefile names and locations. ONLY THE GEOMETRY IS COPIED!"
# ---------------------------------------------------------------------------

#Helpful sources of info used
#http://webhelp.esri.com/arcgisdesktop/9.3/index.cfm?TopicName=Reading_geometries

# Import system modules
import os, math, arcgisscripting, string #win32com.client

from easygui import *

#http://tkinter.unpythonic.net/wiki/ProgressMeter
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

##---------------------------------------------------------##

msg += "\n\nDo you want to continue?"
if ccbox(msg, title):     # show a Continue/Cancel dialog
    pass  # user chose Continue
else:  # user chose Cancel
    raise SystemExit

#def _demo(meter, value):
#    meter.set(value)
#    if value < 1.0:
#        value = value + 0.005
#        meter.after(10, lambda: _demo(meter, value))
#    else:
#        meter.set(value, 'Demo successfully finished')

tkProgress = Tkinter.Tk(className='Script Progress')
m = Meter(tkProgress, relief='ridge', bd=3, fillcolor='cornflower blue')
m.pack(fill='x')
m.set(0.0, 'Waiting for input ...')
#m.after(1000, lambda: _demo(m, 0.0))
#tkProgress.mainloop()

m.set(0.0, 'Select input polyline shape file...')
#myPointDefault = 'X:\Geography\LNaylor\Glamorgan\GIS\L24\*.shp'
myPointDefault = 'D:\Scratch\*.shp'
myShapeFile1 = fileopenbox(msg="Select Polyline Shapefile", title=None, default=myPointDefault, filetypes=["*.shp"])
if myShapeFile1:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit
#myShapeFile1 = "X:\Geography\LNaylor\Glamorgan\GIS\L24\XLS2SHP_L24_09Mar08.shp"
(dirName1, fileName1) = os.path.split(myShapeFile1)
(fileBaseName1, fileExtension1)=os.path.splitext(fileName1)

m.set(0.0, 'Select output polygon shapefile ...')
dirName = dirName1
#dirName = 'X:\Geography\LNaylor\Glamorgan\GIS\L24\Test'
#defaultFile = dirName + '\PLYL2PLYG_TEST.shp'
defaultFile = dirName + '\PLYL2PLYG_' + fileBaseName1 + '.shp'
myOutputFile = filesavebox(msg=None, title=None, default=defaultFile, filetypes=["*.shp"])
if myOutputFile:     # file chosen
    pass
else:  # user chose Cancel
    raise SystemExit
print "myOutputFile = " + myOutputFile

m.set(0.0, 'Creating ArcGIS objects ...')
# Create the Geoprocessor object
gp = arcgisscripting.create(9.3)
#gp = win32com.client.Dispatch('esriGeoprocessing.GPDispatch')

#Turn on overwrite of existing output files
gp.overwriteoutput = True

# Set the Geoprocessing environment...
CurrentWorkingDirectory = os.getcwd()
print "cwd          = " + CurrentWorkingDirectory
print "myShapeFile1 = " + myShapeFile1
print "myOutputFile = " + myOutputFile + "\n"

# Create a spatial reference object
sr = gp.CreateObject("spatialreference")

# Use a projection file to define the spatial reference's properties
sr.CreateFromFile(r'C:\Program Files\ArcGIS\Coordinate Systems\Projected Coordinate Systems\National Grids\British National Grid.prj')

# Create the output feature class using the spatial reference object
gp.CreateFeatureClass(os.path.dirname(myOutputFile),os.path.basename(myOutputFile), "POLYGON","","","ENABLED", sr)

myOtherCols = []

print "Input Shapefile Fields:"
print "FieldName   |Editable|IsNullable|Required|Length|Type    |Precision|Scale"
print "------------|--------|----------|--------|------|--------|---------|-----"
for field in fields:
    print field.Name.ljust(12), str(field.Editable).ljust(8), str(field.IsNullable).ljust(10), str(field.Required).ljust(8), str(field.Length).ljust(6), field.Type.ljust(8), field.Precision, "        ", field.Scale
    if field.Name not in ('FID', 'Shape', 'Id'):
        if field.Type == 'Single':
            myFieldType = 'FLOAT'
        elif field.Type == 'Double':
            myFieldType = 'DOUBLE'
        elif field.Type == 'String':
            myFieldType = 'TEXT'
        else:
            myFieldType = field.Type

        #Add the additional field
        gp.AddField_management(myOutputFile, field.Name, myFieldType, field.Precision, field.Scale, field.Length, "", field.IsNullable, field.Required)

        #Add field.namne to myOtherCols list
        myOtherCols.append(field.name)
print "-------------------------------------------------------------------------\n"
    
# Open an insert cursor for the new feature class
curInsert = gp.InsertCursor(myOutputFile)

# Create array to hold polyline's point data
shapeArray = gp.createobject("Array")

#Identify input shapefile's geometry field
desc = gp.Describe(myShapeFile1)
fields = desc.Fields
shapefieldname = desc.ShapeFieldName

# Open an search cursor for the polyline feature class
curSearch = gp.SearchCursor(myShapeFile1)

maxparts = 1

m.set(0.0, 'Working ...')
row = curSearch.Next()
while row:
    # Create the geometry object
    feat = row.GetValue(shapefieldname)

    # Print the current multipoint's ID
    #print "Feature " + str(row.getvalue(desc.OIDFieldName)) + ":"

    partnum = 0

    # Count the number of points in the current multipart feature
    partcount = feat.PartCount

    # Enter while loop for each part in the feature (if a singlepart feature this will occur only once)
    while partnum < partcount:
        # Print the part number
        #print "Part " + str(partnum) + ":"

        part = feat.GetPart(partnum)
        pnt = part.Next()
        pntcount = 0
        # Enter while loop for each vertex
#        while pnt:
#            # Print x,y(,z) coordinates of current point
#            #
#            if str(pnt.z) == "1.#QNAN":
#                print pnt.x, pnt.y
#            else:
#                print pnt.x, pnt.y, pnt.z
#                
#            pnt = part.Next()
#            pntcount += 1
#            # If pnt is null, either the part is finished or there is an 
#            #   interior ring
#            #
#            if not pnt: 
#                pnt = part.Next()
#                if pnt:
#                    print "Interior Ring:"
        partnum += 1
    print "Feature " + str(row.getvalue(desc.OIDFieldName)) + " : " + str(partnum) + " Part(s)"
    if partnum > maxparts:
        maxparts = partnum

    #Insert feature into polygon shapefile
    newRow = curInsert.newrow()
    newRow.shape = feat
    for myField in myOtherCols:
        newRow.SetValue(myField, row.GetValue(myField))
    curInsert.InsertRow(newRow)
    
    row = curSearch.Next()

if maxparts > 1:
    print "\nOne of the features has more than one part!\n(Some of the polygons might not look as you expect)"
    msgbox("One of the features has more than one part!\n(Some of the polygons might not look as you expect)")
    
tkProgress.destroy()

print "\nFINISHED :-)"
msgbox("Finished!", title, ok_button="OK")

del row, pnt, curSearch, curInsert
