# ---------------------------------------------------------------------------
title = "CheckSHPDetails.py"
msg   = "Copyright (c) 2009 - Andrew Cowley, University of Exeter \n\n"
# ---------------------------------------------------------------------------
# Created on: Wed Apr 29 2009 15:46
# Description: 
msg += "This script will check a specified shape file for details such as "
msg += "coordinate system etc."
# ---------------------------------------------------------------------------

# Import system modules
import os, arcgisscripting

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

# Create the Geoprocessor object
gp = arcgisscripting.create()

# Describe a feature class
desc = gp.Describe(myShapeFile)
type = desc.ShapeType

# Get the shape type (Polygon, Point, Polyline) of the feature class
type = desc.ShapeType

# Get the spatial reference 
sr = desc.SpatialReference

msg =  "Shape File Name = " + fileName + "\n"
msg += "File Location = " + dirName  + "\n\n"
msg += "Geometry Field Name = " + desc.ShapeFieldName + "\n"
msg += "Geometry Type = " + type + "\n\n"
msg += "Spatial Reference Type = " + sr.Type + "\n"
msg += "Spatial Reference Name = " + sr.Name + "\n\n"
if sr.Type == 'Projected':
    msg += "PCS Name = " + str(sr.PCSName) + "\n"
    msg += "Central Meridian = " + str(sr.CentralMeridian) + "\n"
    msg += "Longitude of Origin = " + str(sr.LongitudeOfOrigin) + "\n"
    msg += "False Easting = " + str(sr.FalseEasting) + "\n"
    msg += "False Northing = " + str(sr.FalseNorthing) + "\n"
    msg += "Azimuth = " + str(sr.Azimuth) + "\n"
    msg += "Classification = " + str(sr.Classification) + "\n"
    msg += "ProjectionName = " + str(sr.ProjectionName) + "\n"
    msg += "Linear Unit Name = " + str(sr.LinearUnitName) + "\n"
else:
    msg += "GSC Name = " + str(sr.GCSName) + "\n"
    msg += "Spheroid Name = " + str(sr.SpheroidName) + "\n"
    msg += "Semi Major Axis = " + str(sr.SemiMajorAxis) + "\n"
    msg += "Semi Minor Axis = " + str(sr.SemiMinorAxis) + "\n"
    msg += "Flattening = " + str(sr.Flattening) + "\n"
    msg += "Longitude = " + str(sr.Longitude) + "\n"
    msg += "Radians Per Unit = " + str(sr.RadiansPerUnit) + "\n"
    msg += "Datum Name = " + str(sr.DatumName) + "\n"
    msg += "Prime Meridian Name = " + str(sr.PrimeMeridianName) + "\n"
    msg += "Angular Unit Name = " + str(sr.AngularUnitName) + "\n"

msgbox(msg, title, ok_button="OK")
