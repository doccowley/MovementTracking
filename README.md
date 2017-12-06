# MovementTracking
Python / ArcGIS scripts to manage movement data ... e.g. boulders along shoreline

Script Name                             Purpose
-----------                             -------
CreateEdgePointSHPFromExcel.pyc         Main script used to convert Excel GPS data
CreateMultiPointPolyFrom1PointSHP.pyc   Convert point data from (1) to poly-line
CreatePolyFrom2PointSHPs.pyc            Convert point data from two  poly-lines created using (1) to block movement
Arc_BearingCalc_Glamorgan_v2.pyc        Add movement distance and bearing to output shapefile from (4)
Arc_BearingCalc_SingleSHP.py            As above but for a single polyline shapefile e.g. layer edge
ExportAttTableAndShapeAsCSV.py          Export point shapefile attribute data including X,Y(,Z,M?) to .CSV file
ExportAttTableAndShapeAsCSV_v2.py       As above but can handle polylines … work-in-progress/experimental
CheckSHPDetails.py                      Quickly check info on spatial reference in use
CreatePolygonFromPolylineSHP.py         Create a block island polygon etc

Excel files need to have specific layout/columns for above to work.

See 'GIS Python Scripts.docx' for further information
