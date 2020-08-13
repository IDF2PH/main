#
# IDF2PHPP: A Plugin for exporting an EnergyPlus IDF file to the Passive House Planning Package (PHPP). Created by blgdtyp, llc
# 
# This component is part of IDF2PHPP.
# 
# Copyright (c) 2020, bldgtyp, llc <info@bldgtyp.com> 
# IDF2PHPP is free software; you can redistribute it and/or modify 
# it under the terms of the GNU General Public License as published 
# by the Free Software Foundation; either version 3 of the License, 
# or (at your option) any later version. 
# 
# IDF2PHPP is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of 
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the 
# GNU General Public License for more details.
# 
# For a copy of the GNU General Public License
# see <http://www.gnu.org/licenses/>.
# 
# @license GPL-3.0+ <http://spdx.org/licenses/GPL-3.0+>
#
"""
Takes in a list of geometry (brep) and 'orders' it left->right. Used to order
window surfaces or similar for naming 0...1...2...etc...
Make sure the surface normals are all pointing 'out' for this to work properly.
-
EM Feb. 25, 2020
    Args:
        _geom: (list) A list of Breps to 'order'.
    Returns:
        geomOrdered_: A list of the Breps, ordered from left-to-right (when viewed from 'front')
"""

ghenv.Component.Name = "BT_OrderGeometry"
ghenv.Component.NickName = "Order Geometry"
ghenv.Component.Message = 'FEB_25_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import rhinoscriptsyntax as rs
import Rhino.Geometry

def getOrderedGeom(_geomList):
    """
    Takes in a list of geometry (brep) and 'orders' it left->right. Used to order
    window surfaces or similar for naming 0...1...2...etc...
    Make sure the surface normals are all pointing 'out' for this to work properly.
    """
    
    setXs = []
    setYs = []
    setZs = []
    Centroids = []
    CentroidsX = []
    #Go through all the selected window geometry, get the information needed
    for i in range( len(_geomList) ):
        # Get the surface normal for each window
        windowBrep = rs.coercebrep(_geomList[i])
        surfaceList =  windowBrep.Surfaces
        for eachSurface in surfaceList:
            srfcCentroid = rs.SurfaceAreaCentroid(eachSurface)[0]
            b, u, v = eachSurface.ClosestPoint(srfcCentroid)
            srfcNormal = eachSurface.NormalAt(u, v)
            setXs.append( srfcNormal.X )
            setYs.append( srfcNormal.Y )
            setZs.append( srfcNormal.Z )
            Centroids.append( srfcCentroid )
            CentroidsX.append(srfcCentroid.X)
    
    # Find the average Normal Vector of the set
    px = sum(setXs) / len(setXs)
    py = sum(setYs) / len(setYs)
    pz = sum(setZs) / len(setZs)
    avgNormalVec = Rhino.Geometry.Point3d(px, py, pz)
    
    # Find a line through all the points and its midpoint
    #print Centroids
    fitLine = rs.LineFitFromPoints(Centroids)
    newLine = Rhino.Geometry.Line(fitLine.From, fitLine.To)
    
    # Find the Midpoint of the Line
    #rs.CurveMidPoint(newLine) #Not working....
    midX = (fitLine.From.X + fitLine.To.X) / 2
    midY = (fitLine.From.Y + fitLine.To.Y) / 2
    midZ = (fitLine.From.Z + fitLine.To.Z) / 2
    lineMidpoint = Rhino.Geometry.Point3d(midX, midY, midZ)
    
    # New Plane
    newAvgWindowPlane = rs.CreatePlane(lineMidpoint, avgNormalVec, [0,0,1] )
    
    # Rotate new Plane to match the window
    finalPlane = rs.RotatePlane(newAvgWindowPlane, 90, [0,0,1])
    
    
    # Plot the window Centroids onto the selection Set Plane
    centroidsReMaped = []
    for eachCent in Centroids:
        centroidsReMaped.append( finalPlane.RemapToPlaneSpace(eachCent) )
    
    # Sort the original geometry, now using the ordered centroids as key
    return [x for _,x in sorted(zip(centroidsReMaped, _geom))]

if len(_geom)>1:
    geomOrdered_ = getOrderedGeom(_geom)
else:
    geomOrdered_ = _geom[0]
