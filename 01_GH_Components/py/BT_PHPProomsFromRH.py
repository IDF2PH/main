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
Will try and figure out the 'Rooms' inside each of the HB zones.
-
> Rooms hold onto floor-area data, Ventilation system configurations, fresh-air flow rate, etc.
> This component allows for user-input of geometry for floor areas and rooms in the Rhino Scene.
> If no room geometry is input, wil create a new room for each TFA surface with a height of 2.5m
> Note that a 'Room' can have more than one floor area / volume / space (closets in bedrooms, for instance).
> This component will need to be able to read TFA and Room Name/Number data from the Rhino Scene (User-Text 'Object Name', 'Room_Number', 'TFA_Factor').
-
EM July 31, 2020

    Args:
        _roomTFASurfaces: (List) An input for the user-determined room floor area(s) to use.
        Supply inputs as surfaces. Use the Rhino tools to apply TFA information (TFA Factor, Roomname, etc..)
        in the Rhino scene before passing in here. Will bundle together rooms with the same number+name and
        will try and join together tfa surfaces that touch (if they have the same number+name).
        _roomGeometry: (List) <Optional> An input for user-determined room geometry describing the space-shape of the rooms. 
        If no inputs are supplied, the component will just build a room with a standard ceiling height of 2.5m (8.2ft).
        If supplied, best-practice is to join together each room's surfaces (walls+ceilings) into a single poly-surface (1 per room). Be sure
        to leave off the TFA surface (leave 'open' on the bottom) so that the TFA srfc can be joined to it in the component here.
        _HBZones: (List) A list of all the HB Zones to use.
    Returns:
        roomBreps_: A preview of all the PHPP-Room Breps built by this component
        rooms_: A list of all the PHPP-Room Objects built
        HBZones_: the HB Zone Objects with the new PHPP-Rooms added. (Added as attribute 'PHPProoms' - a list of all the Room Objects)
"""

ghenv.Component.Name = "BT_PHPProomsFromRH"
ghenv.Component.NickName = "PHPP Rooms from Rhino"
ghenv.Component.Message = 'JUL_31_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import rhinoscriptsyntax as rs
import Grasshopper.Kernel as ghK
import ghpythonlib.components as ghc
import scriptcontext as sc
import Rhino
from contextlib import contextmanager
import copy
from collections import defaultdict

# Classes and Defs
preview=sc.sticky['Preview']
PHPP_TFA_Surface = sc.sticky['PHPP_TFA_Surface']
PHPP_Room = sc.sticky['PHPP_Room']
PHPP_RoomVolume = sc.sticky['PHPP_RoomVolume']
PHPP_Sys_Ventilation = sc.sticky['PHPP_Sys_Ventilation']

hb_hive = sc.sticky["honeybee_Hive"]()
HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)

@contextmanager
def rhDoc():
    """For reaching into the Rhino document
    to get params for Objects"""
    
    try:
        sc.doc = Rhino.RhinoDoc.ActiveDoc
        yield
    finally:
        sc.doc = ghdoc

def joinTouchingTFAsurfaces(_tfaSrfcObjs, _HBzoneObjs):
    # Takes in a set of TFA Surface Objects that are touching
    # Returns a new single TFA Surface Obj with averaged / joined values
    srfcExtPerims = []
    AreaGross = []
    TFAs = []
    
    # Figure out the new joined room's Ventilation Flow Rates
    ventFlowRates_Sup = []
    ventFlowRates_Eta = []
    ventFlowRates_Tran = []
    # Get all the already input flow rates for the TFA Surfaces (if any)
    for srfcObj in _tfaSrfcObjs:
        ventFlowRates_Sup.append(srfcObj.V_sup)
        ventFlowRates_Eta.append(srfcObj.V_eta)
        ventFlowRates_Tran.append(srfcObj.V_trans)
    
    # Use the max values from the input set as the new Unioned objs' vent flow rates
    unionedSrfcVentRates = [max(ventFlowRates_Sup), max(ventFlowRates_Eta), max(ventFlowRates_Tran) ]
    
    for srfcObj in _tfaSrfcObjs:
        TFAs.append( srfcObj.getArea_TFA() )
        AreaGross.append(srfcObj.Area_Gross)
        srfcExtEdges = ghc.BrepEdges( rs.coercebrep(srfcObj.Surface) )[0]
        srfcExtPerims.append( ghc.JoinCurves( srfcExtEdges, preserve=False ) )
    
    unionedSrfc = ghc.RegionUnion(srfcExtPerims)
    unionedTFAObj = PHPP_TFA_Surface(unionedSrfc, _HBzoneObjs, unionedSrfcVentRates, _inset=0, _offsetZ=0)
    
    # Set the TFA Surface Attributes
    unionedTFAObj.Area_Gross = sum(AreaGross)
    unionedTFAObj.TFAfactor = sum(TFAs) / sum(AreaGross)
    unionedTFAObj.RoomNumber = _tfaSrfcObjs[0].RoomNumber
    unionedTFAObj.RoomName = _tfaSrfcObjs[0].RoomName
    
    return unionedTFAObj

def createTFASurfaces(_tfaSrfcs, _HBZoneObjects):
    """Creates TFA Surface Objects from input with all their parameter values"""
    
    tfaSurfacObjs = []
    tfaSurfacObjsDict = {}
    
    for srfc in _tfaSrfcs:
        # Create a new TFA Surface Object for the geometry item (Curve, Surface) input
        try:
            newTFASurfaceObj = PHPP_TFA_Surface(srfc, _HBZoneObjects, _roomVentFlowRates='Automatic', _inset=0, _offsetZ=0)
            tfaSurfacObjs.append(newTFASurfaceObj)
            
            # Add to the dictionary
            key = "{}-{}".format(newTFASurfaceObj.RoomNumber, newTFASurfaceObj.RoomName)
            if key not in tfaSurfacObjsDict.keys():
                tfaSurfacObjsDict[key] = [newTFASurfaceObj]
            else:
                tfaSurfacObjsDict[key].append(newTFASurfaceObj)
        except:
            errorMsg = "Something went wrong building the TFA Floor Surfaces. Are you sure you applied TFA and Room Name info for all the surfaces?"
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, errorMsg)
        
    return tfaSurfacObjsDict

def findNeighbors(_srfcList):
    """ Takes in a list of surfaces. tests against others in the
    set to see if they are touching. Adds a 'Neighbor' marker if so"""
    
    for i, srfc in enumerate(_srfcList):
        
        for k in range(0, len(_srfcList)):
            #print 'Surface: {} looking at Surface: {}...'.format(i, k)
            vertsI =  ghc.DeconstructBrep( _srfcList[i].Surface ).vertices
            vertsK =  ghc.DeconstructBrep( _srfcList[k].Surface ).vertices
            
            if ghc.BrepXBrep(_srfcList[i].Surface, _srfcList[k].Surface).curves:
                if _srfcList[i].Neighbors != None:
                    matchSetID = _srfcList[i].Neighbors
                elif _srfcList[k].Neighbors != None:
                    matchSetID = _srfcList[k].Neighbors
                else:
                    matchSetID = i
                
                _srfcList[i].addNeighbor(matchSetID) #branch[k].ID)
                _srfcList[k].addNeighbor(matchSetID) #branch[i].ID)
    
    return _srfcList

def binByNeighbor(_srfcList):
    """Creates a dict of surfaces, using the 'Neighbor' as key. So
    will return a dict with lists of surfaces with multiple values if
    they are touching"""
    
    srfcSets = {}
    
    for each in _srfcList:
        if each.Neighbors == None:
            srfcSets[each.ID] = [each]
        else:
            if each.Neighbors not in srfcSets.keys():
                srfcSets[each.Neighbors] = [each]
            else:
                srfcSets[each.Neighbors].append(each)
    
    return srfcSets

roomGeomBreps = []
roomBreps_ = []
tfaSrfcBreps = []
rooms_ = []
tfaSrfcObjs_Unioned = []
tfaSrfcObjs = {}
roomsDict = defaultdict()

#------------------------------------------------------------------------------
# Get the Brep Geometry from Rhino Scene
# Get all the Room Geomerty Breps, TFA Geometry Breps, TFA Surface Objects
with rhDoc():
    for i, brepGUID in enumerate(_roomGeometry): roomGeomBreps.append(rs.coercebrep(brepGUID) )
    for i, brepGUID in enumerate(_roomTFASurfaces): tfaSrfcBreps.append(rs.coercebrep(brepGUID) )
    if len(_roomTFASurfaces)>0 and len(_HBZones)>0: tfaSrfcObjs = createTFASurfaces(_roomTFASurfaces, HBZoneObjects)

#------------------------------------------------------------------------------
# Build a Default room if nothing is passed in
if len(_HBZones)>0 and len(_roomTFASurfaces)==0:
    # Create the TFA surface from the floor
    for i in range(len(HBZoneObjects)):
        for surface in HBZoneObjects[i].surfaces:
            if surface.type >= 2 and surface.type <3: # 2-3 are the Floor surface types
                # Get the Vent flow rates, if any
                try:
                    ventFlowRates = roomVentFlowRates_.Branch(0)
                except:
                    ventFlowRates = ['Automatic']
                try:
                    newTFASurfaceObj = PHPP_TFA_Surface(surface.geometry, HBZoneObjects, ventFlowRates, _inset=0.1, _offsetZ=0.1)
                except:
                    errorMsg = "Something went wrong building the TFA Floor Surfaces."\
                    "Are you sure you applied TFA and Room Name info for all the surfaces?"
                    ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, errorMsg)
                tfaSrfcObjs = {'000-DEFAULT' : [newTFASurfaceObj] }

#------------------------------------------------------------------------------
# Join Together any touching TFA Surfaces
for k, tfaSrfcObjList in tfaSrfcObjs.items():
    # Join TFA surfaces if needed
    if len(tfaSrfcObjList)>1:
        srfcSets = findNeighbors(tfaSrfcObjList)
        srfcSets_Joined = binByNeighbor(srfcSets)
        for k, each in srfcSets_Joined.items():
            if len(each)>1:
                joinedSrfc = joinTouchingTFAsurfaces(each, HBZoneObjects)
                tfaSrfcObjs_Unioned.append( joinedSrfc )
            else:
                tfaSrfcObjs_Unioned.append( each[0] )
    else:
        tfaSrfcObjs_Unioned.append( tfaSrfcObjList[0] )

#------------------------------------------------------------------------------
# Build the Rooms for each TFA Surface
# Sort the rooms into a dict based on RoomNumber and RoomName

for tfaSrfcObj in tfaSrfcObjs_Unioned:
    if tfaSrfcObj.HostError == True:
        roomName = getattr(tfaSrfcObj, 'RoomName', 'No Room Name')
        roomNum = getattr(tfaSrfcObj, 'RoomNumber', 'No Room Number')
        msg = "Couldn't figure out which zone the room '{}-{}' should "\
        "go in?\nMake sure the room is completely inside one or another zone.".format(roomNum, roomName)
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
        continue
    
    # If Room Geometry passed in, do:
    # See if you can make a closed room Brep
    # If you can, create a new Room Volume from the closed Brep set
    # Then remove that Room's Geom from the set to test (to speed future search?)
    # If no closables match found, create a default room volume
    # If no room geom input, just build a default size room for each
    if len(roomGeomBreps)>0:
        for i, roomGeom in enumerate(roomGeomBreps):
            tfaFloorJoinedToGeom = ghc.BrepJoin([roomGeom, tfaSrfcObj.Surface])
            if tfaFloorJoinedToGeom.closed==True:
                newRmVol = PHPP_RoomVolume(tfaSrfcObj, roomGeom)
                roomGeomBreps.pop(i)
                break
        else:
            newRmVol = PHPP_RoomVolume(tfaSrfcObj, _roomGeom=None, _roomHeightUD=2.5)
            roomName = getattr(tfaSrfcObj, 'RoomName', 'No Room Name')
            roomNum = getattr(tfaSrfcObj, 'RoomNumber', 'No Room Number')
            msg = "I could not join the room TFA surface and any room\n"\
            "geometry together to make a closed Brep for room: '{}-{}'.\n"\
            "Please ensure that the geometry can be joined and try again.".format(roomNum, roomName)
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Remark, msg)
    else:
        newRmVol = PHPP_RoomVolume(tfaSrfcObj, _roomGeom=None, _roomHeightUD=2.5)
    
    # Sort the new Room Volume into the master Dict based on name
    if tfaSrfcObj.RoomNumber != 'None' and tfaSrfcObj.RoomName.upper() != 'None':
        key = str(tfaSrfcObj.RoomNumber) + '_' + str(tfaSrfcObj.RoomName).upper()
    else:
        key = len(roomsDict.keys())
    
    if key in roomsDict.keys():
        roomsDict[key].append( newRmVol )
    else:
        roomsDict[key] = [ newRmVol ]

#------------------------------------------------------------------------------
# Build the final Rooms from all the the Room Volume Objects
for roomVolumes in roomsDict.values():
    # Create a single new Room from the list of Room Volumes
    newRoomObj = PHPP_Room( roomVolumes)
    
    # Preview Outputs
    rooms_.append(newRoomObj)
    for eachRoomBrep in newRoomObj.RoomBreps:
        roomBreps_.append(eachRoomBrep)

#------------------------------------------------------------------------------
# Add a list of the Room Objects to the Zone as an attribute
if len(_HBZones)>0:
    # Add the new Rooms to the right host Honeybee Zone
    for zone in HBZoneObjects:
        # Add a Default Vent System to the Zone
        defaultVentSystem = PHPP_Sys_Ventilation()
        setattr(zone, 'PHPP_VentSys', defaultVentSystem)
        
        zoneRooms = []
        for room in rooms_:
            if room.HostZoneName == zone.name:
                # Assign the Default Vent Unit for the Zone's rooms
                setattr(room, 'VentUnitName', defaultVentSystem.Unit_Name)
                setattr(room, 'VentSystemName', defaultVentSystem.SystemName)
                zoneRooms.append(room)
        setattr(zone, 'PHPProoms', zoneRooms)
    
    # Add the modified Honeybee zone objects back to the HB dictionary
    HBZones_  = hb_hive.addToHoneybeeHive(HBZoneObjects, ghenv.Component)
