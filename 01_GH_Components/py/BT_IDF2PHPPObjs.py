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
Takes in the IDF 'Objects' from the reader and organizes them for export to the PHPP. Gets all the relevant Materials, Constructions and Surfaces from the IDF file.
-
EM September 29, 2020

    Args:
        _HBZones: <Optional> If connected, the component will try and read detailed 'Frame' and 'Glass' Object data for each window in building. If this isn't hooked up, the normal EP windows will be used to create PHPP-Style Window Components. 
        _IDF_Objs_List: Takes in a list if IDF objects. Connect to the 'IDF_Objs_List' output on the 'IDF Reader' Component
    Returns:
        opaqueSurfaces: A List of the opaque surface IDF Objects found 
        windowObjects: A List of the window surface IDF Objects found
        zoneNames_: A List of the Zone names found
        PHPPObjs_: DataTree of all the organized, setup PHPP objects to pass to the writer
"""

ghenv.Component.Name = "BT_IDF2PHPPObjs"
ghenv.Component.NickName = "IDF-->PHPP Objs"
ghenv.Component.Message = 'SEP_29_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "02 | IDF2PHPP"

from System import Object
from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path
import Grasshopper.Kernel as ghK
import scriptcontext as sc
from collections import namedtuple
import ghpythonlib.components as ghc
import math
import copy
from collections import defaultdict
import rhinoscriptsyntax as rs
from collections import namedtuple

# Classes and Defs
preview=sc.sticky['Preview']
phpp_calcNorthAngle=sc.sticky['phpp_calcNorthAngle']
phpp_GetWindowSize=sc.sticky['phpp_GetWindowSize']
phpp_makeHBMaterial=sc.sticky['phpp_makeHBMaterial']
phpp_makeHBConstruction=sc.sticky['phpp_makeHBConstruction']

phpp_ClimateData = sc.sticky['phpp_ClimateData']

PHPP_XL_Obj=sc.sticky['PHPP_XL_Obj']
PHPP_Glazing=sc.sticky['PHPP_Glazing']
PHPP_Frame=sc.sticky['PHPP_Frame']
PHPP_Window_Install=sc.sticky['PHPP_Window_Install']
PHPP_ClimateDataSet = sc.sticky['PHPP_ClimateDataSet']

IDF_Zone = sc.sticky['IDF_Zone']
IDF_ZoneInfilFlowRate = sc.sticky['IDF_ZoneInfilFlowRate']
IDF_ZoneList = sc.sticky['IDF_ZoneList']
IDF_Obj_building=sc.sticky['IDF_Obj_building']
IDF_Obj_MaterialLayer=sc.sticky['IDF_Obj_MaterialLayer']
IDF_Obj_MaterialWindowSimple=sc.sticky['IDF_Obj_MaterialWindowSimple']
IDF_Obj_MaterialWindowGlazing = sc.sticky['IDF_Obj_MaterialWindowGlazing']
IDF_Obj_MaterialWindowGas = sc.sticky['IDF_Obj_MaterialWindowGas']
IDF_Obj_Construction=sc.sticky['IDF_Obj_Construction']
IDF_Obj_surfaceWindow=sc.sticky['IDF_Obj_surfaceWindow']
IDF_Obj_surfaceOpaque=sc.sticky['IDF_Obj_surfaceOpaque']
IDF_Obj_location = sc.sticky['IDF_Obj_location']

hb_hive = sc.sticky["honeybee_Hive"]()
HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)

def parseIDFObjects(_IDF_Objs):
    # Looks at the IDF Objects and parses them  out
    # Builds class objects as appropriate
    zones = []
    zoneInfiltrationRates = []
    zonesList = []
    opaqueMaterials = []
    windowMaterialsSimple = {}
    windowMaterialGas = {}
    windowMaterialGlazing = {}
    allConstructions = []
    opaqueSurfaces = []
    location = []
    
    # First, need to find the North Direction. Have to do that before the rest
    for each in _IDF_Objs:
    # Look at each of the IDF objects coming in....
        try:
            idfObjName = getattr(each, 'objName')
        except:
            idfObjName = ''
        
        if idfObjName == 'Building':
            # If the IDF Obj is a 'Building'
            # Create the Building Object and get the Project's North Angle Vector
            bldg = IDF_Obj_building(each)
            bldgNorthVec = bldg.NorthVector
    
    # Now go through and pull out each class object
    for eachIDFobj in _IDF_Objs:
        try:
            idfObjName = getattr(eachIDFobj, 'objName')
        except:
            idfObjName = ''
        
        # If its an opaque Building Surface object
        if 'BuildingSurface:Detailed' in idfObjName:
            opaqueSurfaces.append(  IDF_Obj_surfaceOpaque(eachIDFobj, bldgNorthVec)  )
        
        # If its a 'Material' or 'Material:AirGap' object
        elif idfObjName == 'Material:AirGap' or idfObjName == 'Material':
            opaqueMaterials.append( IDF_Obj_MaterialLayer(eachIDFobj) )
        
        elif idfObjName == 'Material:NoMass':
            opaqueMaterials.append( IDF_Obj_MaterialLayer(eachIDFobj, noMass=True) )
        
        # If its a simple EP Style Window Material
        elif 'WindowMaterial:SimpleGlazingSystem' in idfObjName:
            windowMaterialsSimple[eachIDFobj.Name] = IDF_Obj_MaterialWindowSimple(eachIDFobj)
        
        elif 'WindowMaterial:Gas' in idfObjName:
            windowMaterialGas[eachIDFobj.Name]  = IDF_Obj_MaterialWindowGas( eachIDFobj )
        
        elif 'WindowMaterial:Glazing' in idfObjName:
            windowMaterialGlazing[eachIDFobj.Name]  = IDF_Obj_MaterialWindowGlazing( eachIDFobj )
        
        elif 'Construction' in idfObjName:
            allConstructions.append( IDF_Obj_Construction( eachIDFobj )  )
        
        elif idfObjName == 'Zone':
            newObj = IDF_Zone( eachIDFobj )
            zones.append( IDF_Zone( eachIDFobj ) )
        
        elif idfObjName == 'ZoneList':
            zonesList.append( IDF_ZoneList( eachIDFobj ) )
            
        elif 'ZoneInfiltration:DesignFlowRate' in idfObjName:
            zoneInfiltrationRates.append( IDF_ZoneInfilFlowRate( eachIDFobj  ) )
        
        elif 'Site:Location' in idfObjName:
            location = eachIDFobj
            
    return opaqueSurfaces, opaqueMaterials, windowMaterialsSimple, windowMaterialGas, windowMaterialGlazing, allConstructions, zones, zoneInfiltrationRates, zonesList, location

def materialWindowSimpleFromLayers(_const):
    # If its a Window 'construction' of multiple layers, calc an approximate effective Uw
    # https://bigladdersoftware.com/epx/docs/8-5/engineering-reference/window-heat-balance-calculation.html#equivalent-layer-thermal-model
    # NOTE: Not doing this correctly right now. Neglecting radiation of convective effects. Only a super simplified approximiation for now....
    
    class temp:
        # temp holder for params to pass
        def __init__(self):
            pass
    
    # Ok... so its a window with multiple layers. So need to calc equiv conductivity
    newWinUg = [0.04, 0.13] # Start with the Surface Film Resistances...
    for eachLayer in _const.Layers:
        if eachLayer[1] in windowMaterialGlazing:
            newWinUg.append( 1 / windowMaterialGlazing[ eachLayer[1] ].uValue ) # Resistance of Glass Layers
        elif eachLayer[1] in windowMaterialGas:
            newWinUg.append( 1 / windowMaterialGas[ eachLayer[1] ].uValue * 0.5 ) # Resistance of Gas Layers
    newWinUg = 1/sum(newWinUg)
    
    # Create a New WindowMaterial:SimpleGlazingSystem Object to approximate this built-up construction
    tempObj = temp()
    setattr(tempObj, 'U-Factor {W/m2-K}', newWinUg)
    setattr(tempObj, 'Solar Heat Gain Coefficient', 0.4)
    setattr(tempObj, 'Visible Transmittance', 0.75)
    setattr(tempObj, 'Name', _const.Name)
    
    return IDF_Obj_MaterialWindowSimple( tempObj )

def filterConstructions(_allConst, _materialsWindowSimple, _windowMaterialGas, _windowMaterialGlazing):
    # Takes in all the construction and splits them into 
    # Window and Opaque constructions
    
    opaqueConstructions = []
    windowConstructionsSimple = {}
    
    for construction in allConstructions:
        isWindow = False
        
        # Is it a Window Construction?
        # If the Construction includes any materials found in the WindowMaterialsSimple, WindowGas or WindowGlazing, then yes
        for eachMat in _materialsWindowSimple.keys():
            if _materialsWindowSimple[eachMat].Name in construction.LayerNames:
                 isWindow = True
        for eachMat in _windowMaterialGas:
            if eachMat in construction.LayerNames:
                 isWindow = True
        for eachMat in _windowMaterialGlazing:
            if eachMat in construction.LayerNames:
                 isWindow = True
        
        # Now branch off the construction as appropriate
        if isWindow==True:
            # Its a window, so add to the Window Construction list
            
            if len(construction.Layers) > 1:
                # Its a built up window. So turn that into a Simple Window
                # Modify window construction using approximation and create a 'Simple' Window
                materialWindowSimple = materialWindowSimpleFromLayers( construction )  # Create a new MaterialWindowSimple
                _materialsWindowSimple[construction.Name] = materialWindowSimple       # Add the new MaterialWindowSimple to the list
                construction.Layers = [ ['Layer1', materialWindowSimple.Name]  ]       # Change the Construction Layers to ONLY include the MaterialWindowSimple now  
                windowConstructionsSimple[construction.Name] = construction            # Add the modified window construction to the list
            else:
                windowConstructionsSimple[construction.Name] =  construction
        else:
            # its not a window, so add to the Opaque Construction List
            opaqueConstructions.append( construction )
    
    return opaqueConstructions, windowConstructionsSimple, _materialsWindowSimple

def getPHPPRooms(_HBZoneObjects):
    # Looks at for all the rooms and ventilation systems
    
    HBZonePHPPRooms=[]
    HBZoneVentSystemsDict = {}
    if len(_HBZones)>0:
        # Pull out the rooms if there are any
        for zone in _HBZoneObjects:
            # Look at each HB Zone and see if there have been PHPP rooms added
            if 'PHPProoms' in zone.__dict__.keys():
                # Some PHPP Rooms there. Read them.
                for eachRoom in zone.PHPProoms:
                    HBZonePHPPRooms.append(eachRoom)
        
        # Get the Ventilation Systems from the HB Zones if there are any
        for zone in _HBZoneObjects:
            if 'PHPP_VentSys' in zone.__dict__.keys():
                # There is a Ventilation System applied. Pull it out
                HBZoneVentSystemsDict[zone.PHPP_VentSys.Unit_Name] = zone.PHPP_VentSys
        HBZoneVentSystems = [HBZoneVentSystemsDict]
    else:
        HBZoneVentSystems = []
        
    return HBZonePHPPRooms, HBZoneVentSystems

def getIDFWindowObjects(_IDF_Objs, _windowConstructionsSimple, _windowMaterialsSimple):
    # Finds all  the widnow surfaces and builds window objects
    windowSurfaces = []
    windowObjs_raw = []
    windowObjs_filtered = []
    windowObjs_triangulated = {}
    
    for eachIDFobj in _IDF_Objs:
        try:
            idfObjName = getattr(eachIDFobj, 'objName')
            
        except:
            idfObjName = ''
        
        # If its an EP Window Object
        if 'FenestrationSurface:Detailed' in idfObjName:
            windowObjs_raw.append(eachIDFobj)
    
    ##################################################
    # Fix for window triangulation
    for windowObj in windowObjs_raw:
        # Honeybee adds the code '..._glzP_0, ..._glzP_1, etc..' suffix to the name for its triangulated windows
        if '_glzP_' in windowObj.Name:
            # See if it has only 3 vertices as well just to double check
            numOfVerts = 0
            for key in windowObj.__dict__.keys():
                if 'XYZ Vertex' in key:
                    numOfVerts += 1
            if numOfVerts == 3:
                # Ok, so its a triangulated window.
                # File the triangulated window in the dictionary using its name as key
                
                tempWindowName = windowObj.Name.split('_glzP_')[0]
                if tempWindowName not in windowObjs_triangulated.keys():
                    windowObjs_triangulated[ tempWindowName ] = [windowObj]
                else:
                    windowObjs_triangulated[ tempWindowName ].append(windowObj)
            else:
                windowObjs_filtered.append(windowObj)
        else:
            windowObjs_filtered.append(windowObj)

    # Unite the triangulated objects
    for key in windowObjs_triangulated.keys():
        perims = []
        for windowObj in windowObjs_triangulated[key]:
            triangleVerts = []
            # Get the verts
            for key in windowObj.__dict__.keys():
                if 'XYZ Vertex' in key:
                    verts = getattr(windowObj, key).split(' ')
                    verts = [float(x) for x in verts]
                    point = ghc.ConstructPoint(verts[0], verts[1], verts[2])
                    triangleVerts.append( point )
            
            # Union the Segments, find the outside perimeter
            perim = ghc.PolyLine(triangleVerts, closed=True)
            perims.append(perim)
        
        unionedPerim = ghc.RegionUnion(perims)
        
        # Build a new Window Obj using this now unioned geometry
        newVertPoints = ghc.ControlPoints(unionedPerim).points #windowObj.__dict__
        newWindowObj = copy.deepcopy(windowObj)
        
        for i in range(len(newVertPoints)):
            setattr(newWindowObj, 'XYZ Vertex {} {}'.format(i+1, '{m}'), str(newVertPoints[i]).replace(',',' '))
            setattr(newWindowObj, 'Name', windowObj.Name[:-7])
        
        windowObjs_filtered.append(newWindowObj)
    
    ##################################################
    
    # Build the Window Objects
    for eachWindowObj in windowObjs_filtered:
            # Find the windows's CONSTRUCTION and MATERIAL information in the IDF
            thisWindowEP_CONST_Name = getattr(eachWindowObj, 'Construction Name')                         # Get the name of the Windows' Construction  
            thisWindowEP_MAT_Name = _windowConstructionsSimple[  thisWindowEP_CONST_Name  ].Layers[0][1]  # Find the Material name of 'Layer 1' in the Window's Construction
            thisWindowEP_WinSimp_Obj = _windowMaterialsSimple[  thisWindowEP_MAT_Name  ]                  # Find the 'WindowMaterial:SimpleGlazingSystem' Object with the same name as 'Layer 1' 
            winterShadingFactor = 0.75
            summerShadingFactor = 0.75
            
            # Create the new IDF_Obj_surfaceWindow Object
            windowSurfaces.append( IDF_Obj_surfaceWindow(eachWindowObj, thisWindowEP_WinSimp_Obj, winterShadingFactor, summerShadingFactor) )
    
    return windowSurfaces

def updatePHPPStyleWindows(_zoneObjs, _IDFwindowSurfaces):
    # Used to update / overwrite the IDF window Params with the more detailed
    # Params from the HB Zone 'phppWindowDict' <if they exist>
    
    for IDFwindowObj in _IDFwindowSurfaces:
        # Find the IDFWindow's host zone
        for zone in _zoneObjs:
            for surface in zone.surfaces:
                if surface.name in IDFwindowObj.HostSrfc:
                    # Found the window's host zone
                    try:
                        #print 'Updating IDFWindow Object: <{}> with Params from HB Zone'.format(IDFwindowObj.Name)
                        # Get the HB Zone's detailed PHPP Style Window Data
                        phppWindowObj = zone.phppWindowDict[ IDFwindowObj.Name ]
                        
                        # Re-set the IDF-Window Obj's param data with the detailed HB Data
                        setattr(IDFwindowObj, 'Type_Variant', phppWindowObj.Type_Variant)
                        setattr(IDFwindowObj, 'Type_Frame', phppWindowObj.Type_Frame)
                        setattr(IDFwindowObj, 'Type_Glass', phppWindowObj.Type_Glass)
                        setattr(IDFwindowObj, 'Installs', phppWindowObj.Installs)
                        
                        shadingFactors = phppWindowObj.getShadingFactors_Simple()
                        setattr(IDFwindowObj, 'winterShadingFac', shadingFactors[0])
                        setattr(IDFwindowObj, 'summerShadingFac', shadingFactors[1])
                    except:
                        pass
                    break

def filterSurfaces(_surfaces):
    # Filter to only include the surface if its 'exposed' to the outdoors or Ground (not an interior floor / wall)
    exposedSurfaces = []
    
    for srfc in _surfaces:
        if srfc.exposure != 'Surface':
        #if srfc.exposure == 'Outdoors' or srfc.exposure == 'Ground':
            exposedSurfaces.append(srfc)
        
    return exposedSurfaces

def buildZoneBrep(_zoneObjs, _opaqueSurfaces):
    # Takes in the IDF Surfaces and builds Zone Breps from them
    # Sets the ZoneObj as an attr using the new Brep
    
    zoneBreps = []
    
    for zone in _zoneObjs:    
        zoneSurfaces = []
        for srfc in _opaqueSurfaces:
            if srfc.HostZoneName == zone.ZoneName:
                zoneSurfaces.append( ghc.BoundarySurfaces(srfc.Boundary) )
        zoneBrep = ghc.BrepJoin( zoneSurfaces ).breps
        zoneBreps.append( zoneBrep )
        setattr(zone, 'ZoneBrep', zoneBrep)
    
    return zoneBreps

def calcZoneParams(_zonesList, _zoneObjs, _opaqueSurfaces, _rooms, _HBZoneObjs):
    # Looks at the ZoneList, and for each Zone in the List
    # Finds the Floor Area, the exposed Surface Area (Outdoors) and the Zone Volume
    # And the right Infiltration Rate. Calcs the total Infiltration Rate in m3/h for each Zone
    # Sets Zone Attributes for ACH, Volume, Floor Area
    
    blowerPressure = 50 #Pa - for calc'ing the tested/input ACH
    
    # Get the Infiltration Rate, Volume, Floor Area from the IDF File
    for zoneList in _zonesList:
        for zoneInfilObj in zoneInfiltrationRates:
            if zoneInfilObj.ZoneName == zoneList.Name:
                # Use this InfilObj's params...
                # So now for each actual ZONE name in the Zone List...
                for eachKey in zoneList.__dict__.keys():
                    if eachKey != 'Name': # Cus' this key is only the name of the ZoneList Object
                        zoneNameinList = getattr(zoneList, eachKey)
                        for zoneObj in _zoneObjs:
                            if zoneObj.ZoneName == zoneNameinList:
                                # Find the HB Zone Obj
                                for HBzone in _HBZoneObjs:
                                    if HBzone.name == zoneObj.ZoneName:
                                        thisHBZone = HBzone
                                
                                # Ok, found the right Zone Obj. Get the Reference Values
                                # Find the PHPProom Volumes for the zone
                                roomVn50s = []
                                for room in _rooms:
                                    if room.HostZoneName == zoneObj.ZoneName:
                                        roomVn50s.append(room.RoomNetClearVolume)
                                
                                # Get the Gross Exposed Surface Area of the Zone
                                zoneExposedFacadeArea = thisHBZone.getExposedArea()
                                
                                # The Gross Reference Floor area to use for calcs
                                zoneRefFloorArea = thisHBZone.getFloorArea()
                                
                                # Zone Volume, but use Room Vn50s if there are any
                                zoneRefVolume = thisHBZone.getZoneVolume()
                                if len(roomVn50s) != 0:
                                    zoneRefVolume = sum(roomVn50s)
                                
                                # Calc the Infiltration Rate by Floor Area or by Facade Area
                                try:
                                    infilRatebyFloor = float( getattr(zoneInfilObj, 'FlowRatePerFloorArea') )# m3/s-m2
                                    infilRatebyFacadeArea = 0
                                    zoneInfilRate = infilRatebyFloor * zoneRefFloorArea * 60 * 60 # sec/min * min/hour
                                    zoneACH50 = ((math.pow((blowerPressure/4),0.63)) * zoneInfilRate ) / zoneRefVolume
                                except:
                                    try:
                                        infilRatebyFloor = 0
                                        infilRatebyFacadeArea = float( getattr(zoneInfilObj, 'FlowRatePerSurfaceArea') )# m3/s-m2
                                        zoneInfilRate = infilRatebyFacadeArea * zoneExposedFacadeArea * 60 * 60
                                        zoneACH50 = ((math.pow((blowerPressure/4),0.63)) * zoneInfilRate ) / zoneRefVolume
                                    except:
                                        print 'something went wrong. Maybe no Infiltration rate applied?'
                                
                                # Set the attributes on the Zone Objects
                                setattr(zoneObj, 'InfiltrationACH50', zoneACH50)
                                setattr(zoneObj, 'Volume_Gross', zoneRefVolume)
                                setattr(zoneObj, 'FloorArea_Gross', zoneRefFloorArea)
                                setattr(zoneObj, 'Volume_Vn50', zoneRefVolume)
                                setattr(zoneObj, 'TFA', zoneRefFloorArea)
                                
    
    # Try and get the room Vn50 and TFA from the PHPP_Roooms in the HB Zone
    for zone in _zoneObjs:
        zoneTFAs = []
        zoneVn50s = []
        
        for room in _rooms:
            if room.HostZoneName == zone.ZoneName:
                zoneTFAs.append( room.FloorArea_TFA )
                zoneVn50s.append( room.RoomNetClearVolume )
        
        if zoneTFAs and zoneVn50s:
            setattr(zone, 'TFA', sum(zoneTFAs) )
            setattr(zone, 'Volume_Vn50', sum(zoneVn50s) )

def getDHWSys(_zoneObjs):
    dhwSystems = defaultdict()
    
    for zone in _zoneObjs:
        dhwSystemObj = getattr(zone, 'PHPP_DHWSys', None)
        if dhwSystemObj != None:
            dhwSystems[dhwSystemObj.SystemName] = dhwSystemObj
    
    if len(dhwSystems.keys())>1:
        warning = 'Found more than one type of DHW System applied. I will try and combine them\n'\
        'together, but you might be better off exporting the different zones to separate PHPPs\n'\
        'if you really have multiple different DHW systems in the building?'
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
    
    return list(dhwSystems.values())

def getGround(_zoneObjs):
    # Pull out the ground objects from the Honeybee Zones
    groundObjs = []
    
    for zone in _zoneObjs:
        if hasattr(zone, 'PHPP_ground'):
            groundObjs.append( getattr(zone, 'PHPP_ground') )
    
    return groundObjs

def findNearestPHPPclimateZone(_lat, _long, _climateData):
    """ Finds the nearest PHPP Climate zone to the EPW Lat /Long 
    
    Methodology copied from the PHPP v 9.6a (SI) Climate worksheet
    """
    def sortByDist(e):
        return e['distToEPW']
    
    for each in _climateData:
        eachLat = float(each.get('Latitude', 0))
        eachLong = float(each.get('Longitude', 0))
        
        a = math.sin(math.pi/180*eachLat)*math.sin(math.pi/180*_lat)+math.cos(math.pi/180*eachLat)*math.cos(math.pi/180*_lat)*math.cos(math.pi/180*(eachLong-_long))
        b = max([-1, a])
        c = min([1, b])
        d = math.acos(c)
        kmFromEPWLocation = 6378 * d
        
        each['distToEPW'] = kmFromEPWLocation
    
    _climateData.sort(key=sortByDist)
    climateSetToUse = _climateData[0]
    
    dataSet = climateSetToUse.get('Dataset', 'US0055b-New York')
    alt = '=J23'
    country = climateSetToUse.get('Country', 'US-United States of America')
    region = climateSetToUse.get('Region', 'New York')
    
    climateSetToUse = PHPP_ClimateDataSet(dataSet, alt, country, region)
    
    return [climateSetToUse]

def get_appliances(_zones):
    appliances = []
    for zone in _zones:
        try:
            zone_appliances =  getattr(zone, 'PHPP_ElecEquip')
            for app in zone_appliances:
                setattr(app, 'ZoneFloorArea', zone.getFloorArea())
        except:
            zone_appliances = []
        
        appliances = appliances+zone_appliances
    
    return appliances

def get_phpp_lighting(_zones):
    lighting = []
    for zone in _zones:
        try:
            lightingObj = getattr(zone, 'PHPP_LightingEfficacy')
            setattr(lightingObj, 'ZoneFloorArea', zone.getFloorArea())
            lighting.append( lightingObj )
        except:
            pass
    
    return lighting

def calcFootprint(_zoneObjs, _opaqueSurfaces):
    # Finds the 'footprint' of the building for 'Primary Energy Renewable' reference
    # 1) Re-build the zone Breps
    # 2) Join all the zone Breps into a single brep
    # 3) Find the 'box' for the single joined brep
    # 4) Find the lowest Z points on the box, offset another 10 units 'down'
    # 5) Make a new Plane at this new location
    # 6) Projects the brep onto the new Plane
    
    #-----
    zoneBreps = []
    for zone in _zoneObjs:    
        zoneSurfaces = []
        for srfc in _opaqueSurfaces:
            if srfc.HostZoneName == zone.ZoneName:
                zoneSurfaces.append( ghc.BoundarySurfaces(srfc.Boundary) )
        zoneBrep = ghc.BrepJoin( zoneSurfaces ).breps
        zoneBreps.append( zoneBrep )
    
    bldg_mass = ghc.SolidUnion(zoneBreps)
    
    if bldg_mass == None:
        return None
    
    #------- Find Corners, Find 'bottom' (lowest Z)
    bldg_mass_corners = [v for v in ghc.BoxCorners(bldg_mass)]
    bldg_mass_corners.sort(reverse=False, key=lambda point3D: point3D.Z)
    rect_pts = bldg_mass_corners[0:3]
    
    #------- Project Brep to Footprint
    projection_plane1 = ghc.Plane3Pt(rect_pts[0], rect_pts[1], rect_pts[2])
    projection_plane2 = ghc.Move(projection_plane1, ghc.UnitZ(-10)).geometry
    matrix = rs.XformPlanarProjection(projection_plane2)
    footprint_srfc = rs.TransformObjects (bldg_mass, matrix, copy=True)
    footprint_area = rs.Area(footprint_srfc)
    
    #------- Output
    Footprint = namedtuple('Footprint', ['Footprint_surface', 'Footprint_area'])
    fp = Footprint(footprint_srfc, footprint_area)
    
    return fp


#-------------------------------------------------------------------------------
##### Read the IDF Objects and Build class objects  ##########

# Get Material Layers, Constructions, Surfaces
(opaqueSurfaces,
opaqueMaterials,
windowMaterialsSimple,
windowMaterialGas,
windowMaterialGlazing,
allConstructions,
zones,
zoneInfiltrationRates,
zonesList,
location) = parseIDFObjects(_IDF_Objs_List)

opaqueSurfaces_Exposed = filterSurfaces(opaqueSurfaces)

# Filter Constructions into Opaque/Window
(opaqueConstructions,
windowConstructionsSimple,
windowMaterialsSimple) = filterConstructions(allConstructions, windowMaterialsSimple, windowMaterialGas, windowMaterialGlazing)

# IDf Window Objects
windowObjects = getIDFWindowObjects(_IDF_Objs_List, windowConstructionsSimple, windowMaterialsSimple)

# Zone Rooms, Ventialtion from HB, Update windows to Detailed data from HB Zones
if len(_HBZones)>0 and len(_IDF_Objs_List)>1:
    HBZonePHPPRooms, HBZoneVentSystems = getPHPPRooms(HBZoneObjects)
    updatePHPPStyleWindows(HBZoneObjects, windowObjects)
    
    # Calc and  set Zone Attributes
    buildZoneBrep(zones, opaqueSurfaces)  # Build the Zone Breps and add to Zone Objects
    footprint = calcFootprint(zones, opaqueSurfaces)
    calcZoneParams(zonesList, zones, opaqueSurfaces, HBZonePHPPRooms, HBZoneObjects)   # Determine Infiltation and add to Zone Objects
    dhwSystemObj = getDHWSys(HBZoneObjects)
    groundObjs = getGround(HBZoneObjects)
    elec_equip_appliances = get_appliances(HBZoneObjects)
    phpp_lighting = get_phpp_lighting(HBZoneObjects)
else:
    HBZonePHPPRooms, HBZoneVentSystems = [], []
    dhwSystemObj = []
    groundObjs = []
    elec_equip_appliances = []
    phpp_lighting = []
    footprint = []

# Figure out the Closest PHPP Climate Zone
latitude = float(getattr(location, 'Latitude {deg}', 51.30))
longitude = float(getattr(location, 'Longitude {deg}', 9.44))
climate = findNearestPHPPclimateZone(latitude, longitude, phpp_ClimateData)
    
try:
    latitude = float(getattr(location, 'Latitude {deg}', 51.30))
    longitude = float(getattr(location, 'Longitude {deg}', 9.44))
    climate = findNearestPHPPclimateZone(latitude, longitude, phpp_ClimateData)
except:
    print 'Error finding the nearest PHPP Climate Zone?'
    climate = []

#-------------------------------------------------------------------------------
#########       Package up all the output objects    #########   
#########       into a single tree for passing       #########    

PHPPObjs_ = DataTree[Object]()
# Materials and Constructions
PHPPObjs_.AddRange(opaqueMaterials, GH_Path(0))                                 # Opaque Material Objects
PHPPObjs_.AddRange(opaqueConstructions, GH_Path(1))                             # Opaque IDF_Obj_Construction Objects
PHPPObjs_.AddRange(windowMaterialsSimple, GH_Path(2))                           # EP Style Window Material Objects
PHPPObjs_.AddRange(windowConstructionsSimple, GH_Path(3))                       # EP Style Window IDF_Obj_Construction Objects
# Surfaces
PHPPObjs_.AddRange(opaqueSurfaces_Exposed, GH_Path(4))                          # Opaque Surface Objects
PHPPObjs_.AddRange(windowObjects, GH_Path(5))                                   # PHPP Style Window IDF_Obj_Construction Objects
PHPPObjs_.AddRange(HBZonePHPPRooms, GH_Path(6))                                 # List of all the HB Zone PHPP Rooms
PHPPObjs_.AddRange(HBZoneVentSystems, GH_Path(7))                               # List the Fresh Air Ventilation System Dictionary
# Zones
PHPPObjs_.AddRange(zones, GH_Path(8))                                           # All of the Zones
PHPPObjs_.AddRange(zoneInfiltrationRates, GH_Path(9))                           # Zone Infiltration Airflow rates
PHPPObjs_.AddRange(dhwSystemObj, GH_Path(10))
PHPPObjs_.AddRange(groundObjs, GH_Path(11))
PHPPObjs_.AddRange(climate, GH_Path(12)) 
PHPPObjs_.AddRange(elec_equip_appliances, GH_Path(13)) 
PHPPObjs_.AddRange(phpp_lighting, GH_Path(14)) 
PHPPObjs_.Add(footprint, GH_Path(15)) 

#-------------------------------------------------------------------------------
# Output Preview of Zone Names
zoneNames_ = []
for zoneObj in PHPPObjs_.Branch(8):
    zoneNames_.append(zoneObj.ZoneName)