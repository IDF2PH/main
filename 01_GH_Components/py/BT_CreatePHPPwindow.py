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
Use this component BEFORE an HB 'addHBGlz' component to add PHPP-Style parameters to the window. 
Component will:
    - Create PHPP-Style Constructions from EP input or Default ('Exterior Window') with a U-W-Installed cacl'd for each window
    - Read window params from the Rhino Scene (object User-Text) and get data from the Rhino Window Library (Document User-Text)
    - Accept direct user-input of frame, glass and install objects
---------
Will work in that order, and later inputs will overwrite previous ones (direct user input overwrites Rhino scene params, which overwrites EP-Constructions)
Once it has gotten all the window params, builds a new PHPP-Style Window Object and writes to the master dictionary attached to the zone (creates this dict if it doesn't already exist)
Will lastly, write the NEW EP-Construction and EP-Material with the U-W-Installed value to the HB Library for use in the EP Simulation. 
-
EM July 25, 2020
    Args:
        _HBZones: (list) The HB Zone object(s) from HB Consttructors
        names_: (list) <Optional> An optional entry for user-defined Window Names to use. Input either a single name or a list matching the length of the geometry. 
        frames_: (list) <Optional> An optional entry for user-defined Frame Objects. Use the 'PHPP | Win. Frame' Component to create a Frame Object and pass in. Either pass in a single Frame object, or a list of Frames that matches the Geometry. If the length of the Frames list doesn't match the Geometry, the first Frame Object input will be used for all Windows.
        glazings_: (list) <Optional> An optional entry for user-defined Glass Objects. Use the 'PHPP | Win. Glass' Component to create a Glass Object and pass in. Either pass in a single Glass object, or a list of Frames that matches the Geometry. If the length of the Glass list doesn't match the Geometry, the first Glass Object input will be used for all Windows.
        psi_installs_: (list) <Optional> An optional entry for user-defined Psi-Install Values (W/m-k). Either pass in a single number which will be used for all edges, or a list of 4 numbers (left, right, bottom, top) - one for each edge.
        installs_:(list) <Optional> An optional entry for user-defined Install Conditions (1|0) for each window edge (1=Apply Psi-Install, 0=Don't apply Psi-Install). Either pass in a single number which will be used for all edges, or a list of 4 numbers (left, right, bottom, top) - one for each edge.
        EPConsts: (list) <Optional> A list of EP Construction Names to use for the windows. Accepts a single name which is then applied to all windows, or a list of names corresponding (length / order) to the windows. Be sure that the EP Constructions and EP Materials are already added to the HB Library.
        _windowGeom: (list) The Window Geometry to use. Accepts Surfaces or Curves from the Rhino scene. Pass all Rhino Geometry through an 'ID' component in order to have the component read the parameters from the Rhino Scene. Set the type hint to 'srfc' if you are trying to pass in geometry that you created in Grasshopper rather than Rhino geometry.
    Returns:
        HBZones_: The HBZone(s) with the new 'phppWindowDict' entries added.
        windowSurfaces_: A List of the Window Surfaces
        windowNames_: A List of the Window Names
        EPConstructions_: A List of the NEW EPConstructions to be applied to the Windows.
"""

ghenv.Component.Name = "BT_CreatePHPPwindow"
ghenv.Component.NickName = "New PHPP Window"
ghenv.Component.Message = 'JUL_25_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc
import rhinoscriptsyntax as rs
import Rhino
import Grasshopper.Kernel as ghK
import json
import ghpythonlib.components as ghc
import random
import re
from System import Object
from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path

# Defs and Classes
idf2ph_rhDoc = sc.sticky['idf2ph_rhDoc']
PHPP_Glazing = sc.sticky['PHPP_Glazing']
PHPP_Frame = sc.sticky['PHPP_Frame']
PHPP_Window_Install = sc.sticky['PHPP_Window_Install']
PHPP_WindowObject = sc.sticky['PHPP_WindowObject']
phpp_makeHBMaterial = sc.sticky['phpp_makeHBMaterial']
phpp_makeHBConstruction = sc.sticky['phpp_makeHBConstruction']
phpp_getWindowLibraryFromRhino = sc.sticky['phpp_getWindowLibraryFromRhino']

hb_hive = sc.sticky["honeybee_Hive"]()
hb_EPMaterialAUX = sc.sticky["honeybee_EPMaterialAUX"]()
HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)

#################################################################
# Def from HB 'addToEPLibrary'
def HB_addToEPLibrary(EPObject, overwrite):
    # import the classes
    
    hb_EPObjectsAux = sc.sticky["honeybee_EPObjectsAUX"]()
    added, name = hb_EPObjectsAux.addEPObjectToLib(EPObject, overwrite)
    
    if not added:
        msg = name + " is not added to the project library!"
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
        print msg
    else:
        print name + " is added to this project library!"

def createHBMats(_winName, _uwInst, _gVal):
    # Add the EP Construction name to the output list
    # Create a new 'U-w-installed' Material for the window
    matName = "PHPP_MAT_{}".format(_winName).upper()
    newHBMat = phpp_makeHBMaterial(
            matName, # Material Name
            _uwInst, # The new U-w-Installed
            _gVal,
            0.75) # Default VT
    
    # Create a new 'U-w-installed' Construction for the window
    constructionName = "PHPP_CONST_{}".format(_winName)
    new_EPConstruction = phpp_makeHBConstruction(constructionName, matName)
    
    return newHBMat, constructionName, new_EPConstruction

def cleanUpName(_baseName):
    nameNoSpaces = re.sub(r'\s+', '_', _baseName)
    nameNoPeriods = nameNoSpaces.replace(".", "-")
    nameAddPrefix = '__Win_' + nameNoPeriods
    
    return nameAddPrefix

def getWindowBasics(_in):
    with idf2ph_rhDoc():
        # Get the Window Geometry
        geom = rs.coercegeometry(_in)
        windowSurface  = ghc.BoundarySurfaces( geom )
        
        # Inset the window just slightly. If any windows touch one another or the zone edges
        # will failt to create a proper closed Brep. So shink them ever so slightly. Hopefully
        # not enough to affect the results. 
        windowPerim = ghc.JoinCurves( ghc.DeconstructBrep(windowSurface).edges, preserve=False )
        try:
            windowPerim = ghc.OffsetonSrf(windowPerim, 0.004, windowSurface) # 0.4mm so hopefully rounds down
        except:
            windowPerim = ghc.OffsetonSrf(windowPerim, -0.004, windowSurface)
        windowSurface = ghc.BoundarySurfaces( windowPerim )
        
        # Pull in the Object Name from Rhino Scene
        try:
            windowName = rs.ObjectName(_in)
        except:
            print('No!')
            warning = "Can't get the Name for some reason?\nDouble check:\n"\
            "If you are passing in Rhino Geometry,\nbe sure to pass it through\n"\
            "an 'ID' component before inputing into the '_windowGeom' input?"
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
            
        windowName = windowName if windowName!=None else 'Unnamed_Window'
        windowName = cleanUpName(windowName)
        
        
        # Double check that the surface Normal didn't get flipped
        c1 = ghc.Area(geom).centroid
        n1 = rs.SurfaceNormal(geom, c1)
        
        c2 = ghc.Area(windowSurface).centroid
        n2 = rs.SurfaceNormal(windowSurface, c2)
        
        normAngleDifference = ghc.Degrees( ghc.Angle(n1, n2).angle )
        if round(normAngleDifference, 0) != 0:
            # Flip the surface if it doesn't match the source
            windowSurface = ghc.Flip(windowSurface).surface
        
        return windowName, windowSurface

def getWindowUserText(_in):
    with idf2ph_rhDoc():
        windowDataFromRhino = {}
        
        if rs.IsUserText(_in):
            print 'Pulling data from Rhino model for Window: {}'.format(rs.ObjectName(_in))
            # Get the User-Text data from the Rhino Scene
            for eachKey in rs.GetUserText(_in):
                windowDataFromRhino[eachKey] = rs.GetUserText(_in, eachKey) 
            return windowDataFromRhino
        else:
            return None

def getWindowParamsFromLib(_in, _lib):
    instDepth = _in.get('InstallDepth', 0.1)
    varType = _in.get('VariantType', 'a')
    psiInstallType = _in.get('PsiInstallType', None)
    
    frameTypeName = _in['FrameType']
    frameTypeObj = _lib['lib_FrameTypes'][frameTypeName] # the dict is inside a list....
    
    glassTypeName = _in['GlazingType']
    glassTypeObj = _lib['lib_GlazingTypes'][glassTypeName]
    
    installs = [
    0 if _in['InstallLeft']=='False' else 1,
    0 if _in['InstallRight']=='False' else 1,
    0 if _in['InstallBottom']=='False' else 1,
    0 if _in['InstallTop']=='False' else 1
    ]
    
    # Update the Psi-Installs with UD Values if found in the Library
    if psiInstallType:
        sc.doc = Rhino.RhinoDoc.ActiveDoc
        psiInstalls_UD =  json.loads(rs.GetDocumentUserText('PHPP_lib_PsiInstall_'+psiInstallType) )
        sc.doc = ghdoc
        
        if psiInstalls_UD['Left'] != None:
            installs[0] = psiInstalls_UD['Left'] * installs[0]
        if psiInstalls_UD['Right'] != None: 
            installs[1] = psiInstalls_UD['Right'] * installs[1]
        if psiInstalls_UD['Bottom'] != None:
            installs[2] = psiInstalls_UD['Bottom'] * installs[2]
        if psiInstalls_UD['Top'] != None:
            installs[3] = psiInstalls_UD['Top'] * installs[3]
    
    return frameTypeObj, glassTypeObj, installs, varType, instDepth

def isZoneHost(_winSrfc, _HBzones):
    # Performs a BrepXBrep collision to see if the Windw
    # is 'on' any of the HB Zone Surfaces or not?
    
    for zone in _HBzones:
        if 'surfaces' not in dir(zone):
            warning = "Double check the input into _HBZones? It looks like perhaps you are\n"\
            "passing in just the HB Surfaces? Be sure to pass those surfaces to an HB 'createHBZones'\n"\
            "before this in order to create a 'Zone' before passing in here."
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning) 
            return None
        for srfc in zone.surfaces:
            if ghc.BrepXBrep(srfc.geometry, _winSrfc).curves != None:
                hostZone = zone.name
                return hostZone
    
    return None

def uniqueNames(_windowName, _HBzones, _hostZoneName):
    # Looks at the HB Zones to see if there are duplicate window names
    # Creates a unique name if so. Returns the name and sets the attr on the HB Zone
    
    for zone in _HBzones:
        if zone.name == _hostZoneName:
            try:
                # Get the master list of window names for the zone
                windowNameList = getattr(zone, 'windowNameList')
            except:
                # If there isn't one, start it now
                windowNameList = []
                setattr(zone, 'windowNameList', windowNameList)
            
            # See if the name already in the Master List. If not, add it
            if _windowName not in windowNameList:
                windowNameList.append(_windowName)
            else:
                # if it is a duplicate name, add a new unique ID number and then add it
                _windowName = "{}_{:03d}".format(_windowName, len(windowNameList) )
                windowNameList.append(_windowName)
            
            setattr(zone, 'windowNameList', windowNameList)
    
    return _windowName

def getParamsFromMultiLayerConsr(_materialsList):
    rValues = []
    
    for material in _materialsList:
        values, comments, UValue_SI, UValue_IP = hb_EPMaterialAUX.decomposeMaterial(material.upper(), ghenv.Component)
        rValues.append( 1/UValue_SI )
    
    gValue = 0.4 #DEFAULT
    
    return 1/sum(rValues), gValue

def getWindowParamsFromHBLib(_constName):
    # Calls the EP Construction from the HB Library
    frameUf_ = []
    frameW_ = []
    framePsiG_ = []
    framePsiInst_ = []
    glassU_ = None
    glassG_ = None
    glassVT_= None
    
    # First, Get the Construction Object from the HB Library
    construction = hb_EPMaterialAUX.decomposeEPCnstr(_constName.upper())
    if construction != -1:
        materials, comments, UValue_SI, UValue_IP = construction
        
        # Get all the Material Params as well
        if len(materials) == 1:
            # Simple Material Window (Single Layer), pull the parms
            values, comments, UValue_SI, UValue_IP = hb_EPMaterialAUX.decomposeMaterial(materials[0].upper(), ghenv.Component)
            EP_uValue = values[1]
            EP_gValue = values[2]
        elif len(materials) > 1:
            # MultiLayer Assembly. Calc the effective U-w
            EP_uValue, EP_gValue = getParamsFromMultiLayerConsr( materials )
    
    # Set up the outputs 
    for i in range(4):
        frameUf_.append(float( EP_uValue) )
        frameW_.append(0.12)
        framePsiG_.append(0)
        framePsiInst_.append(0)
        
    glassU_ = float( EP_uValue )
    glassG_ = float( EP_gValue )
    
    return glassU_, glassG_, frameUf_, frameW_, framePsiG_, framePsiInst_

def getWindowIDinfo(_count, _windowGeom):
    # Get the Window Name, Host Zone and Geometry from Rhino Scene
    windowName, windowSurface = getWindowBasics(_windowGeom)
    hostZoneName = isZoneHost(windowSurface, HBZoneObjects)
    
    try: windowName = inputs.getInputNames()[_count]
    except:pass
    
    # Fix Duplicate window names if there are any
    if hostZoneName:
        windowName = uniqueNames(windowName, HBZoneObjects, hostZoneName)
    else:
        warning = 'Cound not find the host zone for "{}"? Check the window geometry is Co-Planar with a Zone surface?'.format( windowName )
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning) 
    
    outputs.getWinNames().append( windowName )
    outputs.getWinSrfcs().append( windowSurface )
    
    return windowName, windowSurface, hostZoneName

def getWindowParams(_count, _windowGeom, _phppLibrary):
    
    #####################################
    # Sort out the Window's Params to use
    
    # 1)
    # Get the EP Construction params from the HB Library
    EPConstName = inputs.getEPConst()[_count] 
    glassU, glassG, frameUf, frameW, framePsiG, framePsiInst = getWindowParamsFromHBLib(EPConstName)
    
    glassTypeObj = PHPP_Glazing(EPConstName+'_Glass', _gValue=float(glassG), _uValue=float(glassU) )
    frameTypeObj = PHPP_Frame(EPConstName+'_Frame', _uValues=frameUf , _frameWidths=frameW , _psiGlazings=framePsiG , _psiInstalls=framePsiInst )
    installsObj = PHPP_Window_Install( [1,1,1,1] ) # Default
    variantType = 'a'
    instDepth = 0.1
    
    # 2)
    # Next, Try and pull Params from the Rhino scene if there are any
    windowDataFromRhino = getWindowUserText(_windowGeom)
    if windowDataFromRhino and _phppLibrary:
        windowParams = getWindowParamsFromLib(windowDataFromRhino, _phppLibrary)
        frameTypeObj, glassTypeObj, installValues, variantType, instDepth = windowParams
        installsObj = PHPP_Window_Install( installValues )
    
    # 3)
    # Last, get any user input Param Objects for frame, glass, installs, etc....
    try: glassTypeObj = inputs.getInputGlazing()[_count]
    except:pass
    
    try: frameTypeObj = inputs.getInputFrames()[_count]
    except:pass
    
    try:
        if len( inputs.getInputInstalls() ) == 1:
            # Could be a single value passed in, if so - use it for all four
            installsUI = []
            for i in range(4):
                installsUI.append( inputs.getInputInstalls()[0] )
            installsObj = PHPP_Window_Install(installsUI)
        elif len( inputs.getInputInstalls()[0] ) == 4:
            # Could be a list of 4 values, if so, create a new InstallObj with all four
            installsObj = PHPP_Window_Install( inputs.getInputInstalls()[0] )
    except:
        pass
    
    try:
        if len( inputs.getInputPsiInst() ) == 1:
            # Single value input, use this for all four
            setattr( frameTypeObj, 'psiInstLeft', float(inputs.getInputPsiInst()[0]) )
            setattr( frameTypeObj, 'psiInstRight', float(inputs.getInputPsiInst()[0]) )
            setattr( frameTypeObj, 'psiInstBottom', float(inputs.getInputPsiInst()[0]) )
            setattr( frameTypeObj, 'psiInstTop', float(inputs.getInputPsiInst()[0]) )
            
        elif len( inputs.getInputPsiInst() ) == 4:
            # List of 4 values, use this 
            setattr( frameTypeObj, 'psiInstLeft', float(inputs.getInputPsiInst()[0]) )
            setattr( frameTypeObj, 'psiInstRight', float(inputs.getInputPsiInst()[1]) )
            setattr( frameTypeObj, 'psiInstBottom', float(inputs.getInputPsiInst()[2]) )
            setattr( frameTypeObj, 'psiInstTop', float(inputs.getInputPsiInst()[3]) )
    except:
        pass
    
    return glassTypeObj, frameTypeObj, installsObj, variantType, instDepth

def cleanInput(_inputList, targetListLength, outputLocation, _defaultVal=None, ):
    # Used to make sure the input lists are all the same length
    # if the input param list isn't the same, use only the first item for all
    
    if len(_inputList) == targetListLength:
        for each in _inputList:
            outputLocation.append(each)
    elif len(_inputList) >= 1:
        for i in range(targetListLength):
            outputLocation.append(_inputList[0])
    elif len(_inputList) == 0 and _defaultVal != None:
        for i in range(targetListLength):
            outputLocation.append(_defaultVal)

class Outputs:
    """ Temporary object to store and manage the output values of the component
    
    Note to self: playing around with this method.... seems like I made
    it more complicated than it needs to be though. Maybe just go back to normal?
    """
    
    __windowSurfaces = None
    __windowNames = None
    __EPConstructions = None
    __windowObjs = None
    __HBConstructions = None
    __HBMaterials = None
    
    @staticmethod
    def getWinSrfcs():
        if Outputs.__windowSurfaces == None:
            Outputs.__windowSurfaces = []
        return Outputs.__windowSurfaces
    
    @staticmethod
    def getWinNames():
        if Outputs.__windowNames == None:
            Outputs.__windowNames = []
        return Outputs.__windowNames
    
    @staticmethod
    def getEPConsts():
        if Outputs.__EPConstructions == None:
            Outputs.__EPConstructions = []
        return Outputs.__EPConstructions
        
    @staticmethod
    def getWinObjs():
        if Outputs.__windowObjs == None:
            Outputs.__windowObjs = DataTree[Object]()
        return Outputs.__windowObjs
        
    @staticmethod
    def getHBConsts():
        if Outputs.__HBConstructions == None:
            Outputs.__HBConstructions = []
        return Outputs.__HBConstructions
    
    @staticmethod
    def getHBMats():
        if Outputs.__HBMaterials == None:
            Outputs.__HBMaterials = []
        return Outputs.__HBMaterials

class Inputs:
    
    __InputNames = None
    __InputFrames = None
    __InputGlazing = None
    __InputPsiInst = None
    __InputEPConst = None
    __Installs = None
    
    @staticmethod
    def getInputNames():
        if Inputs.__InputNames == None:
            Inputs.__InputNames = []
        return Inputs.__InputNames
    
    @staticmethod
    def getInputFrames():
        if Inputs.__InputFrames == None:
            Inputs.__InputFrames = []
        return Inputs.__InputFrames
    
    @staticmethod
    def getInputGlazing():
        if Inputs.__InputGlazing == None:
            Inputs.__InputGlazing = []
        return Inputs.__InputGlazing
    
    @staticmethod
    def getInputPsiInst():
        if Inputs.__InputPsiInst == None:
            Inputs.__InputPsiInst = []
        return Inputs.__InputPsiInst
    
    @staticmethod
    def getEPConst():
        if Inputs.__InputEPConst == None:
            Inputs.__InputEPConst = []
        return Inputs.__InputEPConst
    
    @staticmethod
    def getInputInstalls():
        if Inputs.__Installs == None:
            Inputs.__Installs = []
        return Inputs.__Installs

# Master Output class obj
outputs = Outputs()
inputs = Inputs()

cleanInput(names_, len(_windowGeom), inputs.getInputNames() )
cleanInput(frames_, len(_windowGeom), inputs.getInputFrames() )
cleanInput(glazings_, len(_windowGeom), inputs.getInputGlazing() )
cleanInput(psi_installs_, 4, inputs.getInputPsiInst() )
cleanInput(installs_, 4, inputs.getInputInstalls() )
cleanInput(EPConsts, len(_windowGeom), inputs.getEPConst(), 'Exterior Window')

# Get the Rhino Scene Document User-Text (Window Library)
if len(_windowGeom)>0 and len(_HBZones)>0:
    _phppLibrary = phpp_getWindowLibraryFromRhino()
else:
    _phppLibrary = None

# Get the Window Name, Host Zone and Geometry from Rhino Scene
if len(_windowGeom)>0 and len(_HBZones)>0:
    for i, winGeom in enumerate(_windowGeom):
        # Get the Basic Window Information, Data Parameters
        windowName, windowSurface, hostZoneName = getWindowIDinfo(i, winGeom)
        glassTypeObj, frameTypeObj, installsObj, variantType, instDepth = getWindowParams(i, winGeom, _phppLibrary)
        
        #####################################
        # Create the 'WINDOW' object from frame, glass, installs passed in
        newWindowObj = PHPP_WindowObject(
                                    windowName,
                                    glassTypeObj,
                                    frameTypeObj,
                                    installsObj,
                                    windowSurface,
                                    variantType,
                                    instDepth
                                    )
        
        # Calc the Uw Installed for the new Window Object (based on size, params...)
        uW_Inst = newWindowObj.getUwInstalled()
        
        # Create a NEW HB Materials and Constructions for the window
        # based on the U-W-installed, add to the output lists
        newHBMat, constructionName, new_EPConstruction = createHBMats(windowName, uW_Inst, newWindowObj.Type_Glass.gValue)
        outputs.getEPConsts().append(constructionName)
        
        # The full IDF text to write
        outputs.getHBMats().append(newHBMat) # The full txt to write to IDF
        outputs.getHBConsts().append(new_EPConstruction) # The full txt to write to IDF
        
        # Add the PHPP Style Window + UNIQUE name to a HB Zone master list, so can pull out later
        # Find the right zone to write to 
        for i, zone in enumerate(HBZoneObjects):
            if zone.name == hostZoneName:
                # If the Window Dict doesn't yet exist
                if 'phppWindowDict' in zone.__dict__.keys():
                    phppWindowDict = zone.phppWindowDict
                else:
                    setattr(zone, 'phppWindowDict', {})
                    phppWindowDict = zone.phppWindowDict
                
                # Add the new window Object to the Zone's Dict
                phppWindowDict[newWindowObj.Name] = newWindowObj
                
                # Add to a preview output
                outputs.getWinObjs().Add(newWindowObj, GH_Path(i))

# Pass along the Component Outputs
windowSurfaces_ = outputs.getWinSrfcs()
windowNames_ = outputs.getWinNames()
EPConstructions_ = outputs.getEPConsts()
windowObjs_ = outputs.getWinObjs()

#################################################################
# Add the HB Materials and Constructions to the HB Library
addToHBLib_ = True

if addToHBLib_ and len(outputs.getHBConsts())>0 and len(outputs.getHBMats())>0:
    print('----')
    for eachMat in outputs.getHBMats():
        HB_addToEPLibrary(eachMat, overwrite=True)
    print('----')
    for eachConst in outputs.getHBConsts():
        HB_addToEPLibrary(eachConst, overwrite=True)

# Add modified Surfaces / Zones back to the HB dictionary
if len(_HBZones)>0:
    HBZones_  = hb_hive.addToHoneybeeHive(HBZoneObjects, ghenv.Component)

