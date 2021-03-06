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
Use this before a Honeybee 'createHBSrfs' component to pull geometry and relevant surface parameters
from the Rhino scene rather than setting them in GH. This component tries to read data from each
surface's 'Attribute User Text'. Use the Rhino-scene Set Params tool to set surface data before
trying to import it here. 
-
EM July 31, 2020
    Args:
        _srfcs: The Zone's Opaque surfaces as a list (walls, floors, ceilings, etc...). By Default Type-Hint is set to 'GUID' in order to get geom data parameters from the Rhino scene. If passing in Grasshopper generated surfaces be sure to set Type-Hint to 'No Type Hint'.
        autoOrientation_: (bool Default='False') Set to 'True' to have this component automatically assign surface type ('wall', 'floor', 'roof'). useful if you are testing massings / geometry and don't want to assign explicit type everytime. If you have already assigned the surface type in Rhino, leave this set to False. If 'True' this will override any values found in the Rhino scene.
    Returns:
        srfcGeometry_: Connect to the '_geometry' Input on a Honeybee 'createHBSrfs' component
        srfcNames_: Connect to the 'srfName_' Input on a Honeybee 'createHBSrfs' component
        srfcTypes_: Connect to the 'srfType' Input on a Honeybee 'createHBSrfs' component
        srfcEPBCs_: Connect to the 'EPBC_' Input on a Honeybee 'createHBSrfs' component
        srfcEPConstructions_: Connect to the '_EPConstruction' Input on a Honeybee 'createHBSrfs' component. If blank or any Null values pased, will use HB defaults as usual.
        srfcRADMaterials_: Connect to the '_RADMaterial' Input on a Honeybee 'createHBSrfs' component
"""

ghenv.Component.Name = "BT_GetSurfaceParams"
ghenv.Component.NickName = "Get Surface Params"
ghenv.Component.Message = 'JUL_31_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import Rhino
import scriptcontext as sc
import rhinoscriptsyntax as rs
import json
import Grasshopper.Kernel as ghK
import System
import rhinoscriptsyntax as rs
import re
import math
import ghpythonlib.components as ghc

# Defs and Classes
hb_hive = sc.sticky["honeybee_Hive"]()
hb_EPMaterialAUX = sc.sticky["honeybee_EPMaterialAUX"]()

phpp_makeHBMaterial_NoMass = sc.sticky['phpp_makeHBMaterial_NoMass']
phpp_makeHBMaterial_Opaque = sc.sticky['phpp_makeHBMaterial_Opaque']
phpp_makeHBConstruction = sc.sticky['phpp_makeHBConstruction']
phpp_createSrfcHBMatAndConst = sc.sticky['phpp_createSrfcHBMatAndConst']
idf2ph_rhDoc = sc.sticky['idf2ph_rhDoc']

#################################################################
class Outputs:
    """ Temporary object to store and manage the output values of the component
    """
    
    __srfcGeom = None
    __srfcNames = None
    __srfcTypes = None
    __srfcEPBCs = None
    __EPConsts = None
    __RADMats = None
    
    @staticmethod
    def getSrfcGeom():
        if Outputs.__srfcGeom == None:
            Outputs.__srfcGeom = []
        return Outputs.__srfcGeom
    
    @staticmethod
    def getSrfcNames():
        if Outputs.__srfcNames == None:
            Outputs.__srfcNames = []
        return Outputs.__srfcNames
    
    @staticmethod
    def getSrfcTypes():
        if Outputs.__srfcTypes == None:
            Outputs.__srfcTypes = []
        return Outputs.__srfcTypes
        
    @staticmethod
    def getSrfcEPBCs():
        if Outputs.__srfcEPBCs == None:
            Outputs.__srfcEPBCs = []
        return Outputs.__srfcEPBCs
        
    @staticmethod
    def getEPConstructions():
        if Outputs.__EPConsts == None:
            Outputs.__EPConsts = []
        return Outputs.__EPConsts
    
    @staticmethod
    def getRADMaterials():
        if Outputs.__RADMats == None:
            Outputs.__RADMats = []
        return Outputs.__RADMats

def HB_addToEPLibrary(EPObject, overwrite, printLog=False):
    """ This definition is take right from the Honeybee 'addToEPLibrary' Component
    
    Takes in the EP Object (material or construction) and writes it out to the
    'Hive' so it can be used / access later when writing the IDF file. 
    Arguments:
        EPObject: The EP Material or EP Construction to write to the Hive
        overwrite: True / False to overwright existing values with same name in the Hive
        printLog: True / False to print to log what this is doing / succes or not. 
    Returns:
        None
    """
    hb_EPObjectsAux = sc.sticky["honeybee_EPObjectsAUX"]()
    added, name = hb_EPObjectsAux.addEPObjectToLib(EPObject, overwrite)
    
    if not added:
        msg = name + " is not added to the project library!"
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
        if printLog: print msg
    else:
        if printLog: print name + " is added to this project library!"

def getAllUserText(_GUID):
    """ Takes in an objects GUID and returns the full dictionary of
    Attribute UserText Key and Value pairs. Cleans up a bit as well.
    
    Arguments:
        _GUID: the Rhino GUID of the surface object to try and read from
    Returns:
        dict: a dictionary object with all the keys / values found in the Object's UserText
    """
    dict = {}
    
    if not _GUID.GetType() == System.Guid:
        remark = "Unable to get parameter data for the surface? If trying to pull data\n"\
        "from Rhino, be sure the '_srfc' input Type Hint is set to 'Guid'\n"\
        "For now, using default values for all surface parameter values."
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Remark, remark)
        return dict
    
    with idf2ph_rhDoc():
        dict['Object Name'] = rs.ObjectName(_GUID) # Always get the name
        
        for eachKey in rs.GetUserText(_GUID):
            if 'Object Name' not in eachKey:
                val =  rs.GetUserText(_GUID, eachKey)
                if val != 'None':
                    dict[eachKey] = val
                else:
                    dict[eachKey] = None
    
    return dict

def warnNoName(_dict):
    """ Gives warning if there isn't any name found in the UserText for the surface
    
    Arguments:
        dict: the Dictionary with the UserText keys / values
    Returns:
        srfcName: 'NoName' if none found, otherwise the user-determined name
    """
    
    srfcName = _dict.get("Object Name", "NoName")
    if srfcName == 'None' or srfcName == None:
        warning = "Warning: Some Surfaces look like they are missing names? It is likely that\n"\
        "the Honeybee solveAdjc component will not work correctly without names.\n"\
        "This is especially true for 'interior' surfaces up against other thermal zones.\n"\
        "Please apply a unique name to all surfaces before proceeding. Use the\n"\
        "IDF2PH 'Set Window Names' tool to do this automatically (works on regular srfcs too)"
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
    return srfcName

def addEPMatAndEPConstToLib(_thisConst, _HBMaterials, _EPConstructions, _HBConstructions, _udConstDict):
    """ Creates and adds new HB Material and Constructions to the master output lists
    
    Arguments:
        _thisConst: The name of the Construction found on the Surface's UserText
        _HBMaterials: A List of the Honeybee Material Objects
        _EPConstructions: A List of the EnergyPlus Construction Objects
        _HBConstructions: A List of the Honeybee Construction Objects
        _udConstDict: The dictionary of Assembly Parameters from the Rhino DocumentUserText
    Returns:
        _HBMaterials: A List of the Honeybee Material Objects
        _EPConstructions A List of the EnergyPlus Construction Objects
        _HBConstructions: A List of the Honeybee Construction Objects
    """
    
    if _thisConst == None:
        return _HBMaterials, _EPConstructions, _HBConstructions
    
    # Get the UD Assembly dict parameter values, depending on type
    if isinstance(_udConstDict, str):
        vals = json.loads(_udConstDict.get(_thisConst, {}))
    elif isinstance(_udConstDict, dict):
        vals = _udConstDict.get(_thisConst, {})
    else:
        vals = dict()
    
    # Create a new EP Material and EP Construction from the read-in values
    try:
        nm = vals.get('Name','NO_NAME').upper()
        uval = float(vals.get('uValue',1))
        intInsul = float(vals.get('intInsulation',0))
    except:
        nm = vals.get('Name','NO_NAME').upper()
        uval = 1
        intInsul = 0
        warning = 'Something went wrong getting the Construction Assembly info for the Assembly type: "{}" from the Rhino Library?\n'\
        'No assembly by that name, or a non-numeric entry found in the Library? Please check the Library File inputs.\n'\
        'For now, I will assign a U-Value of 1.0 W/m2k to this assembly as a default value.'.format(_thisConst.upper())
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
    
    new_HBMat, new_HBMat_Mass, constructionName, new_EPConstruction = phpp_createSrfcHBMatAndConst(nm, uval, intInsul)
    
    _HBMaterials.append(new_HBMat) # The full txt to write to IDF
    _HBMaterials.append(new_HBMat_Mass)
    _EPConstructions.append(constructionName)
    _HBConstructions.append(new_EPConstruction) # The full txt to write to IDF
        
    return _HBMaterials, _EPConstructions, _HBConstructions

def getSurfaceParams(_srfc):
    """ Takes in a GUID of a surface and reads UserText, bundles output
    
    This will log all the geom and parameters it finds to the master 'output' object
    
    Arguments:
        _srfc: The surface GUID
    Returns:
        None
    """
    # Check in the inputs
    if _srfc == None:
        warning = "Warning: Inputs should be Surface objects.\n"\
        "If you are passing in Grasshopper generated surfaces, (rather than\n"\
        "referenced surfaces from Rhino scene) be sure to set the 'Type Hint'\n"\
        "for the '_srfcs' input on this component to 'No Type Hint'"
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
        return
    
    # Get access to the output lists
    srfcGeom = outputs.getSrfcGeom()
    srfcNames = outputs.getSrfcNames()
    srfcTypes = outputs.getSrfcTypes()
    srfcEPBCs = outputs.getSrfcEPBCs()
    srfcEPConstructions = outputs.getEPConstructions()
    srfcRADMaterials = outputs.getRADMaterials()
    
    # Bring in the Attribute UserText from Rhino
    dict = getAllUserText(_srfc)
    
    # Add the new Geom and Params to the output lists
    with idf2ph_rhDoc(): srfcGeom.append( rs.coercebrep(_srfc) )
    
    srfcNames.append(  warnNoName(dict)  )
    srfcEPConstructions.append( cleanGet(dict, 'EPConstruction', None) )
    srfcRADMaterials.append( cleanGet(dict, 'RADMaterial', None) )
    
    # Fix for the UndergroundSlab issue
    srfcTypeval = srfcTypeCodes.get( cleanGet(dict, 'srfType', 'WALL') )
    srfcTypes.append( srfcTypeval )
    if srfcTypeval == 2.25:
        srfcEPBCs.append( None )
    else:
        srfcEPBCs.append( cleanGet(dict, 'EPBC', 'Outdoors') )

def cleanGet(_dict, _k, _default=None):
    """ Protects against None or space(s) (' ') accidental inputs """
    if _k == None or _dict==None:
        return None
    
    v = _dict.get(_k, _default)
    
    # re will return None if v == only one or more spaces (no alph numeric)
    if v == None or re.search('\S+', v) == None:
        return None
    else:
        return v

def cleanSrfcInput(_srfcsInput):
    """If Brep or Polysrfc are input, explode them"""
    
    outputSrfcs = []
    
    with idf2ph_rhDoc():
        for inputObj in _srfcsInput:
            if isinstance(rs.coercesurface(inputObj), Rhino.Geometry.BrepFace):
                # Catches Bare surfaces
                outputSrfcs.append(inputObj)
            elif isinstance(rs.coercebrep(inputObj), Rhino.Geometry.Brep):
                # Catches Polysurfaces / Extrusions or other Masses
                faces = ghc.DeconstructBrep(rs.coercebrep(inputObj)).faces
                if isinstance(faces, list):
                    for face in faces:
                        outputSrfcs.append(face)
            elif isinstance(rs.coercegeometry(inputObj), Rhino.Geometry.PolylineCurve):
                # Catches PolylineCurves
                if not rs.coercegeometry(inputObj).IsClosed:
                    warn = 'Non-Closed Polyline Curves found. Make sure all curves are closed.'
                    ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Remark, warn)
                else:
                    faces = ghc.DeconstructBrep(rs.coercegeometry(inputObj)).faces
                    if isinstance(faces, list):
                        for face in faces:
                            outputSrfcs.append(face)
                    else:
                        outputSrfcs.append(faces)
        
        return outputSrfcs

def determineSurfaceType(_srfc):
    # Code here adapted from Honeybee 'decomposeZone' method
    # Checks the surface normal and depending on the direction, 
    # assigns it as a 'wall', 'floor' or 'roof'
    
    def getGHSrfNormal(GHSrf):
        centroid = ghc.Area(GHSrf).centroid
        normalVector = ghc.EvaluateSurface(GHSrf, centroid).normal
        return normalVector, GHSrf
    
    maximumRoofAngle = 30
    # check surface Normal
    try:
        normal, surface = getGHSrfNormal(_srfc)
    except:
        print('Failed to find surface normal. Are you sure it is a surface?')
        return 0
    
    angle2Z = math.degrees(Rhino.Geometry.Vector3d.VectorAngle(normal, Rhino.Geometry.Vector3d.ZAxis))
    
    if  angle2Z < maximumRoofAngle or angle2Z > 360 - maximumRoofAngle:
        return 1 #roof
    elif  160 < angle2Z <200:
        return 2 # floor
    else: 
        return 0 #wall

# Surface Type Schema
srfcTypeCodes = {'WALL':0,
    'UndergroundWall':0.5,
    'ROOF':1,
    'UndergroundCeiling':1.5,
    'FLOOR':2,
    'UndergroundSlab':2.25,
    'SlabOnGrade':2.5,
    'ExposedFloor':2.75,
    'CEILING':3,
    'AIRWALL':4,
    'WINDOW': 5,
    'SHADING': 6
    }

# Master Output class obj
outputs = Outputs()

#################################################################
# Pull the geometry and params for each Surface in the input set
srfcsToUse = cleanSrfcInput(_srfcs)
for srfc in srfcsToUse:
    getSurfaceParams( srfc )

# Pass along the Component Outputs
srfcGeometry_ = outputs.getSrfcGeom() 
srfcNames_ = outputs.getSrfcNames()
srfcTypes_ = outputs.getSrfcTypes()
srfcEPBCs_ = outputs.getSrfcEPBCs()
srfcEPConstructions_ =  outputs.getEPConstructions()
srfcRADMaterials_ =  outputs.getRADMaterials()

# Get the Library of EP Construction Params from the Rhino Document User Text dictionary
udConstParams = {}
with idf2ph_rhDoc():
    if rs.IsDocumentUserText():
        for eachKey in rs.GetDocumentUserText():
            if 'PHPP_lib_Assmbly_' in eachKey:
                assmblyName = json.loads(rs.GetDocumentUserText(eachKey)).get('Name')
                udConstParams[assmblyName] = json.loads(rs.GetDocumentUserText(eachKey))

#################################################################
# For each construction, make a new EP Material for it, then a new EP Construction
HBMaterials_ = []
EPConstructions_ = []
HBConstructions_ = []

uniqueEPConstTypes = list(set(srfcEPConstructions_))
for eachConst in uniqueEPConstTypes:
    HBMaterials_, EPConstructions_, HBConstructions_  = addEPMatAndEPConstToLib(eachConst, HBMaterials_, EPConstructions_, HBConstructions_, udConstParams)

# Find the right Surface Construction Name to pass
for i, srfcEPConst in enumerate(srfcEPConstructions_):
    if srfcEPConst == None:
        continue
    thisSrfcName = srfcEPConst.upper().replace(" ", "_")
    for epConstrName in EPConstructions_:
        if thisSrfcName in epConstrName:
            srfcEPConstructions_[i] = epConstrName

#################################################################
# Check if there are NO floors in the set, if there are not any, try and 
# make one or another surface the floor as default
if autoOrientation_:
    for i, srfc in enumerate(srfcGeometry_):
        srfcTypes_[i] = determineSurfaceType(srfc)
else:
    if not any(item in srfcTypes_ for item in [2, 2.25, 2.5, 2.75]):
        warn = 'No "Floor" type surfaces found? Note that if you try and pass\n'\
        "a Honeybee zone to the OpenStudio exporter without a floor it will cause\n"\
        "an error. Either add a floor surface, or set 'autoOrientation_' to TRUE"
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Remark, warn)

#################################################################
# Add the HB Materials and Constructions to the HB Library
addToHBLib_ = True
showLog = True #For debugging

if addToHBLib_ and HBConstructions_ and HBMaterials_:
    for eachMat in HBMaterials_:
        HB_addToEPLibrary(eachMat, overwrite=True, printLog=showLog)
    print '----------'
    for eachConst in HBConstructions_:
        HB_addToEPLibrary(eachConst, overwrite=True, printLog=showLog)
