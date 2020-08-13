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
Creates DHW Recirculation loops for the 'DHW+Distribution' PHPP worksheet. Can create up to 5 recirculation loops. Will take in a DataTree of curves from Rhino and calculate their lengths automatically. Will try and pull curve object attributes from Rhino as well - use attribute setter to assign the pipe diameter, insulation, etc... on the Rhino side.
-
EM August 11, 2020
    Args:
        pipe_geom_: <Tree> A DataTree where each branch describes one 'set' of recirculation piping. 
        PHPP allows up to 5 'sets' of recirc piping. Each 'set' should include the forward and return piping lengths for that distribution leg.
        Use the 'Entwine' component to organize your inputs into branches before inputing if more than one set of piping is passed in.
------
The input here will accept either:
>  A single number representing the length (m) of the total loop in meters
>  A list (multiline input) representing multiple pipe lengths (m). These will be summed together for each branch
>  A curve or curves with no parameter values. The length of the curves will be summed together for each branch. You'll then need to enter the diam, insulation, etc here in the GH Component.
>  A curve or curves with paramerer values applied back in the Rhino scene. If passing in geometry with parameter value, be sure to use an 'ID' component before inputing the geometry here.
        pipe_diam_: <Optional> (mm) A List of numbers representing the nominal diameters (mm) of the pipes in each branch. This will override any values goten from the Rhino scene objects. If only one value is passed, will be used for all objects.
        insulThickness_: <Optional> (mm) A List of numbers representing the insulation thickness (mm) of the pipes in each branch. This will override any values goten from the Rhino scene objects. If only one value is passed, will be used for all objects.
        insulConductivity_: <Optional> (W/mk) A List of numbers representing the insulation conductivity (W/mk) of the pipes in each branch. This will override any values goten from the Rhino scene objects. If only one value is passed, will be used for all objects.
        insulReflective_: <Optional> A List of True/False values for the pipes in each branch. True=Reflective Wrapper, False=No Reflective Wrapper. This will override any values goten from the Rhino scene objects. If only one value is passed, will be used for all objects.
        insul_quality_: ("1-None", "2-Moderate", "3-Good") The Quality of the insulaton installation at the mountings, pipe-suspentions, couplings, valves, etc.. (Note: not the quality of the overall pipe insulation).
        daily_period_: (hours) The usage period in hours/day that the recirculation system operates. Default is 18 hrs/day.
    Returns:
        circulation_piping_: The Recirculation Piping object(s). Connect this to the 'circulation_piping_' input on the 'DHW System' component in order to pass along to the PHPP.
"""

ghenv.Component.Name = "BT_DHW_Piping_Recirc"
ghenv.Component.NickName = "Piping | Recirc"
ghenv.Component.Message = 'AUG_11_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import Rhino
import scriptcontext as sc
from contextlib import contextmanager
import rhinoscriptsyntax as rs
import ghpythonlib.components as ghc
import Grasshopper.Kernel as ghK
import inspect


# Classes and Defs
preview = sc.sticky['Preview']
PHPP_DHW_RecircPipe = sc.sticky['PHPP_DHW_RecircPipe']

@contextmanager
def rhDoc():
    try:
        sc.doc = Rhino.RhinoDoc.ActiveDoc
        yield
    finally:
        sc.doc = ghdoc

def cleanPipeDimInputs(_diam, _thickness):
    # Clean diam, thickness
    if _diam != None:
        if " (" in _diam: 
            _diam = float( _diam.split(" (")[0] )
    
    if _thickness != None:
        if " (" in _thickness: 
            _thickness = float( _thickness.split(" (")[0] )
    
    return _diam, _thickness

def getPipe_FromRhino(_pipeObj, _i):
    # If its a numeric input, just return that as the length and the rest as blank
    try:
        return float(_pipeObj), None, None, None, None
    except:
        pass
    
    # Otherwise, try and pull the required info from the Rhino scene
    # Note, in order to keep the component's Type-Hint as 'No Type Hint' so that
    # I can input either crvs or numbers, need to get the Rhino GUID using the 
    # Compo.Params.Input method below.
    try:
        rhinoGuid = ghenv.Component.Params.Input[0].VolatileData[0][_i].ReferenceID.ToString()
        pipeLen = ghc.Length( rs.coercecurve(_pipeObj) )
    except:
        rhinoGuid = None
        pipeLen = None
    
    with rhDoc():
        if pipeLen and rhinoGuid:
            try:
                conductivity =  float( rs.GetUserText(rhinoGuid, 'insulation_conductivity') )
                reflective = rs.GetUserText(rhinoGuid, 'insulation_reflective')
                thickness =  rs.GetUserText(rhinoGuid, 'insulation_thickness')
                diam = rs.GetUserText(rhinoGuid, 'pipe_diameter')
                
                diam, thickness = cleanPipeDimInputs(diam, thickness)
                
                return pipeLen, diam, thickness, conductivity, reflective
            except:
                msg = "Note: I could not find any parameters assigned to the pipe geometry in the Rhino Scene.\n"\
                "Be sure that you enter parameter information for the diam, thickness, etc... in the component here.\n"\
                "------\n"\
                "If you did assign information to the curves in the Rhino scene, be sure to use an 'ID' component\n"\
                "BEFORE you pass those curves into the 'pipe_geom_' input."
                ghenv.Component.AddRuntimeMessage( ghK.GH_RuntimeMessageLevel.Remark, msg)
                
                return pipeLen, None, None, None, None
        else:
            msg = "Something went wrong getting the Curve length?\n"\
            "Be sure you are passing in a curve / polyline object or a number to the 'pipe_geom_'?"
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
            
            return None, None, None, None, None

def checkFloat(_inputVal, _nm):
    if _inputVal == None:
        return None
    else:
        try:
            return float(_inputVal)
        except:
            unitWarning = "Inputs for '{}' should be numbers".format(_nm)
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, unitWarning)
            return _inputVal

def getParams_fromGH(_ghParams, _pipeObj, _attrName):
        # Check the input type
        for eachParam in _ghParams:
            checkFloat(eachParam, _attrName)
        
        if len(_ghParams) == len(_pipeObj):
            for i, eachObj in enumerate(_pipeObj):
                setattr(eachObj, _attrName, _ghParams[i])
        else:
            for eachObj in _pipeObj:
                setattr(eachObj, _attrName, _ghParams[0])
        
        return _pipeObj

def combinePipeObjs(_pipeObjList):
    # for each pipe object, make a dict of all the values 
    masterDict = {}
    for pipeObj in _pipeObjList:
        for k, v in pipeObj.__dict__.items():
            if k in masterDict.keys():
                masterDict[k] = masterDict[k] + [v]
            else:
                masterDict[k] = [v]
    
    # Go through all the keys, if any variation occurs, throw an error
    for k, v in masterDict.items():
        if 'length' not in k:
            if len(set(v)) != 1:
                paramWarning = "Warning: Component found multiple parameters for '{}' on the geometry?\nOnly the parameter values from the first object will be used for this pipe-set.".format(k) 
                ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, paramWarning)
    
    # combine the lengths
    totalLength = sum(masterDict['length'])
    
    # return a single pipe object
    new_pipe_obj = _pipeObjList[0]
    new_pipe_obj.length = totalLength
    
    return new_pipe_obj

# ------------------------------------------------------------------------------
# Default Inputs, Clean Inputs
if insul_quality_:
    if "3" in str(insul_quality_)  or "Good" in str(insul_quality_): insulQual = "3-Good"
    elif "2" in str(insul_quality_) or "Moderate" in str(insul_quality_): insulQual = "2-Moderate"
    else: insulQual = "1-None"
else:
    insulQual = "1-None"

dlyperiod = daily_period_ if daily_period_ != None else 18
if dlyperiod:
    try: float(dlyperiod)
    except: ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, '"daily_period_" should be a number.')

# ------------------------------------------------------------------------------
# Build the Re-Circ Pipe Objects
circulation_piping_ = []
if pipe_geom_.BranchCount > 0:
    for branch in pipe_geom_.Branches:
        branch_pipe_objs = []
        for i, geom in enumerate(branch):
            
            newRecircObj = PHPP_DHW_RecircPipe()
            
            (newRecircObj.length,        
            newRecircObj.diam,
            newRecircObj.insulThck,
            newRecircObj.insulCond,
            newRecircObj.insulRefl ) = getPipe_FromRhino(geom, i)
            
            newRecircObj.quality = insulQual
            newRecircObj.period = dlyperiod
            
            branch_pipe_objs.append(newRecircObj)
        
        # If multiple pipes in a single branch, combine them into one
        if len(branch_pipe_objs) > 1:
            circulation_piping_.append( combinePipeObjs(branch_pipe_objs) )
        elif len(branch_pipe_objs) == 1:
            circulation_piping_.append( branch_pipe_objs[0] )

# ------------------------------------------------------------------------------
# Override Rhino obj with GH params (if any)
if len(pipe_diam_)>0: circulation_piping_ = getParams_fromGH(pipe_diam_, circulation_piping_, 'diam')
if len(insulThickness_)>0: circulation_piping_ = getParams_fromGH(insulThickness_, circulation_piping_, 'insulThck')
if len(insulConductivity_)>0: circulation_piping_ = getParams_fromGH(insulConductivity_, circulation_piping_, 'insulCond')
if len(insulReflective_)>0: circulation_piping_ = getParams_fromGH(insulReflective_, circulation_piping_, 'insulRefl')

# ------------------------------------------------------------------------------
# Preview the objects
for each in circulation_piping_:
    print('----\nDHW Recirc Piping:')
    preview(each)
