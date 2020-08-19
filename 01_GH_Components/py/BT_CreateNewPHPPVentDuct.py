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
Collects and organizes data for a duct for a Ventilation System.
-
EM August 19, 2020
    Args:
        ductLength_: List<Float | Curve> Input either a number for the length of the duct from the Ventilation Unit to the building enclusure, or geometry representing the duct (curve / line)
        ductWidth_: List<Float> Input the diameter (mm) of the duct. Default is 101mm (4")
        insulThickness_: List<Float> Input the thickness (mm) of insulation on the duct. Default is 52mm (2")
        insulConductivity_: List<Float> Input the Lambda value (W/m-k) of the insualtion. Default is 0.04 W/mk (Fiberglass)
    Returns:
        hrvDuct_: A Duct object for the Ventilation System. Connect to the 'hrvDuct_01_' or 'hrvDuct_02_' input on the 'Create Vent System' to build a PHPP-Style Ventialtion System.
"""

ghenv.Component.Name = "BT_CreateNewPHPPVentDuct"
ghenv.Component.NickName = "Vent Duct"
ghenv.Component.Message = 'AUG_19_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import rhinoscriptsyntax as rs
import ghpythonlib.components as gh
import scriptcontext as sc
import Rhino
import Grasshopper.Kernel as ghK

# Classes and Defs
preview = sc.sticky['Preview']
PHPP_Sys_Duct = sc.sticky['PHPP_Sys_Duct']
phpp_convertValueToMetric = sc.sticky['phpp_convertValueToMetric']

def getDuctInputIndexNumbers():
    """ Looks at the component's Inputs and finds the ones labeled 'ductLength_'
    
    Returns:
        The list Index value for the "ductLength_" input 
        Returns None if no match found
    """
    
    hrvDuct_inputNum = None
    
    for i, input in enumerate(ghenv.Component.Params.Input):
        if 'ductLength_' == input.Name:
            hrvDuct_inputNum = i
        
    return hrvDuct_inputNum

def determineDuctToUse(_inputList, _inputIndexNum):
    output = []
    
    if len(_inputList) != 0:
        for i, ductItem in enumerate(_inputList):
            if isinstance(ductItem, Rhino.Geometry.Curve):
                duct01GUID = ghenv.Component.Params.Input[_inputIndexNum].VolatileData[0][i].ReferenceID.ToString()
                output.append( duct01GUID )
            else:
                try:
                    output.append(phpp_convertValueToMetric(ductItem, 'M') )
                except:
                    output.append( ductItem )
    else:
        output = [5]
    
    return output

def getParamValueAsList(_targetLen, _inputList, _default, _unit):
    outputList = []
    
    if len(_inputList) == _targetLen:
        for item in _inputList:
            outputList.append(phpp_convertValueToMetric(item, _unit) if item else _default)
    elif len(_inputList) != _targetLen and len(_inputList) != 0:
        for i in range(_targetLen):
            try:
                outputList.append(phpp_convertValueToMetric(_inputList[i], _unit))
            except:
                outputList.append(phpp_convertValueToMetric(_inputList[-1], _unit))
    else:
        outputList = [_default] * _targetLen
    
    return outputList

#-------------------------------------------------------------------------------
# Clean up the inputs
hrvDuct_inputNum = getDuctInputIndexNumbers()
ductLengths = determineDuctToUse(ductLength_, hrvDuct_inputNum)

#-------------------------------------------------------------------------------
# Build the Duct Object
hrvDuct_ = PHPP_Sys_Duct(
        ductLengths,
        getParamValueAsList(len(ductLengths), ductWidth_, 104, 'MM'),
        getParamValueAsList(len(ductLengths), insulThickness_, 52, 'MM'),
        getParamValueAsList(len(ductLengths), insulConductivity_, 0.04, 'W/MK')
        )

#-------------------------------------------------------------------------------
for warning in hrvDuct_.Warnings:
    ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)

preview(hrvDuct_)