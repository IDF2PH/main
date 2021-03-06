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
Builds Thermal Bridge (linear, point) objects to add to the PHPP. Note that these objects are generally not inluced in the EnergyPlus model and so may be a source or discrepancy between the EnergyPlus reslults and the PHPP results. It will be more accurate to include these items in the final model.
-
EM August 1, 2020

    Args:
        estimated_tb_: <Optional> A single number (0 to 1) which represents the % increase in heat loss due to thermal bridging.
        Typical values could be: Passive House: +0.05, Good: +0.10,Medium: +0.25, Bad +0.40
        linear_tb_names_: <Optional> A list of names for the Thermal Bridge items to add to the PHPP. This will override any inputs in 'linear_tb_geom_'
        linear_tb_lengths_: <Optional> A list of Lengths (m) for the Thermal Bridge items to add to the PHPP. This will override any inputs in 'linear_tb_geom_'
        linear_tb_PsiValues_: <Optional> Either a single number or a list of Psi-Values (W/mk) to use for any Thermal Bridge items.
        If no values are passed, a default 0.01 W/mk value will be used for all Thermal Bridge items. If only one number is input, it will be used for all Psi-Value items output.
        linear_tb_geom_: <Optional> A list of Rhino Geometry (curves, lines) to use for finding lengths and names automatically.
        ------
        If you want this to read from Rhino rather than GH, pass all your referenced geometry thorugh an 'ID' (Primitive/GUID) object first before inputting.
        This will try and read the name from the Rhino-Scene object and use the object's name for the PHPP entry (Rhino: Properties/Object Name/...)
        -----
        point_tb_Names_: <Optional> A list of the Point-Thermal-Bridge names. Each name will become a unique point-TB object in the PHPP.
        point_tb_ChiValues_: <Optional> A list of Chi Values (W/mk) to use for the Point thermal bridges. If this list matches the 'point_tb_Names' input each will be used. If only one value is input, it will be used for all Chi-Values. If no values are input, a default of 0.1-W/mk will be used for all.
        Returns:
        thermalBridges_: Thermal Bridge objects to write to Excel. Connect to the 'thermalBridges_' input on the '2XL |  PHPP Geom' Component
"""

ghenv.Component.Name = "BT_SetTB"
ghenv.Component.NickName = "Thermal Bridges"
ghenv.Component.Message = 'AUG_01_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc
import Grasshopper.Kernel as ghK
import rhinoscriptsyntax as rs
import ghpythonlib.components as ghc
import Rhino
from contextlib import contextmanager
import json

preview=sc.sticky['Preview']

class PHPP_ThermalBridge:
    def __init__(self, _nm, _len, _psi, _geom, _groupNo=15, _fRsi=None):
        self.Name = _nm
        self.Length = float(_len)
        
        try: self.PsiValue = float(_psi)
        except: 
            if '=SUM' in str(_psi): # Cus for estimated this wants to be a formula
                self.PsiValue = _psi
        
        self.geometry = _geom
        self.GroupNo = _groupNo
        try: self.fRSI =  float(_fRsi)
        except: self.fRSI = None
    
    def __repr__(self):
        str =("A PHPP Thermal Bridge Object:", 
                " > Name: {}".format(self.Name),
                " > Length: {:.02f} (m)".format(self.Length),
                " > Psi-Value: {} (W/mk)".format(self.PsiValue),
                " > Group Num: {}".format(self.GroupNo),
                " > fRsi: {}".format(self.fRSI),
                )
        
        return "\n".join(str)

def findLength(_inputItem):
    # takes in an Input object and tries to find the length of it
    
    try:
        length = float(_inputItem)
    except:
        try:
            line = rs.coercecurve(_inputItem)
            length = ghc.Length(line)
        except:
            msgLenError = "Couldn't get the length from the inputs. Please input a Curve or Line for each Thermal Bridge item"
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msgLenError)
    
    return length

@contextmanager
def rhDoc():
    try:
        sc.doc = Rhino.RhinoDoc.ActiveDoc
        yield
    finally:
        sc.doc = ghdoc

def getTBLibFromDocText():
    """ Goes and gets the TB library items from the Document User-Text
    Will return a dict of dicts, ie:
        {'TB_Name_01':{'Name':'example', 'fRsi':0.77, 'Psi-Value':0.1},... }
    """
    dict = {}
    
    with rhDoc():
        if rs.IsDocumentUserText():
            keys = rs.GetDocumentUserText()
            
            tbKeys = [key for key in keys if 'PHPP_lib_TB_' in key]
            for key in tbKeys:
                try:
                    val = json.loads( rs.GetDocumentUserText(key) )
                    name = val.get('Name', 'Name Not Found')
                    dict[name] = val
                except:
                    msg = "Problem getting Psi-Values for '{}' from the File. Check the\n"\
                    "DocumentUserText and make sure the TBs are loaded properly".format(name)
                    ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
    return dict

def getTBfromRhino(_tbEdges, _tbLib, _unames, _ulengths, _uPsiVals):
    """Looks at the edges passed in and tries to pull relevant params from Rhino
    Will look for in the Edge's UserText and try and get the 'Typename' and 'Group'
    will also look for any DocumentUserText TB libraries and get params based on
    the object's 'Typename'
    Will also allow for user-override by passing values into the GH Component directly
    for names, lengths and Psi-Values
    """
    tbGeom_ = []
    tbLens_ = []
    tbNames_ = []
    tbPsiVals_ = []
    tbGroups_ = []
    tbrfRSIs_ = []
    
    with rhDoc():
        for i, edge in enumerate(_tbEdges):
            # Get the params from the Rhino Object
            crv = rs.coercecurve(edge)
            tbLen = ghc.Length(crv)
            nm = rs.GetUserText(edge, 'Typename')
            grp = rs.GetUserText(edge, 'Group')
            
            
            if nm not in _tbLib.keys():
                msg = "Can not find the values for TB: '{}' in the file library?\n"\
                "Check the file's DocumentUserText and make sure the name is correct?".format(nm)
                ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
            
            # Get the params from the DocUserTxt Library
            params = _tbLib.get(nm, None)
            psiVal = params.get('psiValue', 0.01) if params != None else 0.01
            fRSI = params.get('fRsi', None) if params != None else None
            
            # Get user input params
            try: nm = _unames[i]
            except: pass
            
            try: tbLen = _ulengths[i]
            except: pass
            
            try: psiVal = _uPsiVals[i]
            except: pass
            
            # Add params to the output lists
            tbNames_.append(nm if nm != None else 'TB_{:0>2d}'.format(i+1))
            tbLens_.append(tbLen)
            tbGroups_.append(grp.split(':')[0] if grp != None else 15) # default = 15: Ambient
            tbGeom_.append(crv if crv !=None else edge)
            tbPsiVals_.append(psiVal)
            tbrfRSIs_.append(fRSI)
    
    return tbNames_, tbLens_, tbPsiVals_, tbGeom_, tbGroups_, tbrfRSIs_

def getTBfromGH(_uNames, _uLens, _uPsiVals):
    
    def cleanAppend(_list, _i, _default=None):
        try:
            val = _list[_i]
        except IndexError:
            try:
                val = _list[0]
            except:
                val = _default
        return val
    
    tbGeom_ = []
    tbLens_ = []
    tbNames_ = []
    tbPsiVals_ = []
    tbGroups_ = []
    tbrfRSIs_ = []
    
    for i, len in enumerate(_uLens):
        tbLens_.append(len)
        tbNames_.append( cleanAppend(_uNames, int(i), 'TB_{:0>2d}'.format(i+1)) )
        tbPsiVals_.append( cleanAppend(_uPsiVals, int(i), 0.01) )
        tbGroups_.append(15)
        tbGeom_.append(None)
        tbrfRSIs_.append(None)
    
    return tbNames_, tbLens_, tbPsiVals_, tbGeom_, tbGroups_, tbrfRSIs_

# Go get the document's TB Library data
tb_lib = getTBLibFromDocText()

# Organize all the Linear TB Parameter data
if len(linear_tb_geom_)==0 and len(linear_tb_lengths_)>0:
    #First, try and get params from GH Doc
    (tbNames,
    tbLengths,
    tbPsiVals,
    tbGeom,
    tbGroups,
    tbfRSIs) = getTBfromGH(linear_tb_names_, linear_tb_lengths_, linear_tb_PsiValues_)
elif len(linear_tb_geom_)>0:
    # Then Try and get params from Rhino Scene
    (tbNames,
    tbLengths,
    tbPsiVals,
    tbGeom,
    tbGroups,
    tbfRSIs) = getTBfromRhino(linear_tb_geom_, tb_lib, linear_tb_names_, linear_tb_lengths_, linear_tb_PsiValues_)

# Put together the new Linear TB Objects
thermalBridges_ = []
if len(linear_tb_geom_)>0 or len(linear_tb_lengths_)>0:
    for i, name in enumerate(tbNames):
        newTBObject = PHPP_ThermalBridge(tbNames[i], tbLengths[i], tbPsiVals[i], tbGeom[i], tbGroups[i], tbfRSIs[i])
        
        thermalBridges_.append(newTBObject)

# Include Point TBs
if len(point_tb_Names_)>0  and len(point_tb_ChiValues_)>0:
    if len(point_tb_Names_) == len(point_tb_ChiValues_):
        for j, point_tb_name in enumerate(point_tb_Names_):
            thermalBridges_.append( PHPP_ThermalBridge(point_tb_name, 1, point_tb_ChiValues_[j], None, 15, None) )
    elif len(point_tb_Names_) != len(point_tb_ChiValues_):
        for j, point_tb_name in enumerate(point_tb_Names_):
            thermalBridges_.append( PHPP_ThermalBridge(point_tb_name, 1, point_tb_ChiValues_[0], None, 15, None) )
elif len(point_tb_Names_)>0  and len(point_tb_ChiValues_)==0:
    for j, point_tb_name in enumerate(point_tb_Names_):
        thermalBridges_.append( PHPP_ThermalBridge(point_tb_name, 1, 0.1, None, 15, None) )

# Inlcude estimated TBs if there are any input
if estimated_tb_:
    if estimated_tb_ > 1:
        estimated_tb_ = float(estimated_tb_) / 100
    thermalBridges_.append( PHPP_ThermalBridge('Estimated', float(estimated_tb_), "=SUM('Annual heating'!O12:O19)/'Annual heating'!M12", None, 15, None)  )

# Preview output
for each in thermalBridges_:
    preview(each)
