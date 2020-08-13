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
Set the parameters for a Supply Air Cooling element (HRV). Sets the values on the 'Cooling Unit' worksheet.
-
EM August 11, 2020
    Args:
        on_offMode_: ('x' or '') Default=''. Cyclical operation works through an on/off regulation of the compressor. If this field is left empty, then the assumption is that the unit has a VRF (variant refrigerant flow), which works by modulating the efficiency of the compressor.
        maxCoolingCap_: (kW) Default = 1000
        SEER_: (W/W) Default = 3
- bad devices: 2.5
- split units, energy efficiency class A: > 3.2
- compact units, energy efficiency class A: > 3.0
- turbo compressor > 500 kW with water cooling: up to more than 6
    Returns:
        supplyAirCooling_: 
"""

ghenv.Component.Name = "BT_CreateCooling_SupplyAir"
ghenv.Component.NickName = "Cooling | Sup Air"
ghenv.Component.Message = 'AUG_11_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc
import Grasshopper.Kernel as ghK

# Classes and Defs
preview = sc.sticky['Preview']

class Cooling_SupplyAir:
    def __init__(self,  _onOff='', _power=1000, _seer=3):
        self.OnOff = _onOff if _onOff != None else ''
        self.MaxPower =  _power if _power != None else 1000
        self.SEER = _seer if _seer != None else 3
    
    def getValsForPHPP(self):
        return self.OnOff, self.MaxPower, self.SEER
    
    def __repr__(self):
        return "A Cooling: Supply Air Params Object"

def cleanInputs(_in, _nm, _default):
    # Apply defaults if the inputs are Nones
    out = _in if _in != None else _default
    
    if _nm == 'onOff':
        if out in ['x', '']:
            return out
        else:
            msg = "on_offModel_ input should be either 'x' or '' (blank) only. 'x' means the equipment only has on/off mode.\n"\
            "Leave blank if equipment has variable variable flow such as VRF."
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
            return _default
    
    try:
        return float(out)
    except:
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, '"{}" input should be a number'.format(_nm))
        return _default

onOff = cleanInputs(on_offMode_, 'onOff', '')
maxCooling = cleanInputs(maxCoolingCap_, 'maxCoolingCap_', 1000)
seer = cleanInputs(SEER_, 'SEER_', 3)

supplyAirCooling_ = Cooling_SupplyAir(onOff, maxCooling,  seer)

preview(supplyAirCooling_)
