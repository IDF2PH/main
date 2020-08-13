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
Will combine parameters together to create a new Window Glazing object which can be used by the 'Create Windows' component.
-
EM Feb. 25, 2020
    Args:
        _name: The name for the new Window Glazing object
        _uValue: (List) Input value for the glazing U-Value (W/m2-k). Glass U-Value to be calculate as per EN-673
        _gValue: (List) Input value for the glazing g-Value (SHGC) . Glass g-Value (SHGC) to be calculate as per EN-410
    Return:
        PHPPGlazing_: A new PHPP-Style Glazing Object for use in a PHPP Window. Connect to the BT_CreatePHPPWindow Component.
"""

ghenv.Component.Name = "BT_CreatePHPPwindowGlazing"
ghenv.Component.NickName = "New PHPP Glazing"
ghenv.Component.Message = 'FEB_25_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import rhinoscriptsyntax as rs
import scriptcontext as sc

PHPP_Glazing = sc.sticky['PHPP_Glazing']

if _name:
    _name = _name.replace(" ", "_")
    
    PHPPGlazing_ = PHPP_Glazing(
                        _name if _name else 'Unnamed_Glazing',
                        float(_gValue) if _gValue else 0.4,
                        float(_uValue) if _uValue else 1.0
                        )


