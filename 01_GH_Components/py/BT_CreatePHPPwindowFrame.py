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
Will combine parameters together to create a new Window Frame object which can be used by the 'Create Windows' component. A PHPP-Style frame includes very detailed information.
-
EM Feb. 25, 2020
    Args:
        _name: The name for the new Window Frame object
        _frameWidths: (List) Input values for the frame face-widths (m). Input values in order: Left, Right, Bottom, Top. If less than 4 values are input, the first value will be used for all four edges.
        _frameUvalues: (List) Input values for the frame U-Values (W/m2-k). Frame U-values to be calculated as per ISO-10077-2. Input values in order: Left, Right, Bottom, Top. If less than 4 values are input, the first value will be used for all four edges.
        _psiGlazings: (List) Input values for the Psi-GlazingEdge Values (W/m-k). Frame Psi-GlazingEdge values to be calculated as per ISO-10077-2. Input values in order: Left, Right, Bottom, Top. If less than 4 values are input, the first value will be used for all four edges.
        _psiInstalls: (List) Input values for the Psi-Install Values (W/m-k). Frame Psi-Install values to be calculated as per ISO 10211. Input values in order: Left, Right, Bottom, Top. If less than 4 values are input, the first value will be used for all four edges.
    Returns:
        PHPPFrame_: A new PHPP-Style Frame Object for use in a PHPP Window. Connect to the BT_CreatePHPPWindow Component.
"""

ghenv.Component.Name = "BT_CreatePHPPwindowFrame"
ghenv.Component.NickName = "New PHPP Frame"
ghenv.Component.Message = 'FEB_25_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import rhinoscriptsyntax as rs
import scriptcontext as sc

def cleanInputs(_inputList, defaultValue):
    # Clean up installs
    if len(_inputList)==4:
        return _inputList
    elif len(_inputList)==1:
        return [_inputList[0], _inputList[0], _inputList[0], _inputList[0]]
    else:
        return [defaultValue, defaultValue, defaultValue, defaultValue]

PHPP_Frame = sc.sticky['PHPP_Frame']

if _name:
    # Clean up the inputs
    _frameWidths= cleanInputs(_frameWidths, 0.12)
    _frameUvalues= cleanInputs(_frameUvalues, 1.0)
    _psiGlazings = cleanInputs(_psiGlazings, 0.04)
    _psiInstalls = cleanInputs(_psiInstalls, 0.04)
    _name = _name.replace(" ", "_")
    
    PHPPFrame_ = PHPP_Frame(
                        _name if _name else 'Unamed_Frame',
                        _frameUvalues,
                        _frameWidths,
                        _psiGlazings, 
                        _psiInstalls
                        )
