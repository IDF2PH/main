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
Collects and organizes data for a Ventilator Unit (HRV/ERV). Used to build up a PHPP-Style Ventilation System.
-
EM June. 26, 2020
    Args:
        name_: <Optional> The name of the Ventilator Unit
        HR_Eff_: <Optional> Input the Ventialtion Unit's Heat Recovery %. Default is 75% 
        MR_Eff_: <Optional> Input the Ventialtion Unit's Moisture Recovery %. Default is 0% (HRV)
        Elec_Eff_: <Optional> Input the Electrical Efficiency of the Ventialtion Unit (W/m3h). Default is 0.45 W/m3h
    Returns:
        ventUnit_: A Ventilator object for the Ventilation System. Connect to the 'ventUnit_' input on the 'Create Vent System' to build a PHPP-Style Ventialtion System.
"""

ghenv.Component.Name = "BT_CreateNewPHPPVentUnit"
ghenv.Component.NickName = "Ventilator"
ghenv.Component.Message = 'JUN_26_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc

# Classes and Defs
preview = sc.sticky['Preview']
PHPP_Sys_VentUnit = sc.sticky['PHPP_Sys_VentUnit']

def checkInput(_in):
    if float(_in) > 1:
        return float(_in) / 100
    else:
        return float(_in)

ventUnit_ = PHPP_Sys_VentUnit(name_ if name_ else 'Default_Unit',
    checkInput(HR_Eff_) if HR_Eff_ else 75,
    checkInput(MR_Eff_) if MR_Eff_ else 0,
    Elec_Eff_ if Elec_Eff_ else 0.45)

preview(ventUnit_)
