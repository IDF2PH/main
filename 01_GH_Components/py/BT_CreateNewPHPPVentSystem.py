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
Collects and organizes data for a simple fresh-air ventilation system (HRV/ERV). Outputs a 'ventilation' class object to apply to a HB Zone.
-
EM Jul. 21, 2020
    Args:
        ventUnitType_: Input Type. Either: "1-Balanced PH ventilation with HR [Default]", "2-Extract air unit", "3-Only window ventilation"
        ventSystemName_: <Optional> A name for the overall system. ie: 'ERV-1', etc.. Will show up in the 'Additional Ventilation' worksheet as the 'Description of Ventilation Units' (E97-E107)
        ventUnit_: Input the HRV/ERV unit object. Connect to the 'ventUnit_' output on the 'Ventilator' Component
        hrvDuct_01_: Input the HRV/ERV Duct object. Connect to the 'hrvDuct_' output on the 'Vent Duct' Component
        hrvDuct_02_: Input the HRV/ERV Duct object. Connect to the 'hrvDuct_' output on the 'Vent Duct' Component
        frostProtectionT_: Min Temp for frost protection to kick in. [deg.  C]. Deffault is -5 C
    Returns:
        ventilation_: A Ventilation Class object with all the params and data related to the simple Ventilation system. Connect this to the '_VentSystem' input on the 'Set Zone Vent' Component.
"""

ghenv.Component.Name = "BT_CreateNewPHPPVentSystem"
ghenv.Component.NickName = "Create Vent System"
ghenv.Component.Message = 'JUL_21_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc

# Classes and Defs
preview = sc.sticky['Preview']
PHPP_Sys_Duct = sc.sticky['PHPP_Sys_Duct']
PHPP_Sys_Ventilation = sc.sticky['PHPP_Sys_Ventilation']

defaultDuct = PHPP_Sys_Duct()

# Create the Ventilation System Object
ventilation_ = PHPP_Sys_Ventilation(
    ventSystemType_ if ventSystemType_ else '1-Balanced PH ventilation with HR', # Unit Type
    ventSystemName_ if ventSystemName_ else 'Vent-1',
    ventUnit_.Name if ventUnit_ else 'Default_Unit',
    ventUnit_.HR_Eff if ventUnit_ else 0.75,
    ventUnit_.MR_Eff if ventUnit_ else 0,
    ventUnit_.Elec_Eff if ventUnit_ else 0.45,
    hrvDuct_01_ if hrvDuct_01_ else defaultDuct, # Duct 1
    hrvDuct_02_ if hrvDuct_02_ else defaultDuct, # Duct 2
    frostProtectionT_ if frostProtectionT_ else -5, # Default -5
    exterior_ if exterior_ else False, # Default is interior
    exhaustVent_)

preview(ventilation_)
