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
Collects and organizes data for an Exhaust Ventilaiton Object such as a Kitchen hood, fireplace makeup air, dryer, etc.
For guidance on modeling and design of exhaust air systems, see PHI's guidebook:
"https://passiv.de/downloads/05_extractor_hoods_guideline.pdf"
-
EM Jul. 21, 2020
    Args:
        _name: (String) A Name to describe the Exhaust Vent item. This will be the name given in the 'Additional
        airFlowRate_On: <Optional> (Int) The airflow (m3/h) of the exhaust ventilation devide when ON.
- PHI Recommened values: 'Efficient Equip.'=250m3/h, 'Standard Equip.'=450m3/h
- Note: if you prefer, add 'cfm' to the input in order to supply values in IP units. It'll convert this to m3/h before passing along to the PHPP
        airFlowRate_Off: <Optional> (Int) The airflow (m3/h) of the exhaust ventilation devide when OFF.
- PHI Recommened values: 'Efficient Equip.'=2m3/h, 'Standard Equip.'=10m3/h
- Note: if you prefer, add 'cfm' to the input in order to supply values in IP units. It'll convert this to m3/h before passing along to the PHPP
        hrsPerDay_On: <Optional> (Float) The hours per day that the device operates. For kitchen extract, assume 0.5 hours / day standard.
        daysPerWeek_On: <Optional> (Int) The days per weel that the devide is used. For kitchen extract, assume 7 days / week standard.
    Returns:
        hrvDuct_: An Exhaust Ventilation Object. Input this into the 'exhaustVent_' input on a 'Create Vent System' component.
"""

ghenv.Component.Name = "BT_CreateNewPHPPExhaustVentilator"
ghenv.Component.NickName = "Exhaust Vent"
ghenv.Component.Message = 'JUL_21_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc

# Classes and Defs
preview = sc.sticky['Preview']
PHPP_Sys_Duct = sc.sticky['PHPP_Sys_Duct']
PHPP_Sys_ExhaustVent = sc.sticky['PHPP_Sys_ExhaustVent'] 

defaultDuct = PHPP_Sys_Duct(1, 152.4, 52, 0.04)
exhaustVent_ = PHPP_Sys_ExhaustVent(_name, airFlowRate_On, airFlowRate_Off, hrsPerDay_On, daysPerWeek_On, defaultDuct)

preview(exhaustVent_)
