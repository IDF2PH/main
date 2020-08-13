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
Used to set up any 'Variants' desired in the output PHPP. The 'Variants' worksheet can contol the PHPP all from a single dashboard worksheet. This is useful when doing iterative studies and examining the impacts of varying assembly insulation levels, window products, etc.. To use, set the desired categories to 'True' to have this setup the PHPP to refer to the 'Variants' worksheet.
-
Note: You'll have to set all the variants parameters in that worksheet either before or after writing out to the PHPP here. Best practice is to set up the 'source' PHPP with all the variant schemes and parameters desired before writing.
-
Note: For the ventilation input, use 'True' to set the ventilation system for a default PHPP. If you are using 
a modified PHPP (XL, added rows to the 'Additional Vent' worksheet, etc..) you can alternately input a list of 
values here in place of a boolean value. These values should correspond to to the:
    Ventilation Unit
    Duct Insualtion Thickness
    Duct Length

ie: if you have modified the number of rows on the 'Additional Vent' worksheetsuch that the Ventilation unit inputs now start on row 141 and the ducting inputs now start on row 171, enter:
    Additional Vent!F141=Variants!D856
    Additional Vent!H171=Variants!D858
    Additional Vent!H172=Variants!D858
    Additional Vent!L171=Variants!D857
    Additional Vent!L172=Variants!D857
in order to set the variant control for any a typical 2-duct ventilation system.
-
EM Jul. 21, 2020

    Args:
        windows_: (bool) True = Set up the windows to refer to the 'Variants' worksheet
        uValues_: (bool) True = Set up the 'Areas' to refer to the 'Variants' worksheet
        airtightness_: (bool) True = Set airtightness (n50) value to refer to the 'Variants' worksheet
        ventilation_: (bool) True = Set Ventilation system to refer to the 'Variants' worksheet
        thermalBridges_: (bool) True = Set Thermal Bridge allocation (% value) to refer to the 'Variants' worksheet
        certification_: (bool) True = Set Certification options on the 'Verification' worksheet to refer to the 'Variants' worksheet
        primaryEnergy_: (bool) True = Set PER worksheet heating, DHW and Cooling system designations to refer to the 'Variants' worksheet
    Returns:
        variants_: A DataTree of the final clean, Excel-Ready output objects. Each output object has a Worksheet-Name, a Cell Range, and a Value. Connect to the 'Variants_' input on the 'PHPP | 2XL Write' Component to write out to Excel.
"""

ghenv.Component.Name = "BT_SetVariants"
ghenv.Component.NickName = "Variants"
ghenv.Component.Message = 'JUL_21_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "02 | IDF2PHPP"

from System import Object
from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path
import Grasshopper.Kernel as ghK
import scriptcontext as sc

# Classes and Defs
PHPP_XL_Obj = sc.sticky['PHPP_XL_Obj'] 
preview = sc.sticky['Preview']

variants_ = DataTree[Object]()

paths = {'vent':1, 'uVals':6, 'air':0, 'tb':2, 'cert':3, 'win':4, 'per':5}

if windows_:
    for i in range(24, 175):
        variants_.Add(  PHPP_XL_Obj('Windows', 'T{}'.format(i), '=G{}'.format(i) ), GH_Path( paths['win'] )  )
        variants_.Add(  PHPP_XL_Obj('Windows', 'U{}'.format(i), '=H{}'.format(i) ), GH_Path( paths['win'] )  )

if uValues_:
    for i in range(0, 10):
        row_Uval = 17+i*21
        row_Variant = 410+i*2
        row_Compo = 15+i
        
        variants_.Add(  PHPP_XL_Obj('U-Values', 'M'+str(row_Uval), '=F'+str(row_Uval) ), GH_Path(paths['uVals'])  )
        variants_.Add(  PHPP_XL_Obj('U-Values', 'S'+str(row_Uval), '=G'+str(row_Uval) ), GH_Path(paths['uVals'])  )
        variants_.Add(  PHPP_XL_Obj('Variants', 'B'+str(row_Variant), '=Components!D'+str(row_Compo) ), GH_Path(paths['uVals'])  )

if airtightness_:
    variants_.Add(  PHPP_XL_Obj('Ventilation', 'N27', '=D27' ), GH_Path( paths['air'] )  )

if len(ventilation_)==1:
    variants_.Add(  PHPP_XL_Obj('Ventilation', 'L12', '=D12' ), GH_Path(100)  )
    variants_.Add(  PHPP_XL_Obj('Additional Vent', 'F97', '=Variants!D856' ), GH_Path(paths['vent'])  )
    variants_.Add(  PHPP_XL_Obj('Additional Vent', 'H127', '=Variants!D858' ), GH_Path(paths['vent'])  )
    variants_.Add(  PHPP_XL_Obj('Additional Vent', 'H128', '=Variants!D858' ), GH_Path(paths['vent'])  )
    variants_.Add(  PHPP_XL_Obj('Additional Vent', 'L127', '=Variants!D857' ), GH_Path(paths['vent'])  )
    variants_.Add(  PHPP_XL_Obj('Additional Vent', 'L128', '=Variants!D857' ), GH_Path(paths['vent'])  )
elif len(ventilation_)==5:
    variants_.Add(  PHPP_XL_Obj('Ventilation', 'L12', '=D12' ), GH_Path(100)  )
    for each in ventilation_:
        first, reference = each.split('=')
        wrksht, rng = first.split('!')
        variants_.Add(  PHPP_XL_Obj(wrksht, rng, '='+reference ), GH_Path(paths['vent'])  )
elif len(ventilation_)!=0:
    msg1 = "Error. ventialtion_ input not understood? Either input TRUE to use the defaults\n"\
    "or input multiline line string with the excel formula to write? Multine string format should look like:\n"\
    "------------\n"\
    "   Additional Vent!F{Your Row Number}=Variants!D856\n"\
    "   Additional Vent!H{Your Row Number}=Variants!D858\n"\
    "   Additional Vent!H{Your Row Number}=Variants!D858\n"\
    "   Additional Vent!L{Your Row Number}=Variants!D857\n"\
    "   Additional Vent!L{Your Row Number}=Variants!D857\n"
    ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg1)

if thermalBridges_:
    variants_.Add(  PHPP_XL_Obj('Areas', 'R145', '=Variants!D933' ), GH_Path(paths['tb'])  )

if certification_:
    variants_.Add(  PHPP_XL_Obj('Verification', 'R78', '=Variants!D927' ), GH_Path(paths['cert'])  )
    variants_.Add(  PHPP_XL_Obj('Verification', 'R80', '=Variants!D928' ), GH_Path(paths['cert'])  )
    variants_.Add(  PHPP_XL_Obj('Verification', 'R82', '=Variants!D929' ), GH_Path(paths['cert'])  )
    variants_.Add(  PHPP_XL_Obj('Verification', 'R85', '=Variants!D930' ), GH_Path(paths['cert'])  )
    variants_.Add(  PHPP_XL_Obj('Verification', 'R87', '=Variants!D931' ), GH_Path(paths['cert'])  )
    
if primaryEnergy_:
    variants_.Add(  PHPP_XL_Obj('PER', 'P10', '=H10' ), GH_Path(paths['per'])  )
    variants_.Add(  PHPP_XL_Obj('PER', 'P12', '=H12' ), GH_Path(paths['per'])  )
    variants_.Add(  PHPP_XL_Obj('PER', 'S10', '=I10' ), GH_Path(paths['per'])  )
    variants_.Add(  PHPP_XL_Obj('PER', 'T10', '=J10' ), GH_Path(paths['per'])  )

