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
The HW usage profile for the project. Input False to turn 'off' any usage type.
Input a DHW Usage Type into _useType. Input either:
        1-Residential
        2-Non-Residential
-
EM June. 27, 2020
    Args:
        _useType:
    Returns:
        usage_: The usage object. Connect this to the 'usage_' input on the 'DHW System' component in order to pass along to the PHPP.
"""

ghenv.Component.Name = "BT_DHW_Usage"
ghenv.Component.NickName = "DHW Usage"
ghenv.Component.Message = 'JUN_27_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc
import Grasshopper.Kernel as ghK
from collections import defaultdict

PHPP_DHW_usage = sc.sticky['PHPP_DHW_usage'] 
PHPP_DHW_usage_NonRes = sc.sticky['PHPP_DHW_usage_NonRes']
preview = sc.sticky['Preview']

def setupInputs(_type):
    
    direction = 'Please input a value DHW Usage Type into _useType. Input either:\n'\
    "    1-Residential\n"\
    "    2-Non-Residential"
    
    #Setup the input names
    inputsRes = {
            1:{ 'name': 'showers_demand_', 'desc':'(Litre/Person/Day) HW usage for showers only at Design-Forward-Temp (un-mixed). Default is 16 ltrs/pers/day.'},
            2: {'name':'other_demand_', 'desc':'(Litre/Person/Day) HW usage for all non-shower use at Design-Forward-Temp (un-mixed). Default is 9 ltrs/pers/day.'}
            }
    
    inputsNonRes = {
        1:{ 'name': 'useDaysPerYear_', 'desc':'(int) Number of Days per Year the hot water is used.'},
        2:{ 'name': 'showers_', 'desc':'(bool) True=Used, False=Not Used'},
        3:{ 'name': 'handWashBasin_', 'desc':'(bool) True=Used, False=Not Used'},
        4:{ 'name': 'washStand_', 'desc':'(bool) True=Used, False=Not Used'},
        5:{ 'name': 'bidet_', 'desc':'(bool) True=Used, False=Not Used'},
        6:{ 'name': 'bathing_', 'desc':'(bool) True=Used, False=Not Used'},
        7:{ 'name': 'toothBrushing_', 'desc':'(bool) True=Used, False=Not Used'},
        8:{ 'name': 'cookingAndDrinking_', 'desc':'(bool) True=Used, False=Not Used'},
        9:{ 'name': 'dishwashing_', 'desc':'(bool) True=Used, False=Not Used'},
        10:{ 'name': 'cleaningKitchen_', 'desc':'(bool) True=Used, False=Not Used'},
        11:{ 'name': 'cleaningRooms_', 'desc':'(bool) True=Used, False=Not Used'}
        }
    
    # Decide which input style to use
    if _type == None:
        inputs = {}
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, direction)
    elif '1' in str(_type):
        inputs = inputsRes
        
    elif '2' in str(_type):
        inputs = inputsNonRes
    else:
        inputs = {}
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, direction)
    
    #Set up the Component Inputs
    for inputNum in range(1, 12):
        item = inputs.get(inputNum, {'name':'-', 'desc':'-'})
        
        ghenv.Component.Params.Input[inputNum].NickName = item.get('name')
        ghenv.Component.Params.Input[inputNum].Name = item.get('name')
        ghenv.Component.Params.Input[inputNum].Description = item.get('desc')
        
    return inputs

# Set the inputs
inputNames = setupInputs(_useType)
ghenv.Component.Attributes.Owner.OnPingDocument()

if _useType == None:
    _useType = ''

if '1' in _useType:
    # Build Residential system    
    defaultInputs = defaultdict()
    
    for input in ghenv.Component.Params.Input:
        for inputValue in input.VolatileData.AllData(True): # Yields and Enum
            defaultInputs[input.Name] = str(inputValue)
    
    usage_ = PHPP_DHW_usage(defaultInputs)
elif '2' in _useType:
    # Build Non-Residential System
    defaultInputs = defaultdict()
    
    for input in ghenv.Component.Params.Input:
        for inputValue in input.VolatileData.AllData(True): # Yields and Enum
            defaultInputs[input.Name] = str(inputValue)
    usage_ = PHPP_DHW_usage_NonRes(defaultInputs)
else:
    pass

if usage_:
    preview(usage_)
