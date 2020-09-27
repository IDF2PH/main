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
Collects and organizes data for simple heating / cooling equipment. Outputs a 'Heating_Cooling' and 'PER' class object with all the data organized and keyed.
-
EM September 27, 2020
    Args:
        primaryHeatGeneration_: <Optional, Default='5-Direct electricity'> The type of heating equipment to use for the 'Primary' heating. Input either:
>  1-HP compact unit
>  2-Heat pump(s)
>  3-District heating, CGS
>  4-Heating boiler
>  5-Direct electricity
>  6-Other
        secondaryHeatGeneration_: <Optional, Default='-'> The type of heating equipment to use for the 'Secondary' heating. Input either:
>  1-HP compact unit
>  2-Heat pump(s)
>  3-District heating, CGS
>  4-Heating boiler
>  5-Direct electricity
>  6-Other
Leave blank if no secondary heat generation equipment is used.
        heatingFracPrimary_: <Optional, Default=100%> A number from 0-1. The percentage of Heating energy that comes from the 'Primary' heater. The rest will come from the 'Secondary' heater.
        dhwFracPrimary_: <Optional, Default=100%> A number from 0-1. The percentage of DHW Energy that comes from the 'Primary' heater. The rest will come from the 'Secondary' heater.
        boiler_: <Optional, Default=None> A typical hot water Boiler system parameters. Connect to the 'Heating:Boiler' Component.
        hp_heating_: <Optional, Default=None> Heat Pump unit for space heating
        hg_DHW_: <Optional, Default=None> Heat Pump unit for Domestic Hot Water (DHW)
        hpGround_: <Optional, Default=None> Not implemented yet....
        compact_: <Optional, Default=None> Not implemented yet....
        distictHeating_: <Optional, Default=None> Not implemented yet....
        supplyAirCooling_: <Optional, Default=None> A 'Supply Air' cooling system. Providing cooling through the ventilation (HRV) air flow.
        recircAirCooling_: <Optional, Default=None> A normal AC system parameters. Connect to the 'Cooling:Recirc' Component.
        addnlDehumid_: <Optional, Default=None> An Additional Dehumidification element. 
        panelCooling_: <Optional, Default=None> Not implemented yet....
    Returns:
        Heating_Cooling_: DataTree. Branch 1 is all the 'PER' related information (heating fractions). Branch 2 is all the Heating / Cooling Equipment related data. Connect to the 'Heating_Cooling_' input on the '2PHPP | Setup' Component.
"""

ghenv.Component.Name = "BT_SetSystems"
ghenv.Component.NickName = "Heating/Cooling Systems"
ghenv.Component.Message = 'SEP_27_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "02 | IDF2PHPP"


from System import Object
from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path
import Grasshopper.Kernel as ghK
import scriptcontext as sc
import re

# Classes and Defs
preview = sc.sticky['Preview']

class PER:
    def __init__(self, _heatPrimary, _heatSecondary, _heatFracPrim, _dhwFracPrim):
        self.heatPrimaryGen = _heatPrimary
        self.heatScondaryGen = _heatSecondary
        self.heatFracPrimary = float(_heatFracPrim)
        self.dhwFracPrrimary = float(_dhwFracPrim)
        
    def __repr__(self):
        str = ("A PER Params Object with:",
                ' > Pimary Heat Generator: {}'.format(self.heatPrimaryGen),
                ' > Secondary Heat Generator: {}'.format(self.heatScondaryGen),
                ' > Heat % From Primary: {:.0f}%'.format(self.heatFracPrimary*100),
                ' > DHW % From Primary: {:.0f}%'.format(self.dhwFracPrrimary*100)
                )
                
        return "\n".join(str)

class HeatingCoolingEquip:
    def __init__(self, _boiler, _hp, _hpDHW, _hpG, _hPOpt, _comp, _dheat, _supC, _recirc, _dehumid, _panelC):
        self.Boiler = _boiler
        self.HP_heating = _hp
        self.HP_dhw = _hpDHW
        self.HP_Ground = _hpG
        self.HP_Options = _hPOpt
        self.CompactUnit = _comp
        self.DistrictHeating = _dheat
        self.SupplyAirCooling = _supC
        self.RecircCooling = _recirc
        self.AddnlDehumid = _dehumid
        self.PanelCooling = _panelC
    
    def __repr__(self):
        str = ("Heating / Cooling Equipment Params",
                ' > Boiler: {}'.format(self.Boiler),
                ' > HP: | Heating {}'.format(self.HP_heating),
                ' > HP | Hot Water: {}'.format(self.HP_dhw),
                ' > GSHP: {}'.format(self.HP_Ground),
                ' > Compact Unit: {}'.format(self.CompactUnit),
                ' > District Heating: {}'.format(self.DistrictHeating),
                ' > Supply Air Cooling: {}'.format(self.SupplyAirCooling),
                ' > Recirc Cooling: {}'.format(self.RecircCooling),
                ' > Additional Dehumidification: {}'.format(self.AddnlDehumid),
                ' > Panel Cooling: {}'.format(self.PanelCooling),
                )
        
        return '\n'.join(str)

def cleanHeatGenInput(_inputVal, _default):
    options = {
        '1':'1-HP compact unit',
        '2':'2-Heat pump(s)',
        '3':'3-District heating, CGS',
        '4':'4-Heating boiler',
        '5':'5-Direct electricity',
        '6':'6-Other'}
    
    if _inputVal == None:
        return _default
    
    match =re.match('\d',str(_inputVal).lstrip())
    if match:
        selection = options.get(match.group(), None)
    else:
        selection = None
    
    if selection:
        return selection
    else:
        msg = 'Please enter a valid Heat Generation type. Select either:\n'\
        '>  1-HP compact unit\n'\
        '>  2-Heat pump(s)\n'\
        '>  3-District heating, CGS\n'\
        '>  4-Heating boiler\n'\
        '>  5-Direct electricity\n'\
        '>  6-Other'
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)

def cleanFractionInput(_input):
    if _input == None:
        return 1
    
    try:
        if float(_input) > 1.0:
            return float(_input)/100
        else:
            return float(_input)
    except:
        msg = 'Please enter a valid heating Fraction input. Enter a number from 0-100%'
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
        return 1



## PER Items
PER_ = PER(cleanHeatGenInput(primaryHeatGeneration_, '5-Direct electricity'), 
            cleanHeatGenInput(secondaryHeatGeneration_, '-'),
            cleanFractionInput(heatingFracPrimary_),
            cleanFractionInput(dhwFracPrimary_)
            )

## Heating / Cooling Equipment
Equipment_ = HeatingCoolingEquip(boiler_,
            hp_heating_,
            hp_DHW_,
            hpGround_,
            hpOptions_,
            compact_,
            distictHeating_,
            supplyAirCooling_,
            recircAirCooling_,
            addnlDehumid_,
            panelCooling_
            )

# Add both PER and Equipment to a master tree
Heating_Cooling_ = DataTree[Object]()
Heating_Cooling_.Add(PER_, GH_Path(0))
Heating_Cooling_.Add(Equipment_, GH_Path(01))

print('----\nPER:')
preview(PER_)
print('----\nEquipment:')
preview(Equipment_)