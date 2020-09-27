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
Set the control options for a Air- or Water-Source Heat Pump (HP). Sets the values on the 'HP' worksheet.
-
EM September 27, 2020
    Args:
        hp_distribution_: The Heat Pump heating distribution method. Input either:
> "1-Underfloor heating"
> "2-Radiators"
> "3-Supply air heating"
        nom_power_: (kW) Only to be entered in special cases. Most users leave blank.
        rad_exponent_: Only to the entered in special cases. Most users leave blank.
        backup_type_: Type of backup heater. Default is "1-Elec Immersion heater". Enter either:
> "1-Elec. Immersion heater"
> "2-Electric continuous flow water heater"
        dT_elec_flow_water_: (Deg C) delta-Temp of electric continuous flow water heater
        hp_priority_: Heat Pump priority. Default is "1-DHW-priority". Enter either:
> "1-DHW-priority"
> "2-Heating priority"
        hp_control_: Default is "1-On/Off". Enter either:
> "1-On/Off"
> "2-Ideal"
        depth_groundwater_: (m) Depth ground water / Ground collector / Ground probe.
        power_groundpump_: (kW) Power of pump for ground heat exchanger.
    Returns:
        hpOptions_: A Heat Pump object. Connect this output to the 'hpOptions_' input on a 'Heating/Cooling System' component.
"""

ghenv.Component.Name = "BT_Heating_ASHP_Options"
ghenv.Component.NickName = "Heating | HP Options"
ghenv.Component.Message = 'SEP_27_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc
import Grasshopper.Kernel as ghK

# Classes and Defs
preview = sc.sticky['Preview']

#-------------------------------------------------------------------------------
class HP_Options:
    def __init__(self, _dist='3-Supply air heating', _pow=None, _rExp=None, 
                _backup='1-Elec. Immersion heater', _dtWater=None, _pior='1-DHW-priority', 
                _c='1-On/Off', _dpth=None, _pumpPow=None):
        self.Distribution = _dist
        self.NominalPower = _pow
        self.RadExponent = _rExp
        self.BackupType = _backup
        self.ElecFlowWater_dT = _dtWater
        self.Priority = _pior
        self.Control = _c
        self.GroundWaterDepth = _dpth
        self.GroundPumpPower = _pumpPow
    
    def __unicode__(self):
        return u'A Heat Pump Options Object'
    def __str__(self):
        return unicode(self).encode('utf-8')
    def __repr__(self):
        return "{}( _dist={!r})".format(
               self.__class__.__name__,
               self.Distribution)

def clean_dist(_in):
    input_str = str(_in)
    if '1' in input_str:
        return '1-Underfloor heating'
    elif '2' in input_str:
        return '2-Radiators'
    else:
        return '3-Supply air heating'
    
    return _in

def clean_backup(_in):
    input_str = str(_in)
    if '2' in input_str:
        return '2-Electric continuous flow water heater'
    else:
        return '1-Elec. Immersion heater'
    
    return _in

def clean_priority(_in):
    input_str = str(_in)
    if '2' in input_str:
        return '2-Heating priority'
    else:
        return '1-DHW-priority'
    
    return _in

def clean_control(_in):
    input_str = str(_in)
    if '2' in input_str:
        return '2-Ideal'
    else:
        return '1-On/Off'
    
    return _in

def checkInput(_in, _nm):
    if _in is None: return None
    
    try:
        return float(_in)
    except:
        msg = 'Error: Input for {} should be a number only.'.format(_nm)
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning ,msg)
        return None

#-------------------------------------------------------------------------------
hp_distribution = clean_dist(hp_distribution_)
backup = clean_backup(backup_type_)
hp_priority = clean_priority(hp_priority_)
hp_control = clean_control(hp_control_)

#-------------------------------------------------------------------------------
hpOptions_ = HP_Options(
                    hp_distribution,
                    checkInput(nom_power_, 'nom_power_'),
                    checkInput(rad_exponent_, 'rad_exponent_'),
                    backup,
                    checkInput(dT_elec_flow_water_, 'dT_elec_flow_water_'),
                    hp_priority,
                    hp_control,
                    checkInput(depth_groundwater_, 'depth_groundwater_'),
                    checkInput(power_groundpump_, 'power_groundpump_'))

preview(hpOptions_)