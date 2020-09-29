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
This Component will build out the typical North-American residential appliance set (refrigerator, stove, etc). Note that the default values here will match the PHPP v9.6a and may not be very representative of your specific models / equipment. Refer to your specific equipment for more detailed values to input. Note also that this component will create an appliance set which, by default, is very different than the EnergyPlus values (see the 'PNNL_Resi_Loads' component for detailed E+ values from PNNL sample files.) In many cases, you'll want your PHPP to use the values here for Certification, and then the PNNL values for your EnergyPlus model. You can of course make them both the  same if you want, but usually for Passive House Certification you'll want to use these values here for the PHPP. 
-
This will add the appliances to EACH of the Honeybee zones input. If you only want to add the appliances to one zone or another, use the 'BT_filterZonesByName' component to split up the zones before passing in. Use one of these components for each zone (or each 'type' of zone) to add appliances and things like plug-loads (consumer elec). 
-
Note that this component will ONLY modify the PHPP appliance set, not the EnergyPlus appliance set. In order to  apply these appliances to your EnergyPlus model, use the 'Set PHPP Res Loads' component and Honeybee 'Set Zone Loads'  and 'Set Zone Schedules' components.
- 
EM Sept. 29, 2020
    Args:
        _HBZones: The Honeybee Zones
        _avg_lighting_efficacy: (Lumens / Watt). Avg. Lamp Efficacy, Default=50. Input of the lighting efficiency averaged over all lamps and their duration of use. The lighting efficiency should be reduced by a factor of 0.5 in case of indirect lighting.
Examples for lighting efficiency [lm/W]:
> INCANDESCENT BULB < 25 W: 8 LM/W
> INCANDESCENT BULB < 50 W: 10 LM/W
> INCANDESCENT BULB < 100 W: 12 LM/W
> HALOGEN BULB < 50W : 12 LM/W
> HALOGEN BULB < 100W : 14 LM/W
> COMPACT FLUORESCENT LAMP < 11 W : 50 LM/W
> COMPACT FLUORESCENT LAMP < 20 W : 57 LM/W
> FLUORESCENT LAMP WITH BALLAST: 80 LM/W
> LED RETRO WHITE: 75 LM/W
> LED RETRO WARM WHITE: 65 LM/W
> LED TUBE: 100 LM/W

Note: LED strips may have substantially lower efficiencies!
        dishwasher_kWhUse_: (kWh / use) Default=1.10
        dishwasher_type_: Default='1-DHW connection'. Input either:
> '1-DHW connection'
> '2-Cold water connection'
        clothesWasher_kWhUse_: (kWh / use) Default=1.10. Assume standard 5kg (11 lb) load
        clothesWasher_type_: Default='1-DHW connection'. Input either:
> '1-DHW connection'
> '2-Cold water connection'
        clothesDryer_kWhUse_: (kWh / use) Default=3.50
        clothesDryer_type_: Default='4-Condensation dryer'. Input either:
> '1-Clothes line'
> '2-Drying closet (cold!)
> '3-Drying closet (cold!) in extract air'
> '4-Condensation dryer'
> '5-Electric exhaust air dryer'
> '6-Gas exhaust air dryer'
        refrigerator_kWhDay_: (kWh / day) Default=0.78. Input value here for fridge without a freezer only. For combo units (fridge + freezer) use 'fridgeFreezer_kWhDay_' input instead.
        freezer_kWhDay_: (kWh / day): Default=0.88. Input value here for freezer only. For combo units (fridge + freezer) use 'fridgeFreezer_kWhDay_' input instead.
        fridgeFreezer_kWhDay_: (kWh / day) Defau;t=1.00. Input value here for combo units with fridge + freezer in a single appliance.
        cooking_kWhUse_:  (kWh / use) Typical values include:
> Gas Cooktop: 0.25 kWh/Use.
> Quartz Halogen Ceramic Cooktop: 0.22 kWh/Use
> Induction Ceramic Cooktop: 0.2 kWh/Use
        cooking_Fuel_: Default='1-Electricity'. Input either:
> '1-Electricity'
> '2-Natural gas'
> '3-LPG'
        consumer_elec_: (W) Default=80 W
        other_: (List) A list of up to three additional electricity consuming appliances / equipment (elevators, hot-tubs, etc). Enter each item in the format "Name, kWh/a". So for instance, for an Elevator consuming 600 kWh/a, input the string: "Elevator, 600"
    Returns:
        HBZones_: The Honeybee Zones with new appliances added
"""

ghenv.Component.Name = "BT_SetResAppliances"
ghenv.Component.NickName = "PHPP Res. Appliances"
ghenv.Component.Message = 'SEP_29_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc

hb_hive = sc.sticky["honeybee_Hive"]()
HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)
preview = sc.sticky['Preview']

class elec_equip_appliance():
    defaults = {
        'dishwasher':1.1,
        'clothesWasher':1.1,
        'clothesDryer':3.5,
        'fridge':0.78,
        'freezer':0.88,
        'fridgeFreezer':1.0,
        'cooking':0.25,
        'other_kWhYear_':0.0,
        'consumerElec': 80.0
    }
    
    def __init__(self, _nm=None, _nomDem=None, _utilFac=1, _freq=1, type=None):
        self.Zone = None
        self.Type = 'Appliance'
        self.Name = _nm
        self.NominalDemand = self.getNomDemand(_nomDem)
        self.UtilizationFactor = float(_utilFac) if _utilFac else 1
        self.Frequency = float(_freq) if _freq else 1
        self.Type = type
    
    def getNomDemand(self, _inputVal):
        if _inputVal is None:
            return None
        
        try:
            return float(_inputVal)
        except:
            return self.defaults.get(self.Name, 0)
    
    def calcAnnualDemand(self, _bldgOccupancy=0, _numResUnits=0):
        demandByOccupancy = self.NominalDemand * self.UtilizationFactor * self.Frequency * _bldgOccupancy 
        demandByHousehold = self.NominalDemand * self.UtilizationFactor * self.Frequency * _numResUnits
        
        if self.Name == 'fridge' or self.Name == 'freezer' or self.Name == 'fridgeFreezer':
            return demandByHousehold
        elif self.Name == 'other':
            return self.NominalDemand
        else:
            return demandByOccupancy
    
    def __unicode__(self):
        return u'PHPP Style Appliance: {}'.format(self.Name)
    def __str__(self):
        return unicode(self).encode('utf-8')
    def __repr__(self):
        return "{}( _nm={!r}, _nomDem={!r}, _utilFac={!r}, _freq={!r}, type={!r})".format(
                self.__class__.__name__,
                self.Name,
                self.NominalDemand,
                self.UtilizationFactor,
                self.Frequency,
                self.Type)

def is_input(_input):
    if _input is None: return False
    else: return True

def clean_dw_type(_in):
    if '2' in str(_in):
        return '2-Cold water connection'
    else:
        return '1-DHW connection'

def clean_wsh_type(_in):
    if '2' in str(_in):
        return '2-Cold water connection'
    else:
        return '1-DHW connection'

def clean_dryer_type(_in):
    if '1' in str(_in):
        return '1-Clothes line'
    if '2' in str(_in):
        return '2-Drying closet (cold!)'
    elif '3' in str(_in):
        return '3-Drying closet (cold!) in extract air'
    elif '5' in str(_in):
        return '5-Electric exhaust air dryer'
    elif '6' in str(_in):
        return '6-Gas exhaust air dryer'
    else:
        return '4-Condensation dryer'

def clean_cooking_type(_in):
    if '2' in str(_in):
        return '2-Natural gas'
    elif '3' in str(_in):
        return '3-LPG'
    else:
        return '1-Electricity'

#-------------------------------------------------------------------------------
# Create Appliances, always include Consumer Electronics
appliances = []
if is_input(dishwasher_kWhUse_):    appliances.append( elec_equip_appliance('dishwasher', dishwasher_kWhUse_ , 1 , 65, type=clean_dw_type(dishwasher_type_) )) 
if is_input(clothesWasher_kWhUse_): appliances.append( elec_equip_appliance('clothesWasher', clothesWasher_kWhUse_ , 1 , 57, type=clean_wsh_type(clothesWasher_type_) ))
if is_input(clothesDryer_kWhUse_):  appliances.append( elec_equip_appliance('clothesDryer', clothesDryer_kWhUse_ , 0.8750 , 57, type=clean_dryer_type(clothesDryer_type_) ))
if is_input(refrigerator_kWhDay_):  appliances.append( elec_equip_appliance('fridge', refrigerator_kWhDay_, 1 , 365 ))
if is_input(freezer_kWhDay_):       appliances.append( elec_equip_appliance('freezer', freezer_kWhDay_ , 1 , 365 ) )
if is_input(fridgeFreezer_kWhDay_): appliances.append( elec_equip_appliance('fridgeFreezer', fridgeFreezer_kWhDay_ , 1 , 365 ))
if is_input(cooking_kWhUse_):       appliances.append( elec_equip_appliance('cooking', cooking_kWhUse_ , 1 , 500, type=clean_cooking_type(cooking_Fuel_) ))

if is_input(other_):        
    for each in other_:
        name, demand = each.split(',') 
        appliances.append( elec_equip_appliance('ud__'+name, demand ))

if is_input(consumer_elec_):
    appliances.append( elec_equip_appliance('consumerElec', consumer_elec_ , 1 , 0.55 ))
else:
    appliances.append( elec_equip_appliance('consumerElec', 80 , 1 , 0.55 ))

#-------------------------------------------------------------------------------
# Add appliances to HB Zone Objects
for zone in HBZoneObjects:
    for appliance in appliances:
        appliance.Zone = zone.name
    
    try:
        zone_appliance_list = getattr(zone, 'PHPP_ElecEquip')
        zone_appliance_list + appliances
        
        setattr(zone, 'PHPP_ElecEquip', zone_appliance_list)
    except Exception as err:
        print err, '::: Adding a new appliance list to the zone'
        setattr(zone, 'PHPP_ElecEquip', appliances)

#-------------------------------------------------------------------------------
#Lighting Efficacy, Add Lighting to HB Zone Objects
lighting_Efficacy = _avg_lighting_efficacy if _avg_lighting_efficacy else 50
for zone in HBZoneObjects:
    lightingObj = elec_equip_appliance('lighting', lighting_Efficacy)
    setattr(lightingObj, 'Zone', zone.name)
    setattr(zone, 'PHPP_LightingEfficacy', lightingObj)

#-------------------------------------------------------------------------------
# Add modified zones back to the HB dictionary
if len(_HBZones)>0:
    HBZones_  = hb_hive.addToHoneybeeHive(HBZoneObjects, ghenv.Component)
    for zone in HBZoneObjects:
        for appliance in zone.PHPP_ElecEquip:
            preview(appliance)
else:
    HBZones_ = _HBZones