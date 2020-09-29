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
This Component is used to calculate and apply PHPP-Style loads to the building. Note that this component should used at the very end and ALL HB-Zones should be input merged into a single list. It needs to figure out the total building floor area to calc the occupancy, so can't use this on one zone at a time. 
Used this to apply simplified constant-schedule loads which will mimic the PHPP standard values. 
-
Use this together with Honeybee 'setEPZoneLoads' and 'setEPZoneSchedules' components in order to set the EnergyPlus values to match the PHPP. Note that you might not want to do this since the PHPP is quitte differenc than typical E+ loads. But if you want the two models to match exactly, use these loads here.
-
EM Sept. 29, 2020
    Args:
        _HBZones: The Honeybee Zones
        num_res_units_: (int) Default=1. 
    Returns:
        HBZones_: The Honeybee Zones
        PHPP_equipmentLoadPerArea_: (W/m2) Connect this to the 'numOfPeoplePerArea_' input on a Honyebee 'equipmentLoadPerArea_' Component if you want to use this value for the E+ model in addition to the PHPP.
        PHPP_lightingDensityPerArea_: (W/m2) Connect this to the 'numOfPeoplePerArea_' input on a Honyebee 'lightingDensityPerArea_' Component if you want to use this value for the E+ model in addition to the PHPP.
        PHPP_numOfPeoplePerArea_: (ppl/m2) Occpancy Load. Connect this to the 'numOfPeoplePerArea_' input on a Honyebee 'setEPZoneLoads' Component if you want to use this value for the E+ model in addition to the PHPP.
        PHPP_occupancySchedule_: Constant Val Schedule (1). Connect this to the 'occupancySchedules_' input on a Honyebee 'setEPZoneSchedules' Component if you want to use this value for the E+ model in addition to the PHPP.
        PHPP_lightingSchedule_: Constant Val Schedule (1). Connect this to the 'lightingSchedules_' input on a Honyebee 'setEPZoneSchedules' Component if you want to use this value for the E+ model in addition to the PHPP.
        PHPP_epuipmentSchedule_: Constant Val Schedule (1). Connect this to the 'equipmentSchedules_' input on a Honyebee 'setEPZoneSchedules' Component if you want to use this value for the E+ model in addition to the PHPP.
"""

ghenv.Component.Name = "BT_SetPHPPResLoads"
ghenv.Component.NickName = "Set PHPP Res Loads"
ghenv.Component.Message = 'SEP_29_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc
import math
import ghpythonlib.components as ghc

hb_hive = sc.sticky["honeybee_Hive"]()
HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)

#-------------------------------------------------------------------------------
# Defs From LB Constant Schedule
def IFDstrFromDayVals(dayValues, schName, daytype, schTypeLims):
    idfStr = 'Schedule:Day:Interval,\n' + \
        '\t' + schName + ' Day Schedule - ' + daytype + ', !- Name\n' + \
        '\t' + schTypeLims + ', !- Schedule Type Limits Name\n' + \
        '\t' + 'No,        !- Interpolate to Timestep\n'
    
    tCount = 1
    for hCount, val in enumerate(dayValues):
        if hCount+1 == len(dayValues):
            idfStr = idfStr + '\t24:00,   !- Time ' + str(tCount) + ' {hh:mm}\n' + \
                '\t' + str(val) + ';     !- Value Until Time ' + str(tCount) + '\n'
        elif val == dayValues[hCount+1]: pass
        else:
            idfStr = idfStr + '\t' + str(hCount+1) + ':00,   !- Time ' + str(tCount) + ' {hh:mm}\n' + \
                '\t' + str(val) + ',     !- Value Until Time ' + str(tCount) + '\n'
            tCount += 1
    
    return idfStr

def IFDstrForWeek(daySchedName, schName):
    idfStr = 'Schedule:Week:Daily,\n' + \
        '\t' + schName + ' Week Schedule' + ', !- Name\n' + \
        '\t' + daySchedName + ',  !- Sunday Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- Monday Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- Tuesday Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- Wednesday Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- Thursday Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- Friday Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- Saturday Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- Holiday Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- SummerDesignDay Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- WinterDesignDay Schedule:Day Name\n' + \
        '\t' + daySchedName + ',  !- CustomDay1 Schedule:Day Name\n' + \
        '\t' + daySchedName + ';  !- CustomDay2 Schedule:Day Name\n'
    
    return idfStr

def IFDstrForYear(weekSchedName, schName, schTypeLims):
    idfStr = 'Schedule:Year,\n' + \
        '\t' + schName + ', !- Name\n' + \
        '\t' + schTypeLims + ', !- Schedule Type Limits Name\n' + \
        '\t' + weekSchedName + ',  !- Schedule:Week Name\n' + \
        '\t' + '1' + ',  !- Start Month 1\n' + \
        '\t' + '1' + ',  !- Start Day 1\n' + \
        '\t' + '12' + ',  !- End Month\n' + \
        '\t' + '31' + ';  !- End Day\n'
    
    return idfStr

def main(values, schedName, schedTypeLimits):
    # Import the classes.
    lb_preparation = sc.sticky["ladybug_Preparation"]()
    hb_EPObjectsAux = sc.sticky["honeybee_EPObjectsAUX"]()
    scheduleTypeLimitsLib = sc.sticky["honeybee_ScheduleTypeLimitsLib"].keys()
    
    # Generate a schedule IDF string if writeSchedule_ is set to True.
    daySchedCollection = []
    daySchNameCollect = []
    schedIDFStrs = []
    daySchedNames = []
    
    # Get the type limits for the schedule.
    if schedTypeLimits == None:
        schTypeLims = 'Fractional'
    else:
        schTypeLims = schedTypeLimits
        if schTypeLims.upper() == 'TEMPERATURE':
            schTypeLims = 'TEMPERATURE 1'
        if schTypeLims.upper() not in scheduleTypeLimitsLib:
            warning = "Can't find the connected _schedTypeLimits_ '" + schTypeLims + "' in the Honeybee EP Schedule Library."
            ghenv.Component.AddRuntimeMessage(gh.GH_RuntimeMessageLevel.Warning, warning)
            return -1
    
    # Write out text strings for the daily schedules
    if len(values) == 1:
        values = [values[0] for x in range(24)]
    elif len(values) == 24:
        pass
    else:
        warning = "_value must be either a single value or a list of 24 values for each hour of the day."
        print warning
        ghenv.Component.AddRuntimeMessage(w, warning)
    
    schedIDFStrs.append(IFDstrFromDayVals(values, schedName, 'Constant', schTypeLims))
    daySchName = schedName + ' Day Schedule - ' + 'Constant'
    
    # Write out text strings for the weekly values.
    schedIDFStrs.append(IFDstrForWeek(daySchName, schedName))
    weekSchedName = schedName + ' Week Schedule'
    
    # Write out text for the annual values.
    schedIDFStrs.append(IFDstrForYear(schedName + ' Week Schedule', schedName, schTypeLims))
    yearSchedName = schedName
    
    # Write all of the schedules to the memory of the GH document.
    for EPObject in schedIDFStrs:
        added, name = hb_EPObjectsAux.addEPObjectToLib(EPObject, overwrite = True)
    
    return yearSchedName, weekSchedName, schedIDFStrs


#-------------------------------------------------------------------------------
def calcOccupancy(_HBZoneObjects, _numDwellingUnits):
    def get_zone_total_TFA(zone):
        return sum([room.FloorArea_TFA for room in zone.PHPProoms])
    
    def get_zone_total_floor_area(zone):
        def is_a_floor(_srfc):
            return _srfc.type >= 2 and _srfc.type <3
        
        def get_surface_area(_srfc):
            return ghc.Area(_srfc.geometry)[0]
        
        return sum([get_surface_area(surface) for surface in zone.surfaces if is_a_floor(surface)])
    
    # 1) Get the total Building TFA
    # 2) Calc the total buiding occupancy based on the total TFA
    #    Formula used here is from PHPP 'Verification' worksheet
    # 3) Figure out the Building's Total EP Floor Area
    # 4) Calc the building's Occupancy Per floor-area
    #
    # Returns PHPP Occupancy normalized by E+ Floor Area (gross)
    
    bldgTFA = sum([get_zone_total_TFA(zone) for zone in _HBZoneObjects])
    bldgOccupancy = (1+1.9*(1 - math.exp(-0.00013 * math.pow((bldgTFA/_numDwellingUnits-7), 2))) + 0.001 * bldgTFA/_numDwellingUnits ) * _numDwellingUnits
    bldgFloorArea = sum([get_zone_total_floor_area(zone) for zone in _HBZoneObjects])
    
    return bldgFloorArea, bldgOccupancy, bldgOccupancy / bldgFloorArea

def calc_lighting(_lightingEff, _bldgOcc, _bldgFloorArea):
    lighting_FrequencyPerPerson = 2.9 #kh/(Person-year)
    lighting_Frequency = lighting_FrequencyPerPerson * _bldgOcc
    lighting_Demand = 720 / _lightingEff
    lighting_UtilizFactor = 1
    lighting_AnnualEnergy = lighting_Demand * lighting_UtilizFactor * lighting_Frequency #kWh / year
    lighting_AvgHourlyWattage =  lighting_AnnualEnergy * 1000  / 8760 # Annual Average W / hour for a constant schedule
    lightingDensityPerArea = lighting_AvgHourlyWattage / _bldgFloorArea # For constant operation schedule
    
    return lightingDensityPerArea

def calc_elec_equip_appliances(_zones, _bldgOcc, _numUnits, _bldgFA):
    appliances = [appliance for zone in _zones for appliance in zone.PHPP_ElecEquip]
    appliance_annual_kWh = [appliance.calcAnnualDemand(_bldgOcc, _numUnits) for appliance in appliances]
    appliance_avg_hourly_W = sum(appliance_annual_kWh) * 1000 / 8760
    
    return appliance_avg_hourly_W / _bldgFA

#-------------------------------------------------------------------------------
if len(_HBZones) is not 0:
    # Occupancy
    numDwellingUnits = 1 if num_res_units_==None else num_res_units_
    bldgGrossFloorArea, bldgOcc, PHPP_numOfPeoplePerArea_ = calcOccupancy(HBZoneObjects, numDwellingUnits)
    
    # Lighting
    PHPP_lightingDensityPerArea_ = calc_lighting(50, bldgOcc, bldgGrossFloorArea)
    
    # Appliances
    PHPP_equipmentLoadPerArea_ = calc_elec_equip_appliances(HBZoneObjects, bldgOcc, numDwellingUnits, bldgGrossFloorArea)


#-------------------------------------------------------------------------------
# Create constant value Schedules
if len(_HBZones) is not 0:
    result = main([1], 'PHPP_Const_Occupancy', 'Fractional')
    if result != -1:
        PHPP_occupancySchedule_, weekSched, schedIDFText = result
        print '\nscheduleValues generated!'
    
    result = main([1], 'PHPP_Const_Lighting', 'Fractional')
    if result != -1:
        PHPP_lightingSchedule_, weekSched, schedIDFText = result
        print '\nscheduleValues generated!'
    
    result = main([1], 'PHPP_Const_Elec_Equip', 'Fractional')
    if result != -1:
        PHPP_epuipmentSchedule_, weekSched, schedIDFText = result
        print '\nscheduleValues generated!'

#-------------------------------------------------------------------------------
if len(_HBZones) is not 0:
    HBZones_ = _HBZones