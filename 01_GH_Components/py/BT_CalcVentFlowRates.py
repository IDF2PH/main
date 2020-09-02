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
Will determine the fresh-air ventilation flow rates for all the PHPP Rooms in the HB Zones. Reads from the HB/EP Schedules and Loads applied to the HB zones
>  If you like, you can pass in a PHPP-Style ventilation schedule into the '_phppVenSched'. If not, this will read the HB Schedule applied to the Zone and use that to create a PHPP-Style schedule with a CONSTANT flow rate.
>  If you want to use this component to align a PHPP and EP model, use an HB 'Constant Schedule' object and set the zone's ventilation schedule to '1'.
-
EM September 1, 2020

    Args:
        _HBZones: List. A list of all the HB Zones to use.
        zoneGrossFloorArea_: <Optional>: (m2) A value representing the total zone floor area (gross) including all interstitial floor levels.
        airFlowInput_: ('EP' | 'UD') <Optional> The Type of values to use for the room's/zone's Air-Flow (m3/h). If 'EP' is input (or none is input) this calculator will figure out the EP Fresh-Air loads based on the HB/EP Program and apply those values to the PHPP Rooms. IF 'UD' is input (User Determined) this calculator will try and use any Rhino-Scene / GH-Scene inputs from the 'Create PHPP Room' component. Any rooms which didn't have values input will default to the HB/EP Zone's Program loads
        _phppVentSched: <Optional> Input a simplified PHPP-Style ventilation schedule from the 'PHPP Vent Sched' component. This will override any HB Schedules found in the zone and will be applied to all rooms in the zone.
    Returns:
        HBZones_: The HB Zones with the new Ventilation Airflow information added
        ventilationPerArea_: (m3/s-m2) the TOTAL average yearly ventilation flow rate for the zone at a CONSTANT rate. Input this value into an HB 'setEPZoneLoads' component and then input a 'Constant Schedule' into an HB 'setEPZoneSchedules' component.
        ventilationPerPerson_: (m3/s-person) an override set to 0. All the ventilation is encompased in the above 'PerArea' value.
"""

ghenv.Component.Name = "BT_CalcVentFlowRates"
ghenv.Component.NickName = "Room Vent Flowrates"
ghenv.Component.Message = 'SEP_01_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"


import scriptcontext as sc
import Grasshopper.Kernel as gh
import os
from collections import deque
import itertools
from datetime import date
from System import Object
from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path
import ghpythonlib.components as ghc
from collections import namedtuple

# Defs and Classes
PHPP_Sys_Ventilation = sc.sticky['PHPP_Sys_Ventilation']
hb_hive = sc.sticky["honeybee_Hive"]()
HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)

# ------------------------------------------------------------------------------
########            FROM HB 'convertEPSCHValues' COMPONENT              ########
def getNationalHols(country, weekStartWith):
    # Dictionary of national hoildays arranged by country.
    # All countries are accounted for with the exception of:
    # BLZ (CENTRAL AMERICA), BRN (SOUTH PACIFIC), GUM (SOUTH PACIFIC), MHL (SOUTH PACIFIC), PLW (SOUTH PACIFIC), UMI (SOUTH PACIFIC)
    # https://energyplus.net/weather
    ###################################################
    # Source: http://www.officeholidays.com/countries/ ------ http://www.officeholidays.com/countries/
    listOfValues = [0, 1, 2, 3, 4, 5, 6]
    week = deque(listOfValues)
    week.rotate(-weekStartWith)
    dayWeekList = list(itertools.chain.from_iterable(itertools.repeat(week, 53)))[:365]
    countries = {
    'USA':[0,[i for i, x in enumerate(dayWeekList) if x == 2][2],[i for i, x in enumerate(dayWeekList) if x == 2][21],184,[i for i, x in enumerate(dayWeekList) if x == 2][35],314,[i for i, x in enumerate(dayWeekList) if x == 5][46],359],
    'CAN': [0,181,[i for i, x in enumerate(dayWeekList) if x == 2][35],358,359],
    'CUB': [0,1,[i for i, x in enumerate(dayWeekList) if x == 6][12],120,205,206,207,282,358,364],
    'GTM': [0,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 0][12],120,180,257,292,304,358],
    'HND': [0,[i for i, x in enumerate(dayWeekList) if x == 4][11],[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 0][12],120,257,278,358],
    'MEX': [0,[i for i, x in enumerate(dayWeekList) if x == 2][4],[i for i, x in enumerate(dayWeekList) if x == 2][11],[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],120,258,[i for i, x in enumerate(dayWeekList) if x == 2][46],345,358,359],
    'MTQ': [0,[i for i, x in enumerate(dayWeekList) if x == 2][12],120,124,127,[i for i, x in enumerate(dayWeekList) if x == 2][19],194,226,304,314,358],
    'NIC': [0,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],120,121,169,256,257,341,358],
    'PRI': [0,5,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 2][12],127,169,327,358],
    'SLV': [0,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 0][12],120,129,168,215,216,217,305,358,359],
    'VIR': [0,5,17,45,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],89,149,153,248,304,314,358,359],
    'ARG': [0,[i for i, x in enumerate(dayWeekList) if x == 2][5],[i for i, x in enumerate(dayWeekList) if x == 3][5],82,[i for i, x in enumerate(dayWeekList) if x == 6][12],91,120,144,170,188,189,[i for i, x in enumerate(dayWeekList) if x == 2][32],282,[i for i, x in enumerate(dayWeekList) if x == 2][47],341,358],
    'BOL': [0,21,[i for i, x in enumerate(dayWeekList) if x == 2][5],[i for i, x in enumerate(dayWeekList) if x == 3][5],[i for i, x in enumerate(dayWeekList) if x == 5][11],120,[i for i, x in enumerate(dayWeekList) if x == 5][20],171,305,358],
    'BRA': [0,[i for i, x in enumerate(dayWeekList) if x == 2][5],[i for i, x in enumerate(dayWeekList) if x == 3][5],[i for i, x in enumerate(dayWeekList) if x == 4][5],[i for i, x in enumerate(dayWeekList) if x == 5][11],110,120,[i for i, x in enumerate(dayWeekList) if x == 5][20],249,284,305,318,323,358],
    'CHL': [0,[i for i, x in enumerate(dayWeekList) if x == 5][11],120,140,179,196,226,261,262,303,304,341,358],
    'COL': [0,10,79,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],120,128,149,156,[i for i, x in enumerate(dayWeekList) if x == 2][26],200,218,226,289,[i for i, x in enumerate(dayWeekList) if x == 2][44],317,341,358],
    'ECU': [0,38,39,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 1][12],120,143,221,281,306,358],
    'PER': [0,[i for i, x in enumerate(dayWeekList) if x == 4][11],[i for i, x in enumerate(dayWeekList) if x == 5][11],120,179,208,209,242,280,304,341,358,359],
    'PRY': [0,59,[i for i, x in enumerate(dayWeekList) if x == 4][11],[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 1][12],120,133,134,162,227,271,284,341,358,364],
    'URY': [0,38,39,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],108,120,169,198,284,236,305,358],
    'VEN': [0,38,39,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],108,120,174,179,185,204,226,284,304,358,359,364],
    'AUS': [0,25,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],114,358,359,360],
    'FJI': [0,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 0][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],174,249,282,303,345,359,360],
    'MYS': [38,120,121,140,155,186,187,242,254,258,274,345,358,359],
    'NZL': [0,3,38,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 0][12],114,[i for i, x in enumerate(dayWeekList) if x == 2][22],[i for i, x in enumerate(dayWeekList) if x == 2][42],358,359,360],
    'PHL': [0,1,38,55,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 0][12],98,120,162,187,232,240,253,303,304,333,357,358,363,364],
    'SGP': [0,38,39,[i for i, x in enumerate(dayWeekList) if x == 6][12],120,121,140,218,220,254,302,358,359],
    'DZA': [0,120,187,253,274,283,304,345],
    'EGY': [6,24,114,120,121,187,188,189,204,253,254,255,274,278,345],
    'ETH': [6,19,60,119,120,124,147,253,255,269,345],
    'GHA': [0,64,65,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,121,144,181,186,253,263,[i for i, x in enumerate(dayWeekList) if x == 6][48],358,359],
    'KEN': [0,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,187,253,292,345,358,359],
    'LBY': [47,120,187,188,252,253,254,255,258,274,295,345,357],
    'MAR': [0,10,120,188,210,225,231,232,253,274,309,321,345],
    'MDG': [0,[i for i, x in enumerate(dayWeekList) if x == 2][12],88,120,124,135,176,226,304,345,358],
    'SEN': [0,[i for i, x in enumerate(dayWeekList) if x == 2][12],93,120,124,[i for i, x in enumerate(dayWeekList) if x == 2][19],187,226,255,304,345,358],
    'TUN': [0,13,78,98,120,187,188,189,205,224,253,254,255,274,287,345],
    'ZAF': [0,79,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],117,120,121,166,220,266,349,358,359],
    'ZWE': [0,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],107,120,144,[i for i, x in enumerate(dayWeekList) if x == 2][31],[i for i, x in enumerate(dayWeekList) if x == 3][31],355,358,359],
    'ARE': [0,124,186,187,252,253,254,255,274,333,335,336,344],
    'BDG': [51,75,84,103,120,140,142,181,183,185,187,226,236,253,254,255,283,284,345,349,358],
    'CHN': [0,37,38,39,40,41,42,43,93,120,121,159,257,258,273,274,275,276,277,278,279],
    'IND': [25,226,274],
    'IRN': [41,71,77,78,79,80,81,89,90,110,124,141,153,154,170,177,187,211,255,263,264,283,284,324,325,350,354],
    'JPN': [0,[i for i, x in enumerate(dayWeekList) if x == 2][1],41,79,118,122,123,124,[i for i, x in enumerate(dayWeekList) if x == 2][28],222,[i for i, x in enumerate(dayWeekList) if x == 2][37],264,282,306,326,356],
    'KAZ': [0,6,59,66,79,80,81,91,98,187,241,253,334,349,352],
    'KOR': [0,37,38,39,40,59,124,133,156,226,256,257,258,275,281,358],
    'KWT': [2,55,56,124,156,157,158,252,253,254,255,274,345],
    'LKA': [14,22,34,52,65,80,[i for i, x in enumerate(dayWeekList) if x == 6][12],102,103,110,120,140,141,169,186,199,228,254,284,301,317,345,346,358],
    'MAC': [0,38,39,40,94,120,258,273,293,353],
    'MDV': [0,11,120,156,186,206,253,254,274,306,314,334,345,364],
    'MNG': [0,38,191,192,193,194,195,362],
    'NPL': [14,49],
    'PAK': [35,81,120,187,188,189,225,253,254,283,312,345,358],
    'PRK': [0,39,46,52,104,114,120,156,169,207,236,251,281,357,360],
    'SAU': [187,190,191,253,254,255,264,265],
    'THA': [0,52,95,102,103,104,121,124,126,127,139,168,169,223,296,338,345,364],
    'TWN': [0,37,38,39,40,41,42,93,120,128,257,282],
    'UZB': [0,13,66,79,98,187,243,255,273,341],
    'VNM': [0,36,37,38,39,40,105,106,119,120,121,122,244],
    'AUT': [0,5,[i for i, x in enumerate(dayWeekList) if x == 2][12],120,124,[i for i, x in enumerate(dayWeekList) if x == 2][19],145,226,298,304,311,341,358,359],
    'BEL': [0,[i for i, x in enumerate(dayWeekList) if x == 2][12],120,124,[i for i, x in enumerate(dayWeekList) if x == 2][19],201,226,304,314,358],
    'BRG': [0,61,62,[i for i, x in enumerate(dayWeekList) if x == 6][17],120,[i for i, x in enumerate(dayWeekList) if x == 2][17],125,141,247,248,264,265,357,358,359],
    'BIH': [0,1,120,121],
    'BLR': [0,6,7,65,66,120,128,129,183,310,358],
    'CHE': [0,[i for i, x in enumerate(dayWeekList) if x == 6][12],124,212,358],
    'CYP': [0,5,[i for i, x in enumerate(dayWeekList) if x == 2][10],83,90,[i for i, x in enumerate(dayWeekList) if x == 6][17],120,[i for i, x in enumerate(dayWeekList) if x == 2][17],[i for i, x in enumerate(dayWeekList) if x == 2][24],226,273,300,358,359],
    'CZE': [0,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,127,185,186,270,300,320,357,358,359],
    'DEU': [0,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,124,[i for i, x in enumerate(dayWeekList) if x == 2][19],275,358,359],
    'DNK': [0,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],[i for i, x in enumerate(dayWeekList) if x == 6][16],124,[i for i, x in enumerate(dayWeekList) if x == 2][19],155,357,358,359],
    'ESP': [0,5,[i for i, x in enumerate(dayWeekList) if x == 6][12],120,226,284,304,339,341,358],
    'FIN': [0,5,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,124,174,175,308,340,357,358,359],
    'FRA': [0,[i for i, x in enumerate(dayWeekList) if x == 2][12],120,124,127,[i for i, x in enumerate(dayWeekList) if x == 2][19],194,226,304,314,358],
    'GBR': [0,[i for i, x in enumerate(dayWeekList) if x == 6][12],121,149,359,360],
    'GRC': [0,5,[i for i, x in enumerate(dayWeekList) if x == 2][10],83,[i for i, x in enumerate(dayWeekList) if x == 6][17],120,[i for i, x in enumerate(dayWeekList) if x == 2][17],[i for i, x in enumerate(dayWeekList) if x == 2][24],226,300,358],
    'HUN': [0,73,[i for i, x in enumerate(dayWeekList) if x == 1][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,[i for i, x in enumerate(dayWeekList) if x == 1][19],[i for i, x in enumerate(dayWeekList) if x == 2][19],231,295,304,358,359],
    'IRL': [0,75,[i for i, x in enumerate(dayWeekList) if x == 2][12],[i for i, x in enumerate(dayWeekList) if x == 2][17],[i for i, x in enumerate(dayWeekList) if x == 2][22],[i for i, x in enumerate(dayWeekList) if x == 2][30],303,358,359,360],
    'ISL': [0,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 1][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],110,120,124,[i for i, x in enumerate(dayWeekList) if x == 1][19],[i for i, x in enumerate(dayWeekList) if x == 2][19],167,212,357,358,359,364],
    'ISR': [82,113,119,131,163,225,275,276,284,289,297],
    'ITA': [0,5,[i for i, x in enumerate(dayWeekList) if x == 2][12],114,120,152,226,304,341,358,359],
    'LTU': [0,46,69,[i for i, x in enumerate(dayWeekList) if x == 2][12],120,[i for i, x in enumerate(dayWeekList) if x == 1][22],174,186,226,304,357,358,359],
    'NLD': [0,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],116,124,[i for i, x in enumerate(dayWeekList) if x == 2][19],358],
    'NOR': [0,[i for i, x in enumerate(dayWeekList) if x == 5][11],[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,124,[i for i, x in enumerate(dayWeekList) if x == 2][19],136,358,359],
    'POL': [0,5,[i for i, x in enumerate(dayWeekList) if x == 1][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,122,134,145,226,304,314,358,359],
    'PRT': [0,[i for i, x in enumerate(dayWeekList) if x == 6][12],114,120,160,226,341,357,358],
    'ROU': [0,1,23,120,[i for i, x in enumerate(dayWeekList) if x == 2][17],170,226,333,334,358,359],
    'RUS': [0,3,4,5,6,53,66,120,128,163,307],
    'SRB': [0,1,6,7,45,46,[i for i, x in enumerate(dayWeekList) if x == 6][17],[i for i, x in enumerate(dayWeekList) if x == 0][17],[i for i, x in enumerate(dayWeekList) if x == 1][17],[i for i, x in enumerate(dayWeekList) if x == 2][17],128,283],
    'SVK': [0,5,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,127,155,240,243,257,273,289,357,358,359],
    'SVN': [0,38,[i for i, x in enumerate(dayWeekList) if x == 2][12],116,120,121,175,226,303,304,358,359],
    'SWE': [0,5,[i for i, x in enumerate(dayWeekList) if x == 6][12],[i for i, x in enumerate(dayWeekList) if x == 2][12],120,124,156,174,175,277,357,358,359,364],
    'SYR': [0,66,106,120,125,187,255,275,278,345,358],
    'TUR': [0,112,120,138,157,158,159,241,253,254,255,256,301],
    'UKR': [0,6,66,120,[i for i, x in enumerate(dayWeekList) if x == 2][17],127,128,[i for i, x in enumerate(dayWeekList) if x == 2][24],235,286,324]
    }
    
    return countries[country]

def addHolidays(startVals, holidayValues, weekStartWith, epwFile, customHolidays, lb_preparation):
    holidayDOYs = []
    if epwFile:
        #get the base code from EPW
        locationData = lb_preparation.epwLocation(epwFile)
        codeNation = lb_preparation.epwDataReader(epwFile, locationData[0])[14][1]
        code = codeNation.split("_")
        if len(code) == 3:
            country = code[2]
        elif len(code) == 2:
            country = code[1]
        else:
            country = code[1]
        
        try:
            holidayDOYs = getNationalHols(country, weekStartWith)
        except:
            holidayDOYs = [0,120,358,359] # international holidays for countries not found.
        
        # Give a message to show the national holidays.
        print "National holidays(DOYs): "
        for item in holidayDOYs:
            print item+1
        print "_"
    
    # Add customHolidays.
    monthsDict = {'01':'JAN', '02':'FEB', '03':'MAR', '04':'APR', '05':'MAY', '06':'JUN',
    '07':'JUL', '08':'AUG', '09':'SEP', '10':'OCT', '11':'NOV', '12':'DEC'}
    def fromDateToDay(holiday, months):
        startDate = date(2015, 1, 1)
        textMonth = holiday.split(' ')[0].upper()
        dictReverse = {value: key for key, value in months.items()}
        intMonth = int(dictReverse[textMonth])
        endDate = date(2015, intMonth, int(holiday.split(' ')[1]))
        period = endDate - startDate
        day = period.days + 1
        return day
    
    if customHolidays != []:
        try:
            if customHolidays is not type(int):
                customHolidays = map(int, customHolidays)
        except ValueError:
            customHolidaysDOY = []
            for item in customHolidays:
                dayDate = fromDateToDay(item, monthsDict)
                customHolidaysDOY.append(dayDate)
            customHolidays = customHolidaysDOY
        
        # Give a message to show the national holidays.
        print "Custom holidays(DOYs): "
        for item in customHolidays:
            print item
            if item-1 not in holidayDOYs:
                holidayDOYs.append(item-1)
    
    # Edit the list of schedule values.
    for day in holidayDOYs:
        for interval in holidayValues:
            if day >= interval[0]-1 and day <= interval[1]-1:
                startVals[day] = interval[2]
    
    # Build up a list of holidays
    def fromDayToDate(day, months):
        dateDay = date.fromordinal(date(2015, 1, 1).toordinal() + day) # 2015 is not leap year
        month, day = str(dateDay).split('-')[1:]
        monthsDate = months[month]
        dateFromDay = monthsDate + ' ' + day
        return dateFromDay
    holidayDates = []
    for day in holidayDOYs:
        holidayDates.append(fromDayToDate(day, monthsDict))
    
    return startVals, holidayDates

def main(schName, startDayOfTheWeek, epwFile, customHol):
    # Check the start day of the week.
    daysOfWeek = {'sun':0, 'mon':1, 'tue':2, 'wed':3, 'thu':4, 'fri':5, 'sat':6,
    'sunday':0, 'monday':1, 'tuesday':2, 'wednesday':3, 'thursday':4, 'friday':5, 'saturday':6}
    
    if startDayOfTheWeek != None:
        try:
            startDayOfTheWeek = int(startDayOfTheWeek)%7
        except:
            try:
                startDayOfTheWeek = daysOfWeek[startDayOfTheWeek.lower()]
            except:
                warning = 'Input for _weekStartDay_ is not valid.'
                print warning
                ghenv.Component.AddRuntimeMessage(gh.GH_RuntimeMessageLevel.Warning, warning)
                return -1
    else:
        startDayOfTheWeek = 0
    
    lb_preparation = sc.sticky["ladybug_Preparation"]()
    readSchedules = sc.sticky["honeybee_ReadSchedules"](schName, startDayOfTheWeek)
    
    values = []
    holidays = []
    dataGotten = False
    if schName.lower().endswith(".csv"):
        # check if csv file exists
        if not os.path.isfile(schName):
            msg = "Cannot find the shchedule file: " + schName
            print msg
            ghenv.Component.AddRuntimeMessage(gh.GH_RuntimeMessageLevel.Warning, msg)
            return -1
        else:
            result = open(schName, 'r')
            lineCount = 0
            for lineCount, line in enumerate(result):
                readSchedules.schType = 'schedule:year'
                readSchedules.startHOY = 1
                readSchedules.endHOY = 8760
                if 'Daysim' in line: pass
                else:
                    if lineCount == 0:readSchedules.unit = line.split(',')[-2].split(' ')[-1].upper()
                    elif lineCount == 1: readSchedules.schName = line.split('; ')[-1].split(':')[0]
                    elif lineCount < 4: pass
                    else:
                        columns = line.split(',')
                        try:
                            values.append(float(columns[4]))
                        except:
                            values.append(float(columns[3]))
                    lineCount += 1
            values.insert(0, values.pop(-1))
            dataGotten = True
    else:
        HBScheduleList = sc.sticky["honeybee_ScheduleLib"].keys()
        if schName.upper() not in HBScheduleList:
            msg = "Cannot find " + schName + " in Honeybee schedule library."
            print msg
            ghenv.Component.AddRuntimeMessage(gh.GH_RuntimeMessageLevel.Warning, msg)
            return -1
        else:
            dataGotten = True
            values = readSchedules.getScheduleValues()
    
    if dataGotten == True:
        # Check for any holidays.
        if epwFile or customHol != []:
            if len(values) == 365 and not schName.lower().endswith(".csv"):
                holidayValues = readSchedules.getHolidaySchedValues(schName)
                values, holidays = addHolidays(values, holidayValues, startDayOfTheWeek, epwFile, customHol, lb_preparation)
        
        strToBeFound = 'key:location/dataType/units/frequency/startsAt/endsAt'
        d, m, t = lb_preparation.hour2Date(readSchedules.startHOY, True)
        startDate = m+1, d, t
        
        d, m, t = lb_preparation.hour2Date(readSchedules.endHOY, True)
        endDate = m+1, d, t
        if readSchedules.endHOY%24 == 0:
            endDate = m+1, d, 24
        
        header = [strToBeFound, readSchedules.schType, readSchedules.schName, \
                  readSchedules.unit, 'Hourly', startDate, endDate]
        
        try: values = lb_preparation.flattenList(values)
        except: pass
        
        return header + values, holidays
    else:
        return -1

# ------------------------------------------------------------------------------

def checkInputs(_HBZoneObjects, _type):
    ## Check the inputs, give warnings
    
    totalModelRooms = []
    for zone in _HBZoneObjects:
        try:
            # Check for demand controls
            if zone.HVACSystem.airDetails.fanControl == 'Variable Volume':
                pass
            elif zone.HVACSystem.airDetails.fanControl == 'Constant Volume':
                msgDemandControl = "It looks like the Zone(s) have no 'Demand Control' for Fresh Air Ventilation?\nFor the EP simulations results to match this calculator, be sure to apply 'Demand Control' to your EP model.\nAdd a 'Honeybee_HVAC Air Details' Component and set 'demandControlledVent_' option to TRUE in order to do this."
                ghenv.Component.AddRuntimeMessage(gh.GH_RuntimeMessageLevel.Warning, msgDemandControl)
        except:
            pass
        
        try:
            # Check for length of rooms, warn about PHPP Exccel table length
            noOfRooms = len(zone.PHPProoms)
            totalModelRooms.append(noOfRooms)
        except:
            pass
    
    if _type != None:
        return _type
    else:
        return 'EP'

def histogram(_data, _nbins):
    # Creates a dictionary Histogram of some data in n-bins
    
    #print _data
    min_val = min(_data)
    max_val = max(_data)
    hist_bins = {} # The number of items in each bin
    hist_vals = {} # The avg value for each bin
    total = 0
    
    # Initialize the dict
    for k in range(_nbins+1):
        hist_bins[k] = 0
        hist_vals[k] = 0
    
    # Create the Histogram
    for d in _data:
        bin_number = int(_nbins * ((d - min_val) / (max_val - min_val)))
        
        hist_bins[bin_number] += 1
        hist_vals[bin_number] += d
        total += 1
    
    # Clean up / fix the data for output
    for n in hist_vals.keys():
        hist_vals[n] =  hist_vals[n] / hist_bins[n]
    
    for h in hist_bins.keys():
        hist_bins[h] = hist_bins[h] / total
    
    return hist_bins, hist_vals # The number of items in each bin, the avg value of the items in the bin

def getHBzoneFloorArea(_HBzoneObj, _zoneGrossFloorArea):
    # Finds and returns the zone total floor area
    # This will use ALL the floor-type surfaces to determing the zone floor area. Is that right? hmm....
    
    if _zoneGrossFloorArea == None:
        floorAreas = []
        for surface in _HBzoneObj.surfaces:
            if surface.type >= 2 and surface.type <3: # 2-3 are the Floor surface types
                floorAreas.append(ghc.Area(surface.geometry)[0] ) # Return 0 is the Area)
        zoneFloorArea = sum(floorAreas)
    else:
        zoneFloorArea = float(_zoneGrossFloorArea)
    
    return zoneFloorArea

def getHBLoadAndSched(_HBzoneObj):
    ############################################################################
    # Get the HB/EP Ventilation Loads and Sched for the Zone from the Hive
    ############################################################################
    
    # Loads from HB Library
    HBZoneLoads = _HBzoneObj.getCurrentLoads(True, ghenv.Component)
    
    HBnumOfPeoplePerArea = _HBzoneObj.numOfPeoplePerArea
    HBventilationPerArea = _HBzoneObj.ventilationPerArea
    HBventilationPerPerson = _HBzoneObj.ventilationPerPerson
    
    # Get the HB/EP Occupancy Schedule for the zone for the Hive
    HBZoneSchedules = _HBzoneObj.getCurrentSchedules(True, ghenv.Component)
    occupancySchedule = HBZoneSchedules['occupancySchedule']
    
    if occupancySchedule:
        result = main(occupancySchedule, None, None, [])[0]
    
    # Clean up the HB/EP Occupancy Schdeule (remove text header)
    HBoccupancySchedualAsValues_ = []
    for eachItem in result:
        try:
            HBoccupancySchedualAsValues_.append(float(eachItem))
        except:
            pass
    
    return HBnumOfPeoplePerArea, HBventilationPerArea, HBventilationPerPerson, HBoccupancySchedualAsValues_

def calcZoneAnnualVentFlowRateFromHB(_HBzoneObj, _userVentSched, _zoneGrossFloorArea):
    # Figure out the HB Zone's Floor Area to use (different than TFA)
    zoneFloorArea = getHBzoneFloorArea(_HBzoneObj, _zoneGrossFloorArea)
    
    if zoneFloorArea == 0:
        warning = "Something wrong with the floor area - are you sure there is at least one Floor surface in the zone?"
        ghenv.Component.AddRuntimeMessage(gh.GH_RuntimeMessageLevel.Warning, warning)
    
    # Pull the Ventilaiton loads, Occupancy Schedule from the Hive
    numOfPeoplePerArea, ventilationPerArea, ventilationPerPerson, zoneOccSchedAsValues = getHBLoadAndSched(_HBzoneObj)
    
    # Calc the Avg Zone Occupancy
    avgZoneOccupancy = (sum(zoneOccSchedAsValues) / len(zoneOccSchedAsValues)) * numOfPeoplePerArea * zoneFloorArea
    
    # Calc the Hourly Flow rates (m3/h) from the HB Hive Schedule
    zoneVentilation_forArea = [ventilationPerArea * zoneFloorArea * 60 * 60] * 8760 # m3/s---> m3/h
    zoneVentilation_forPeople = []
    for eachHourOccRate in zoneOccSchedAsValues:
        #zoneVentilation_forPeople.append(eachHourOccRate * numOfPeoplePerArea * zoneFloorArea * ventilationPerPerson * 60 * 60) # m3/s---> m3/h
        zoneVentilation_forPeople.append( numOfPeoplePerArea * zoneFloorArea * ventilationPerPerson * 60 * 60) # m3/s---> m3/h
    
    zoneVentilation_forArea_Avg = sum(zoneVentilation_forArea) / len(zoneVentilation_forArea)
    zoneVentilation_forPeople_Avg = sum(zoneVentilation_forPeople) / len(zoneVentilation_forPeople)
    zoneVentilation_Total_AnnualAvg = zoneVentilation_forArea_Avg + zoneVentilation_forPeople_Avg
    
    # Return the average Annual Ventialtion Flow rate (m3/h) based on the Zone's HB Schedules
    
    print "The HB Zone '{}' has an average Annual airflow of: {:.2f} m3/h".format(_HBzoneObj.name, zoneVentilation_Total_AnnualAvg)
    print ">Looking at the Honeybee Program parameters:"
    print "   >Taking into account the airflow for 'areas' and the airflow for people."
    print "   >This is the value BEFORE any occupany schedule or operation sched is applied to reduce this (demand control)"
    print "      >Reference Zone Floor Area used is: {:.2f}".format(float(zoneFloorArea))
    print "      >Avg. Annual Zone Occupancy is: {:.2f} PPL".format(float(avgZoneOccupancy))
    print "      >[Ventilation Per Pers: {:.6f} m3/s-prs] x [Avg Zn Occ: {:.2f} ppl] * 3600 s/hr = {:.2f} m3/hr".format(ventilationPerPerson, avgZoneOccupancy, ventilationPerPerson*avgZoneOccupancy*3600)
    print "      >[Ventilation Per Area: {:.6f} m3/s-m2] x [Floor Area: {:.2f} m2] * 3600 s/hr = {:.2f} m3/hr".format(float(ventilationPerArea), float(zoneFloorArea), float(zoneVentilation_forArea_Avg))
    print "      >[Vent For Area: {:.2f} m3/h] + [Vent For PPL: {:.2f} m3/h] = {:.2f} m3/h".format(zoneVentilation_forArea_Avg, zoneVentilation_forPeople_Avg, zoneVentilation_Total_AnnualAvg)
    return zoneVentilation_Total_AnnualAvg, zoneVentilation_forArea_Avg, zoneVentilation_forPeople_Avg

def setRoomVentFlowRates(_HBzoneObj, _type, _zoneAnnualAvgVentRate):
    # --------------------------------------------------------------------------
    ###### Apply the Ventilation flow rate  the Zone's Rooms 

    # if the Zone has Rooms (has the 'PHPProoms' Dict)
    if 'PHPProoms' in _HBzoneObj.__dict__.keys():
        zoneRooms = _HBzoneObj.PHPProoms # Dict of all the rooms in the zone
        
        # Get the total Zone TFA
        totalZoneTFA = []
        for eachRoom in zoneRooms:
            totalZoneTFA.append(eachRoom.FloorArea_TFA)
        
        totalZoneTFA = sum(totalZoneTFA)
        
        if totalZoneTFA == 0:
            warning = "Zone: '{}' looks like it has no TFA surfaces? Check the TFA surface inputs to make sure.\nIt "\
            "should have at least one room floor surfce with valid TFA?".format(_HBzoneObj.name)
            ghenv.Component.AddRuntimeMessage(gh.GH_RuntimeMessageLevel.Warning, warning)
        
        
        for eachRoom in zoneRooms:
            # Compute the Room % of Total TFA and Room's Ventilation Airflows
            percentZoneTotalTFA = eachRoom.FloorArea_TFA / totalZoneTFA
            
            roomAirFlow = percentZoneTotalTFA * _zoneAnnualAvgVentRate
            
            # If the user says 'EP'
            if _type=='EP':
                setattr(eachRoom, 'V_sup', roomAirFlow) 
                setattr(eachRoom, 'V_eta', roomAirFlow)
                setattr(eachRoom, 'V_trans', roomAirFlow)
            elif _type=='UD':
                if eachRoom.V_sup == 'Automatic':
                    setattr(eachRoom, 'V_sup', roomAirFlow/2) # Divide by 2 cus' half goes to supply, half to extract?
                    setattr(eachRoom, 'V_eta', roomAirFlow/2)
                    setattr(eachRoom, 'V_trans', roomAirFlow/2)
                else:
                    # don't do anything. Leave the rooms as-is
                    pass

def setRoomVentSchedule(_HBzoneObj, _type, _userVentSched, _zoneVentilation_Total_AnnualAvg, _annualAvgZoneFlowRate_Area, _annualAvgZoneFlowRate_PPl, _zoneGrossFloorArea):
    # Pull the Ventilation loads, Occupancy Schedule from the Hive
    numOfPeoplePerArea, ventilationPerArea, ventilationPerPerson, zoneOccSchedAsValues = getHBLoadAndSched(_HBzoneObj)
    zoneFloorArea = getHBzoneFloorArea(_HBzoneObj, _zoneGrossFloorArea)
    
    # Convert the HB Sched Values for the zone as PHPP-Stly (bined 3)
    bins, vals = histogram(zoneOccSchedAsValues, 2)
    bined_Sched = namedtuple('phppSched', 'speed_high time_high speed_med time_med speed_low time_low')
    hbRoomVentSched = bined_Sched(vals[2], bins[2], vals[1], bins[1], vals[0], bins[0] )
    
    # if the Zone has Rooms (has the 'PHPProoms' Dict)
    if 'PHPProoms' in _HBzoneObj.__dict__.keys():
        zoneRooms = _HBzoneObj.PHPProoms # Dict of all the rooms in the zone
        
        # Get the Zone's total TFA
        totalZoneTFA = []
        for eachRoom in zoneRooms:
            totalZoneTFA.append(eachRoom.FloorArea_TFA)
        
        totalZoneTFA = sum(totalZoneTFA)
        
        # Go through each room
        for eachRoom in zoneRooms:
            if _userVentSched:
                # If there are user values supplied for the PHPP Schedule, use those 
                # And ignore any HB Library Schedules found
                phppRoomVentSched = _userVentSched[0]
                
                # Calc the Room's flow rates (m3/h) for People based on the PHPP-Style Schedule
                roomFlowrate_People_High = (phppRoomVentSched.speed_high * phppRoomVentSched.time_high * numOfPeoplePerArea * zoneFloorArea * ventilationPerPerson * 60 * 60) # m3/s---> m3/hour
                roomFlowrate_People_Med = (phppRoomVentSched.speed_med * phppRoomVentSched.time_med * numOfPeoplePerArea * zoneFloorArea * ventilationPerPerson * 60 * 60) # m3/s---> m3/hour
                roomFlowrate_People_Low = (phppRoomVentSched.speed_low * phppRoomVentSched.time_low * numOfPeoplePerArea * zoneFloorArea * ventilationPerPerson * 60 * 60) # m3/s---> m3/hour
                
                ventilationPerArea = (roomFlowrate_People_High + roomFlowrate_People_Med + roomFlowrate_People_Low)
            else:
                # Compute the Room % of Total TFA and Room's Ventilation Airflows
                percentZoneTotalTFA = eachRoom.FloorArea_TFA / totalZoneTFA
                roomFlowrate_People_Peak = numOfPeoplePerArea * zoneFloorArea * ventilationPerPerson * 60 * 60 * percentZoneTotalTFA
                
                # Calc total airflow for People
                # Calc the flow rates for people based on the HB Schedule values
                roomFlowrate_People_High = (hbRoomVentSched.speed_high * numOfPeoplePerArea * zoneFloorArea * ventilationPerPerson * 60 * 60 * percentZoneTotalTFA) 
                roomFlowrate_People_Med = (hbRoomVentSched.speed_med * numOfPeoplePerArea * zoneFloorArea * ventilationPerPerson * 60 * 60 * percentZoneTotalTFA) 
                roomFlowrate_People_Low = (hbRoomVentSched.speed_low * numOfPeoplePerArea * zoneFloorArea * ventilationPerPerson * 60 * 60 * percentZoneTotalTFA) 
                
                roomVentilationPerArea = (ventilationPerArea * zoneFloorArea * 60 * 60 * percentZoneTotalTFA)
                
                if roomVentilationPerArea != 0 and roomFlowrate_People_Peak != 0:
                    roomFlowrate_Total_High = (roomFlowrate_People_High + roomVentilationPerArea) / (roomFlowrate_People_Peak + roomVentilationPerArea)
                    roomFlowrate_Total_Med = (roomFlowrate_People_Med + roomVentilationPerArea) / (roomFlowrate_People_Peak + roomVentilationPerArea)
                    roomFlowrate_Total_Low = (roomFlowrate_People_Low + roomVentilationPerArea) / (roomFlowrate_People_Peak + roomVentilationPerArea)
                    
                    # Re-set the Room's Vent Schedule Fan-Speeds based on the calculated rates
                    # taking into account both Floor Area and People
                    phppRoomVentSched = bined_Sched(roomFlowrate_Total_High, bins[2], roomFlowrate_Total_Med, bins[1], roomFlowrate_Total_Low, bins[0] )
                else:
                    phppRoomVentSched = bined_Sched(1.0, 1.0, 0.0, 0.0, 0.0, 0.0)
                
                # Output the HB Schedules and loads as Annual Average 'Constant' flow rate (same as PHPP)
                roomVentilationPerArea = _zoneVentilation_Total_AnnualAvg / zoneFloorArea
                roomVentilationPerPerson = 0.0 # zero out, all the vent is applied as an average to the 'by area'
            
            # Apply the Zone's new PHPP-Stlye Schedule to the room
            setattr(eachRoom, 'phppVentSched', phppRoomVentSched)

def calcHBzoneVentRates(_HBzoneObj, _zoneGrossFloorArea):
    
    ventPerArea = 1
    if 'PHPProoms' not in _HBzoneObj.__dict__.keys():
        msg = 'To be able to properly calculate the airflow for rooms, please\n'\
        'use this component after the "PHPP Rooms From Rhino" component.'
        ghenv.Component.AddRuntimeMessage(gh.GH_RuntimeMessageLevel.Warning, msg)
        return 0, 0 # TODO. Find the a good default....
    
    tfaRoomsTotalAvgAirflow = 0.0
    tfaRoomsTotalFloorArea = 0.0
    
    sup = []
    eta = []
    trans = []
    for eachroom in _HBzoneObj.PHPProoms:
        # Supply
        sup_high = eachroom.V_sup * eachroom.phppVentSched.speed_high * eachroom.phppVentSched.time_high
        sup_med = eachroom.V_sup * eachroom.phppVentSched.speed_med * eachroom.phppVentSched.time_med
        sup_low = eachroom.V_sup * eachroom.phppVentSched.speed_low * eachroom.phppVentSched.time_low
        sup.append(sup_high + sup_med + sup_low)
        
        # Extract
        eta_high = eachroom.V_eta * eachroom.phppVentSched.speed_high * eachroom.phppVentSched.time_high
        eta_med = eachroom.V_eta * eachroom.phppVentSched.speed_med * eachroom.phppVentSched.time_med
        eta_low = eachroom.V_eta * eachroom.phppVentSched.speed_low * eachroom.phppVentSched.time_low
        eta.append(eta_high + eta_med + eta_low)
        
        #Trans
        trans_high = eachroom.V_trans * eachroom.phppVentSched.speed_high * eachroom.phppVentSched.time_high
        trans_med = eachroom.V_trans * eachroom.phppVentSched.speed_med * eachroom.phppVentSched.time_med
        trans_low = eachroom.V_trans * eachroom.phppVentSched.speed_low * eachroom.phppVentSched.time_low            
        trans.append( trans_high + trans_med + trans_low )
        
        # Room's Floor Area (Gross, not TFA)
        tfaRoomsTotalFloorArea += eachroom.FloorArea_Gross
    
    return sup, eta, trans

def preview(_zoneObjs):
    for zone in _zoneObjs:
        for eachroom in zone.PHPProoms:
            print '-----'
            print 'Room {}: {:.0f} m3/h, runs at {:.0f}% fan speed for {:.0f}% of the year'.format(eachroom.RoomName, eachroom.V_sup, eachroom.phppVentSched.speed_high*100, eachroom.phppVentSched.time_high*100)
            print 'Room {}: {:.0f} m3/h, runs at {:.0f}% fan speed for {:.0f}% of the year'.format(eachroom.RoomName, eachroom.V_sup, eachroom.phppVentSched.speed_med*100, eachroom.phppVentSched.time_med*100)
            print 'Room {}: {:.0f} m3/h, runs at {:.0f}% fan speed for {:.0f}% of the year'.format(eachroom.RoomName, eachroom.V_sup, eachroom.phppVentSched.speed_low*100, eachroom.phppVentSched.time_low*100)
            print '-----'

# ------------------------------------------------------------------------------
#### Get / Calculate the Zone's Ventilation Flow rates and Schedule
if len(_HBZones) > 0:
    HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)
    type = checkInputs(HBZoneObjects, airFlowInput_)
    
    supAirflow = []
    etaAirflow = []
    transAirflow = []
    zoneFloorArea = []
    for zone in HBZoneObjects:
        # 1) Figure out the Zone's Annual Average Ventilation Flow Rate
        #    (People + Area) based on HB Program (Load / Schedule)
        #
        # 2) Set the individual room Flow Rates from HB Program, or from Rhino Scene
        #
        # 3) Calc the HB Flow rates to override the Library and
        #    align the EP simulation with the PHPP
        
        print '- - '*25
        print 'Looking at Zone {}'.format(zone.name)
        
        (annualAvgZoneFlowRate,
        annualAvgZoneFlowRate_Area,
        annualAvgZoneFlowRate_PPl) = calcZoneAnnualVentFlowRateFromHB(zone, _phppVentSched, zoneGrossFloorArea_)
        
        setRoomVentFlowRates(zone,
                            type,
                            annualAvgZoneFlowRate)
        
        setRoomVentSchedule(zone,
                            type,
                            _phppVentSched,
                            annualAvgZoneFlowRate,
                            annualAvgZoneFlowRate_Area,
                            annualAvgZoneFlowRate_PPl,
                            zoneGrossFloorArea_)
        
        # Remember, you can't calc this on a zone-by-zone way. You need to find 
        # the TOTAL for ALL zones first, then you can calc the right overall airflows
        sup, eta, trans = calcHBzoneVentRates(zone, zoneGrossFloorArea_)
        supAirflow += sup
        etaAirflow += eta
        transAirflow += trans
        zoneFloorArea.append(getHBzoneFloorArea(zone, zoneGrossFloorArea_))
    
    #---------------------------------------------------------------------------
    # Figure out the Right airflow to use for the whole building
    # These values get passed back to Honeybee for the E+ model
    s = sum(supAirflow)
    e = sum(etaAirflow)
    t = sum(transAirflow)
    fa = sum(zoneFloorArea)
    
    ventilationPerArea_ = (max(s, e, t) / fa)/3600
    ventilationPerPerson_ = 0
    
    HBZones_ = _HBZones

# ------------------------------------------------------------------------------
#### Set the Zone's PHPP Ventilation System as Default
# If no systems already assigned, add in a Default Ventilation System.
# This can be overwriten by the user later
if len(HBZoneObjects) > 0:
    defaultVentSystem = PHPP_Sys_Ventilation() # Uses the class defaults
    for zone in HBZoneObjects:
        
        if getattr(zone, 'PHPP_VentSys'):
            continue
        
        setattr(zone, 'PHPP_VentSys', defaultVentSystem)
        for room in zone.PHPProoms:
            setattr(room, 'VentUnitName', defaultVentSystem.Unit_Name)
            setattr(room, 'VentSystemName', defaultVentSystem.SystemName)

# ------------------------------------------------------------------------------
# Modify the HB Zone's Floor Area parameter
if zoneGrossFloorArea_ != None:
    if len(HBZoneObjects)>0:
        for zone in HBZoneObjects:
            zoneFloorArea = getHBzoneFloorArea(zone, zoneGrossFloorArea_)
            setattr(zone, 'floorArea', zoneFloorArea)

# ------------------------------------------------------------------------------
# Add modified zones back to the HB dictionary
if len(_HBZones)>0:
    HBZones_  = hb_hive.addToHoneybeeHive(HBZoneObjects, ghenv.Component)
else:
    HBZones_ = _HBZones

preview(HBZoneObjects)