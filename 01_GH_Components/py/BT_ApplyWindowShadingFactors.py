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
Used to apply Winter and Summer Shading Factors to the Honeybee Model's windows. These shading factors can come from any source but should always be 0<-->1 with 0=fully shaded window and 1=fully unshaded window.
Note: Be sure that the order of the windowNames and the shading factors match each other.
-
EM June 28, 2020
    Args:
        _HBZones: (List)
        _windowNames: (List) The window names in the HB Model being analyzed. The order of this list should match the order of the Shading Factors input.
        _winterShadingFactors: (List) The winter shading factors (0=fully shaded, 1=fully unshaded). The length and order of this list should match the "_windowNames" input.
        _summerShadingFactors: (List) The summer shading factors (0=fully shaded, 1=fully unshaded). The length and order of this list should match the "_windowNames" input.
    Returns:
        HBZones_: The updated Honeybee Zone objects to pass along to the next step.
"""

ghenv.Component.Name = "BT_ApplyWindowShadingFactors"
ghenv.Component.NickName = "Apply Win. Shading Factors"
ghenv.Component.Message = 'JUN_28_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import Grasshopper.Kernel as ghK
import scriptcontext as sc

hb_hive = sc.sticky["honeybee_Hive"]()
HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)

if len(HBZoneObjects)>0 and len(_windowNames)>0:
    
    if len(_windowNames) != len(_winterShadingFactors) or len(_windowNames) != len(_summerShadingFactors):
        warning = "The number of windows in the HB Model doesn't match the shading factor inputs?\n"\
        "Check the HB Model / shading factor values and try again."
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
    
    for zone in HBZoneObjects:
        for i, windowName in enumerate(_windowNames):
            # Get the PHPP Window Object
            windowObj = zone.phppWindowDict.get(windowName, None)
            
            # Figure out the right Shading Factor to apply
            if windowObj != None:
                try: 
                    winterFac = _winterShadingFactors[i]
                except:
                    winterFac = 0.75
                    warning = "Warning: No Shading Factor found for window '{}'. Using 0.75 for winter.".format(windowName)
                    ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
                
                try:
                    summerFac = _summerShadingFactors[i]
                except:
                    summerFac = 0.75
                    warning = "Warning: No Shading Factor found for window '{}'. Using 0.75 for summer.".format(windowName)
                    ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
                
                # Set the Shading Factors of the object
                windowObj.setShadingFactors(winterFac, summerFac)
                
                #Reset the object in the zone's dict
                zone.phppWindowDict[ windowName ] = windowObj
                
                continue

# Warnings
if _winterShadingFactors.count(0) != 0:
    windowWithShadingFacOfZero = _windowNames[_winterShadingFactors.index(min(_winterShadingFactors))]
    warn = "Warning. One or more windows have a calculated shading factor of 0?\n"\
    "Thats probably not right. Double check window: {}. It might have its surface normal\n"\
    "backwards or something else funny is happening with your shading surfaces? Double check.".format(windowWithShadingFacOfZero)
    ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warn)

# Add modified Surfaces / Zones back to the HB dictionary
if len(_HBZones)>0:
    HBZones_  = hb_hive.addToHoneybeeHive(HBZoneObjects, ghenv.Component)



