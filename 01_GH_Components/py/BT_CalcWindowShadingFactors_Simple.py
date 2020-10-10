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
Will calculate 'shading factor' dimensions for each window in the project. These are written to the 'Shading' worksheet. 
-
Reference: Shading factors go from 0 (fully shaded) to 1 (fully unshaded) and are calculated using the simplified numerical method as implemented in the Passive House Planning Package v9.6 and DesignPH 1.5 or earlier. 
Note that this method is much faster, but a bit less accurate than other methods you could use to determine shading factors. Its useful if you just want a quick picture of the shading condition though or if you are trying to match the exact procedure of an older style PHPP document.
For background and reference on the methodology used, see: "Solar Gains in a Passive House: A Monthly Approach to Calculating Global Irradiaton Entering a Shaded Window" By Andrew Peel, 2007.
-
EM October 10, 2020
    Args:
        runIt_: (bool) Set to 'True' to run the shading calcuation. May take a few seconds. 
        _latitude: (float) A value for the building's latitude. Use the Ladybug 'ImportEPW' to get this value.
        limit_: (int) Default=99 (m) A Value represening how for 'out' the window will look to find potential shading objects. Note, the further out it looks, the more potential shading obejcts it will find. But for any objects beyond 99m away, this simplified calculation method won't really be affected. Leave this set at the default unless you are sure about what you are doing.
        _HBZones: (list) The Honeybee Zones for analysis
        _windowSurrounds: (Tree) Each branch of the tree should represent one window. Each branch should have a list of 4 surfaces corresponding to the Bottom, Left, Top and Right window 'reveals' for windows which are inset into the wall or surface. Use the IDF2PH 'Create Window Reveals' component to automatically create this geometry.
        _bldgEnvelopeSrfcs: (list) The building (HB Zone) surfaces with the windows 'punched' out. Use the IDF2PH 'Create Window Reveals' component to automatically create this geometry.
        _shadingSrfcs: (list) <Optional> Any additional shading geometry (overhangs, neighbors, trees, etc...) you'd like to take into account when generating shading factors. Note that the more elements included, the slower this will run. 
    Returns:
        HBZones_: The updated Honeybee Zone objects to pass along to the next step.
        checklines_: Preview geometry showing the search lines used to find shading geometry.
        windowNames_: A list of the window names in the order calculated.
        winterShadingFactors_: A list of the winter shading factors calcualated. The order of this list matches the "windowNames_" output.
        summerShadingFactors_: A list of the winter shading factors calcualated. The order of this list matches the "windowNames_" output.
"""

ghenv.Component.Name = "BT_CalcWindowShadingFactors_Simple"
ghenv.Component.NickName = "Shading Factors | Simple"
ghenv.Component.Message = 'OCT_10_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import rhinoscriptsyntax as rs
import scriptcontext as sc
import Grasshopper.Kernel as ghk

hb_hive = sc.sticky["honeybee_Hive"]()
HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)

checklines_ = []
winterShadingFactors_ = []
summerShadingFactors_ = []
windowNames_ = []

if len(_HBZones)>0:
    try:
        latitude = float(_latitude)
    except:
       latitude = 40
       warning= "Please input the latitude for the project as a number. Use the Ladybug 'Import EPW' component\n"\
       "to get this value. For now I'll use a value of 40 (~NYC Latitude)."
       ghenv.Component.AddRuntimeMessage(ghk.GH_RuntimeMessageLevel.Warning, warning)

try:
    limit = float(limit_)
    if limit > 200:
        msg = 'Note: This simplified shading factor method does not really get \n'\
        'affected by anything more than 99 meters away from the window. Are you\n'\
        'sure you want to have the window look out that far?'
except:
    limit = 99

# Collect the shading object Geometry
shadingObjs = []
for each in _shadingSrfcs:
    shadingObjs.append( rs.coercebrep(each) )

for each in _bldgEnvelopeSrfcs:
    shadingObjs.append( rs.coercebrep(each) )

for branch in _windowSurrounds.Branches:
    for each in branch:
        shadingObjs.append( rs.coercebrep(each) )

# Calc the Shading Factors for each Window
if runIt_ and len(_HBZones)>0:
    for zone in HBZoneObjects:
        for srfc in zone.surfaces:
            if srfc.hasChild == False:
                continue
            
            for childSrfc in srfc.childSrfs:
                phppWindowObj = zone.phppWindowDict.get(childSrfc.name, None)
                dims = phppWindowObj.calcShadingDims_Simple( shadingObjs, limit)
                
                checklines_.append(phppWindowObj.ShadingDimensions.Horizon.checkline)
                checklines_.append(phppWindowObj.ShadingDimensions.Overhang.checkline)
                checklines_.append(phppWindowObj.ShadingDimensions.Reveal.checkline1)
                checklines_.append(phppWindowObj.ShadingDimensions.Reveal.checkline2)
                
                # Calculated same as PHPP, except I don't think they are calcing correctly yet...
                winter, summer = phppWindowObj.calcShadingFactor_Simple(latitude, shadingObjs, limit )
                windowNames_.append(childSrfc.name)
                winterShadingFactors_.append(winter)
                summerShadingFactors_.append(summer)

# Add modified Surfaces / Zones back to the HB dictionary
if len(_HBZones)>0:
    HBZones_  = hb_hive.addToHoneybeeHive(HBZoneObjects, ghenv.Component)