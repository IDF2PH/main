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
Will filter the PHPP Objects based on the zone name(s) input. Useful if you are exporting multiple PHPP documents (for mixed use / mlti-zone models). Connect this after the 'IDF-->PHPP Objs' component and before the 'Create Excel Obj - Geom'. You can either choose to pass a list of zones to include, or a list of zones to exclude.
-
EM July 10, 2020

    Args:
        _PHPPObjs: (Tree) The PHPP Objects from the 'PHPPObjs_' output on the  'IDF-->PHPP Objs' component
        zonesInclude_: (list) The name or names of the zones to include in the set
        zoneExclude_: (list) The name or names of the zones to exclude from the set.
    Returns:
        ZoneGeom_: The Rhino geometry for the included zone(s), so you can Preview and check that its the right one(s).
        PHPPObjs_: The PHPP Objects for the included zone(s). Pass along to the Excel writer to output to the PHPP.
"""

ghenv.Component.Name = "BT_FilterPHPPObjs"
ghenv.Component.NickName = "Filter PHPP Objs"
ghenv.Component.Message = 'JUL_10_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "02 | IDF2PHPP"

from System import Object
from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path

def filterName(_zoneName, _zonesNamesToFilter):
    flag = True
    for each in _zonesNamesToFilter:
        if each in _zoneName:
            flag = False
    
    return flag

PHPPObjs_ = DataTree[Object]()
ZoneGeom_ =  DataTree[Object]()

##########################################
# Sort out which zones to include in the output
# Filter out any zones not to include
if _PHPPObjs.BranchCount>0:
    zones = [x.ZoneName for x in _PHPPObjs.Branch(8)]
    if len(zonesInclude_)>0: 
        zones = [x for x in zones if not filterName(x, zonesInclude_)]
    if len(zoneExclude_)>0:
        zones = [x for x in zones if filterName(x, zoneExclude_)]
    print('Inlcuding Zones {} in the Export'.format(zones))

if _PHPPObjs.BranchCount >0:
    for srfc in _PHPPObjs.Branch(4):
        if srfc.HostZoneName in zones:
            PHPPObjs_.Add(srfc, GH_Path(4))
            ZoneGeom_.Add(srfc.Srfc, GH_Path(4))
    
    for tfa in _PHPPObjs.Branch(6):
        if tfa.HostZoneName in zones:
            PHPPObjs_.Add(tfa, GH_Path(6))
    
    for zoneObj in _PHPPObjs.Branch(8):
        if zoneObj.ZoneName in zones:
            PHPPObjs_.Add(zoneObj, GH_Path(8))
    
    for dhw in _PHPPObjs.Branch(10):
        if True in [dhwZone in zones for dhwZone in dhw.ZonesAssigned]:
            PHPPObjs_.Add(dhw, GH_Path(10))
    
    for grndSrfc in _PHPPObjs.Branch(11):
        if grndSrfc.Zone in zones:
            PHPPObjs_.Add(grndSrfc, GH_Path(11))
    
    # All of this gets output, regardless of the zones being filtered in/out
    passthrough = [0, 1, 2, 3, 5, 7, 9]
    for branchNum in passthrough:
        PHPPObjs_.AddRange(_PHPPObjs.Branch(branchNum), GH_Path(branchNum))
