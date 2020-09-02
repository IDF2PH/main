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
Will calculate the PHPP Envelope Airtightness using the PHPP Rooms as the reference volume. Connect the ouputs from this component to a Honeybee 'setEPZoneLoads' and then set the Infiltration Schedule to 'CONSTANT'. Use a Honeybee 'Constant Schedule' with a value of 1 and a _schedTypeLimit of 'FRACTIONAL', then connect that to an HB 'setEPZoneSchdeules' component.
-
Note: The results shown here will be a fair bit different than the Honeybee 'ACH2m3/s-m2 Calculator' standard component because for PH Cert we are supposed to use the Net Internal Volume (v50) NOT the gross volume. E+ / HB use the gross volume and so given the same ACH, they will arrive at different infiltration flow rates (flow = ACH * Volume). For PH work, use this component.
-
EM Sept. 02, 2020

    Args:
        _HBZones: Honeybee Zones to apply this leakage rate to. Note, this should be the set of all the zones which were tested together as part of a Blower Door test. IE: if the blower door test included Zones A, B, and C then all three zones should be passed in here together. Use 'Merge' to combine zones if need be.
        _n50: (ACH) The target ACH leakage rate
        _q50: (m3/hr-m2-surface) The target leakage rate per m2 of exposed surface area
        _blowerPressure: (Pascal) Blower Door pressure for the airtightness measurement. Default is 50Pa
    Returns:
        _HBZones: Connect to the '_HBZones' input on the Honeybee 'setEPZoneLoads' Component
        infiltrationRatePerFloorArea_: (m3/hr-m2-floor)
        infiltrationRatePerFacadeArea_: (m3/hr-m2-facade) Connect to the 'infilRatePerArea_Facade_' input on the Honeybee 'setEPZoneLoads' Component
"""

ghenv.Component.Name = "BT_Airtightness"
ghenv.Component.NickName = "Airtightness"
ghenv.Component.Message = 'SEP_02_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc
import rhinoscriptsyntax as rs
import ghpythonlib.components as ghc
import Grasshopper.Kernel as ghK
import math

hb_hive = sc.sticky["honeybee_Hive"]()
HBZoneObjects = hb_hive.callFromHoneybeeHive(_HBZones)
infiltrationRatePerFloorArea_ = []
infiltrationRatePerFacadeArea_ = []

# Clean up inputs
n50 = _n50 if _n50 else None # ACH
q50 = _q50 if _q50 else None # m3/m2-facade
blowerPressure = _blowerPressure if _blowerPressure else 50.0 # Pa

for zone in HBZoneObjects:
    if 'PHPProoms' not in zone.__dict__.keys():
        msg = "Could not get the Volume for zone: '{}'. Be sure that you\n"\
        "Be sure that you use this component AFTER the 'PHPP Rooms from Rhino' component\n"\
        "and that you have valid 'rooms' in the model in order to determin the volume correctly.".format(zone.name)
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
        continue
    
    # --------------------------------------------------------------------------
    # Get all the relevant information needed from the zones
    zoneVolume_Vn50 = sum([room.RoomNetClearVolume for room in zone.PHPProoms])
    zoneFloorArea_Gross = zone.getFloorArea()
    zoneVolume_Gross = zone.getZoneVolume()
    zoneExposedSurfaceArea = zone.getExposedArea()
    
    if n50:
        zone_infil_airflow = zoneVolume_Vn50 * n50 / 3600
    elif q50:
        zone_infil_airflow = zoneExposedSurfaceArea * q50 / 3600
    else:
        zone_infil_airflow = 0.0
    
    print '- '*15, 'Zone: {}'.format(zone.name), '- '*15
    print 'INPUTS:'
    print '  >Zone PHPP Room Volumes (Vn50): {:.2f} m3'.format(zoneVolume_Vn50)
    print '  >Zone E+ Volume (Gross): {:.1f} m3'.format(zoneVolume_Gross)
    print '  >Zone E+ Floor Area (Gross): {:.1f} m2'.format(zoneFloorArea_Gross)
    print '  >Zone E+ Exposed Surface Area: {:.1f} m2'.format(zoneExposedSurfaceArea)
    print '  >Zone Infiltration Flowrate: {:.1f} m3/hr ({:.4f} m3/s) @ {}Pa'.format(zone_infil_airflow*60*60, zone_infil_airflow, _blowerPressure)
    
    # --------------------------------------------------------------------------
    # Decide which flow rate to use and Calc Standard
    # Flow Rate incorporating Blower Pressure
    # This equation comes from Honeybee. The HB Componet uses a standard pressure
    # at rest of 4 Pascals. 
    
    normalAvgPressure = 4 #Pa
    standardFlowRate = zone_infil_airflow/(math.pow((blowerPressure/normalAvgPressure),0.63)) # m3/s
    
    # --------------------------------------------------------------------------
    # Calc the Zone's Infiltration Rate in m3/hr-2 of floor area (zone gross)
    zoneinfilRatePerFloorArea = standardFlowRate /  zoneFloorArea_Gross  #m3/s---> m3/hr-m2
    infiltrationRatePerFloorArea_.append(zoneinfilRatePerFloorArea)
    
    zoneinfilRatePerFacadeArea = standardFlowRate /  zoneExposedSurfaceArea  #m3/s---> m3/hr-m2
    infiltrationRatePerFacadeArea_.append(zoneinfilRatePerFacadeArea)
    
    print 'RESULTS:'
    print '  >Zone Infiltration Flow Per Unit of Floor Area: {:.4f} m3/hr-m2 ({:.6f} m3/s-m2) @ {} Pa'.format(zoneinfilRatePerFloorArea*60*60, zoneinfilRatePerFloorArea, normalAvgPressure)
    print '  >Zone Infiltration Flow Per Unit of Facade Area: {:.4f} m3/hr-m2 ({:.6f} m3/s-m2) @ {} Pa'.format(zoneinfilRatePerFacadeArea*60*60, zoneinfilRatePerFacadeArea, normalAvgPressure)
    print '  >Zone Infiltration ACH: {:.4f} @ {} Pa'.format(standardFlowRate*60*60 / zoneVolume_Vn50, normalAvgPressure)

if _HBZones:
    HBZones_ = _HBZones