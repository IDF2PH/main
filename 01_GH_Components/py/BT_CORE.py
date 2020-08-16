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
Core Classes and Definitiions for IDF2PHPP Exporter. You must run this component before anything else will work. If you are having trouble when opening a GH file for the first time, try hitting 'Recompute'.
-
EM Augsut 16, 2020
"""

print '''Copyright (c) 2020, bldgtyp, llc <info@bldgtyp.com> 
>> This exporter works with Honeybee / Ladybug in order to read, interpret and export an EnergyPlus IDF file into a Passive House Planning Package (PHPP) file.
>> You do NOT need to excecute the EnergyPlus simulation for this run. This interpreter simply reads the input (IDF) file created by Honeybee.
>> You will need Honeybee as well as a Copy of the PHPP for this to work. This interpreter does NOT replace the PHPP. It just exports to it.
>> This has been tested and works with:
    :: Rhino Version 6 SR21 (6.21.19351.9141, 12/17/2019)
    :: PHPP Version 9.6a <SI Verison>
>> Note that this does NOT work with the IP version of the PHPP.
'''

ghenv.Component.Name = "BT_CORE"
ghenv.Component.NickName = "IDF2PHPP"
ghenv.Component.Message = 'AUG_16_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "00 | Core"

import scriptcontext as sc
import ghpythonlib.components as ghc
import rhinoscriptsyntax as rs
import Grasshopper.Kernel as ghK
import math
import Rhino
import json
import random
import re
from contextlib import contextmanager


#-------------------------------------------------------------------------------
##########    From HB    ###########
hb_EPMaterialAUX = sc.sticky["honeybee_EPMaterialAUX"]()


#Data
def getClimateData():
    return [
    {'Dataset': 'AD0001a-Andorra de la Vella', 'Comments': '2015 PHI', 'Code': '0001a', 'Latitude': '42.51', 'Source': 'Meteonorm & EOSWEB satellite data. Load data derived by PHI. ', 'Location': 'Andorra de la Vella', 'Longitude': '1.52', 'Region': '', 'Country': 'AD'},
    {'Dataset': 'AE0001a-Dubai', 'Comments': '2015 PHI & ZEPHIR', 'Code': '0001a', 'Latitude': '25.25', 'Source': 'Different sources; comparison & work by ZEPHIR Passivhaus Italia & PHI', 'Location': 'Dubai', 'Longitude': '55.33', 'Region': '', 'Country': 'AE'},
    {'Dataset': 'AT0001a-Eisenstadt', 'Comments': '2007 PHI. ', 'Code': '0001a', 'Latitude': '47.85', 'Source': 'source: ZAMG', 'Location': 'Eisenstadt', 'Longitude': '16.53', 'Region': 'Burgenland', 'Country': 'AT'},
    {'Dataset': 'AT0002a-Kleinzicken', 'Comments': '2007 PHI. ', 'Code': '0002a', 'Latitude': '47.22', 'Source': 'source: ZAMG', 'Location': 'Kleinzicken', 'Longitude': '16.33', 'Region': 'Burgenland', 'Country': 'AT'},
    {'Dataset': 'AT0003a-Neusiedl am See', 'Comments': '2007 PHI. ', 'Code': '0003a', 'Latitude': '47.95', 'Source': 'source: ZAMG', 'Location': 'Neusiedl am See', 'Longitude': '16.87', 'Region': 'Burgenland', 'Country': 'AT'},
    {'Dataset': u'AT0004a-Feldkirchen/K\xc3\xa4rnten', 'Comments': '2007 PHI. ', 'Code': '0004a', 'Latitude': '46.72', 'Source': 'source: ZAMG', 'Location': u'Feldkirchen/K\xc3\xa4rnten', 'Longitude': '14.1', 'Region': u'K\xc3\xa4rnten', 'Country': 'AT'},
    {'Dataset': u'AT0005a-K\xc3\xb6tschach-Mauthen', 'Comments': '2007 PHI. ', 'Code': '0005a', 'Latitude': '46.68', 'Source': 'source: ZAMG', 'Location': u'K\xc3\xb6tschach-Mauthen', 'Longitude': '13', 'Region': u'K\xc3\xa4rnten', 'Country': 'AT'},
    {'Dataset': 'AT0006a-Spittal/Drau', 'Comments': '2007 PHI. ', 'Code': '0006a', 'Latitude': '46.79', 'Source': 'source: ZAMG', 'Location': 'Spittal/Drau', 'Longitude': '13.49', 'Region': u'K\xc3\xa4rnten', 'Country': 'AT'},
    {'Dataset': u'AT0007a-St. Andr\xc3\xa4/Lavanttal', 'Comments': '2007 PHI. ', 'Code': '0007a', 'Latitude': '46.76', 'Source': 'source: ZAMG', 'Location': u'St. Andr\xc3\xa4/Lavanttal', 'Longitude': '14.83', 'Region': u'K\xc3\xa4rnten', 'Country': 'AT'},
    {'Dataset': 'AT0008a-Weissensee', 'Comments': '2007 PHI. ', 'Code': '0008a', 'Latitude': '46.72', 'Source': 'source: ZAMG', 'Location': 'Weissensee', 'Longitude': '13.29', 'Region': u'K\xc3\xa4rnten', 'Country': 'AT'},
    {'Dataset': 'AT0009a-Amstetten', 'Comments': '2007 PHI. ', 'Code': '0009a', 'Latitude': '48.11', 'Source': 'source: ZAMG', 'Location': 'Amstetten', 'Longitude': '14.9', 'Region': u'Nieder\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0010a-Baden', 'Comments': '2007 PHI. ', 'Code': '0010a', 'Latitude': '48.01', 'Source': 'source: ZAMG', 'Location': 'Baden', 'Longitude': '16.25', 'Region': u'Nieder\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0011a-Lilienfeld', 'Comments': '2007 PHI. ', 'Code': '0011a', 'Latitude': '48.03', 'Source': 'source: ZAMG', 'Location': 'Lilienfeld', 'Longitude': '15.58', 'Region': u'Nieder\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0012a-Stockerau', 'Comments': '2007 PHI. ', 'Code': '0012a', 'Latitude': '48.4', 'Source': 'source: ZAMG', 'Location': 'Stockerau', 'Longitude': '16.19', 'Region': u'Nieder\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0013a-Zwettl', 'Comments': '2007 PHI. ', 'Code': '0013a', 'Latitude': '48.62', 'Source': 'source: ZAMG', 'Location': 'Zwettl', 'Longitude': '15.2', 'Region': u'Nieder\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0014a-Bad Goisern', 'Comments': '2007 PHI. ', 'Code': '0014a', 'Latitude': '47.64', 'Source': 'source: ZAMG', 'Location': 'Bad Goisern', 'Longitude': '13.62', 'Region': u'Ober\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0015a-Gmunden', 'Comments': '2007 PHI. ', 'Code': '0015a', 'Latitude': '47.9', 'Source': 'source: ZAMG', 'Location': 'Gmunden', 'Longitude': '13.78', 'Region': u'Ober\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': u'AT0016a-K\xc3\xb6nigswiesen', 'Comments': '2007 PHI. ', 'Code': '0016a', 'Latitude': '48.41', 'Source': 'source: ZAMG', 'Location': u'K\xc3\xb6nigswiesen', 'Longitude': '14.84', 'Region': u'Ober\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0017a-Linz', 'Comments': '2007 PHI. ', 'Code': '0017a', 'Latitude': '48.32', 'Source': 'source: ZAMG', 'Location': 'Linz', 'Longitude': '14.3', 'Region': u'Ober\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0018a-Ried/Innkreis', 'Comments': '2007 PHI. ', 'Code': '0018a', 'Latitude': '48.22', 'Source': 'source: ZAMG', 'Location': 'Ried/Innkreis', 'Longitude': '13.48', 'Region': u'Ober\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': u'AT0019a-Rohrbach/M\xc3\xbchlkreis', 'Comments': '2007 PHI. ', 'Code': '0019a', 'Latitude': '48.57', 'Source': 'source: ZAMG', 'Location': u'Rohrbach/M\xc3\xbchlkreis', 'Longitude': '14', 'Region': u'Ober\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0020a-Weyer', 'Comments': '2007 PHI. ', 'Code': '0020a', 'Latitude': '47.86', 'Source': 'source: ZAMG', 'Location': 'Weyer', 'Longitude': '14.67', 'Region': u'Ober\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0021a-Windischgarsten', 'Comments': '2007 PHI. ', 'Code': '0021a', 'Latitude': '47.73', 'Source': 'source: ZAMG', 'Location': 'Windischgarsten', 'Longitude': '14.33', 'Region': u'Ober\xc3\xb6sterreich', 'Country': 'AT'},
    {'Dataset': 'AT0022a-Mattsee', 'Comments': '2007 PHI. ', 'Code': '0022a', 'Latitude': '47.98', 'Source': 'source: ZAMG', 'Location': 'Mattsee', 'Longitude': '13.11', 'Region': 'Salzburg', 'Country': 'AT'},
    {'Dataset': 'AT0023a-Salzburg', 'Comments': '2007 PHI. ', 'Code': '0023a', 'Latitude': '47.78', 'Source': 'source: ZAMG', 'Location': 'Salzburg', 'Longitude': '13.05', 'Region': 'Salzburg', 'Country': 'AT'},
    {'Dataset': 'AT0024a-St. Johann/Pongau', 'Comments': '2007 PHI. ', 'Code': '0024a', 'Latitude': '47.32', 'Source': 'source: ZAMG', 'Location': 'St. Johann/Pongau', 'Longitude': '13.18', 'Region': 'Salzburg', 'Country': 'AT'},
    {'Dataset': 'AT0025a-Tamsweg', 'Comments': '2007 PHI. ', 'Code': '0025a', 'Latitude': '47.12', 'Source': 'source: ZAMG', 'Location': 'Tamsweg', 'Longitude': '13.81', 'Region': 'Salzburg', 'Country': 'AT'},
    {'Dataset': 'AT0026a-Zell am See', 'Comments': '2007 PHI. ', 'Code': '0026a', 'Latitude': '47.33', 'Source': 'source: ZAMG', 'Location': 'Zell am See', 'Longitude': '12.8', 'Region': 'Salzburg', 'Country': 'AT'},
    {'Dataset': 'AT0027a-Aigen/Ennstal', 'Comments': '2007 PHI. ', 'Code': '0027a', 'Latitude': '47.53', 'Source': 'source: ZAMG', 'Location': 'Aigen/Ennstal', 'Longitude': '14.13', 'Region': 'Steiermark', 'Country': 'AT'},
    {'Dataset': 'AT0028a-Bad Gleichenberg', 'Comments': '2007 PHI. ', 'Code': '0028a', 'Latitude': '46.88', 'Source': 'source: ZAMG', 'Location': 'Bad Gleichenberg', 'Longitude': '15.91', 'Region': 'Steiermark', 'Country': 'AT'},
    {'Dataset': 'AT0029a-Graz', 'Comments': '2007 PHI. ', 'Code': '0029a', 'Latitude': '47.08', 'Source': 'source: ZAMG', 'Location': 'Graz', 'Longitude': '15.37', 'Region': 'Steiermark', 'Country': 'AT'},
    {'Dataset': 'AT0030a-Kapfenberg', 'Comments': '2007 PHI. ', 'Code': '0030a', 'Latitude': '47.45', 'Source': 'source: ZAMG', 'Location': 'Kapfenberg', 'Longitude': '15.3', 'Region': 'Steiermark', 'Country': 'AT'},
    {'Dataset': 'AT0031a-Ramsau/Dachstein', 'Comments': '2007 PHI. ', 'Code': '0031a', 'Latitude': '47.42', 'Source': 'source: ZAMG', 'Location': 'Ramsau/Dachstein', 'Longitude': '13.63', 'Region': 'Steiermark', 'Country': 'AT'},
    {'Dataset': 'AT0032b-Innsbruck', 'Comments': '2014 PHI. Vergleich mit Meteonorm & IWEC. PassREg. ', 'Code': '0032b', 'Latitude': '47.26', 'Source': 'source: ZAMG', 'Location': 'Innsbruck', 'Longitude': '11.38', 'Region': 'Tirol', 'Country': 'AT'},
    {'Dataset': 'AT0033a-Kirchberg/Tirol', 'Comments': '2007 PHI. ', 'Code': '0033a', 'Latitude': '47.45', 'Source': 'source: ZAMG', 'Location': 'Kirchberg/Tirol', 'Longitude': '12.32', 'Region': 'Tirol', 'Country': 'AT'},
    {'Dataset': u'AT0034a-S\xc3\xb6lden', 'Comments': '2007 PHI. ', 'Code': '0034a', 'Latitude': '46.97', 'Source': 'source: ZAMG', 'Location': u'S\xc3\xb6lden', 'Longitude': '11.01', 'Region': 'Tirol', 'Country': 'AT'},
    {'Dataset': 'AT0035a-Stams', 'Comments': '2007 PHI. ', 'Code': '0035a', 'Latitude': '47.28', 'Source': 'source: ZAMG', 'Location': 'Stams', 'Longitude': '10.98', 'Region': 'Tirol', 'Country': 'AT'},
    {'Dataset': 'AT0036a-Weissenbach/Lech', 'Comments': '2007 PHI. ', 'Code': '0036a', 'Latitude': '47.44', 'Source': 'source: ZAMG', 'Location': 'Weissenbach/Lech', 'Longitude': '10.64', 'Region': 'Tirol', 'Country': 'AT'},
    {'Dataset': u'AT0037a-W\xc3\xb6rgl', 'Comments': '2007 PHI. ', 'Code': '0037a', 'Latitude': '47.49', 'Source': 'source: ZAMG', 'Location': u'W\xc3\xb6rgl', 'Longitude': '12.07', 'Region': 'Tirol', 'Country': 'AT'},
    {'Dataset': 'AT0038a-Zell am Ziller', 'Comments': '2007 PHI. ', 'Code': '0038a', 'Latitude': '47.25', 'Source': 'source: ZAMG', 'Location': 'Zell am Ziller', 'Longitude': '11.9', 'Region': 'Tirol', 'Country': 'AT'},
    {'Dataset': 'AT0039a-Alberschwende', 'Comments': '2007 PHI. ', 'Code': '0039a', 'Latitude': '47.46', 'Source': 'source: ZAMG', 'Location': 'Alberschwende', 'Longitude': '9.85', 'Region': 'Vorarlberg', 'Country': 'AT'},
    {'Dataset': 'AT0040a-Bregenz', 'Comments': '2007 PHI. ', 'Code': '0040a', 'Latitude': '47.5', 'Source': 'source: ZAMG', 'Location': 'Bregenz', 'Longitude': '9.75', 'Region': 'Vorarlberg', 'Country': 'AT'},
    {'Dataset': 'AT0041a-Dornbirn', 'Comments': '2007 PHI. ', 'Code': '0041a', 'Latitude': '47.43', 'Source': 'source: ZAMG', 'Location': 'Dornbirn', 'Longitude': '9.73', 'Region': 'Vorarlberg', 'Country': 'AT'},
    {'Dataset': 'AT0042a-Feldkirch', 'Comments': '2007 PHI. ', 'Code': '0042a', 'Latitude': '47.27', 'Source': 'source: ZAMG', 'Location': 'Feldkirch', 'Longitude': '9.6', 'Region': 'Vorarlberg', 'Country': 'AT'},
    {'Dataset': 'AT0043a-Warth', 'Comments': '2007 PHI. ', 'Code': '0043a', 'Latitude': '47.25', 'Source': 'source: ZAMG', 'Location': 'Warth', 'Longitude': '10.18', 'Region': 'Vorarlberg', 'Country': 'AT'},
    {'Dataset': u'AT0044a-Wien Ost (Gro\xc3\x9f-Enzersdorf)', 'Comments': '2007 PHI. ', 'Code': '0044a', 'Latitude': '48.2', 'Source': 'source: ZAMG', 'Location': u'Wien Ost (Gro\xc3\x9f-Enzersdorf)', 'Longitude': '16.57', 'Region': 'Wien', 'Country': 'AT'},
    {'Dataset': 'AT0045a-Wien-Donaufeld', 'Comments': '2007 PHI. ', 'Code': '0045a', 'Latitude': '48.26', 'Source': 'source: ZAMG', 'Location': 'Wien-Donaufeld', 'Longitude': '16.43', 'Region': 'Wien', 'Country': 'AT'},
    {'Dataset': 'AT0046a-Wien-Hohe Warte', 'Comments': '2007 PHI. ', 'Code': '0046a', 'Latitude': '48.25', 'Source': 'source: ZAMG', 'Location': 'Wien-Hohe Warte', 'Longitude': '16.37', 'Region': 'Wien', 'Country': 'AT'},
    {'Dataset': 'AT0047a-Wien-Innere Stadt', 'Comments': '2007 PHI. ', 'Code': '0047a', 'Latitude': '48.2', 'Source': 'source: ZAMG', 'Location': 'Wien-Innere Stadt', 'Longitude': '16.37', 'Region': 'Wien', 'Country': 'AT'},
    {'Dataset': 'AT0048a-Imst', 'Comments': '2015 PHI', 'Code': '0048a', 'Latitude': '47.25', 'Source': 'Temp = 1981-2010; Strahlung / Lastdaten basierend auf ZAMG', 'Location': 'Imst', 'Longitude': '10.74', 'Region': 'Tirol', 'Country': 'AT'},
    {'Dataset': 'AU0001a-Melbourne', 'Comments': '2011 PHI. Vergleich mit EOSWEB, IWEC.', 'Code': '0001a', 'Latitude': '-37.817', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)  Source: NatHERS (A) RMY data 2012.', 'Location': 'Melbourne', 'Longitude': '144.967', 'Region': 'Victoria', 'Country': 'AU'},
    {'Dataset': 'AU0002a-Perth', 'Comments': '2014 PHI', 'Code': '0002a', 'Latitude': '-31.93', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)  Source: NatHERS (A) RMY data 2012.', 'Location': 'Perth', 'Longitude': '115.98', 'Region': 'Western Australia', 'Country': 'AU'},
    {'Dataset': 'AU0003a-Canberra', 'Comments': '2014 PHI', 'Code': '0003a', 'Latitude': '-35.31', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!) . Source: NatHERS (A) RMY data 2012.', 'Location': 'Canberra', 'Longitude': '149.2', 'Region': 'Australian Capital Territory', 'Country': 'AU'},
    {'Dataset': 'AU0004a-Adelaide', 'Comments': '2014 PHI', 'Code': '0004a', 'Latitude': '-34.93', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!) . Monthly values: MN7 Adelaide (WS, new period). Compared with TMY2 NatHERS and IWEC.', 'Location': 'Adelaide', 'Longitude': '138.53', 'Region': 'South Australia', 'Country': 'AU'},
    {'Dataset': 'AU0005a-Hobart', 'Comments': '2015 PHI', 'Code': '0005a', 'Latitude': '-42.89', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!). Temp = 1981-2010; Other = derived from NatHERS (A) RMY data 2012.', 'Location': 'Hobart', 'Longitude': '147.33', 'Region': 'Tasmania', 'Country': 'AU'},
    {'Dataset': 'AU0006a-Applethorpe', 'Comments': '2015 PHI', 'Code': '0006a', 'Latitude': '-28.61', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!). Source = Meteonorm 7, new period. CL = NatHERS Armidale. PHI January 2015', 'Location': 'Applethorpe', 'Longitude': '151.95', 'Region': 'Queensland', 'Country': 'AU'},
    {'Dataset': 'AU0007a-Armidale', 'Comments': '2015 PHI', 'Code': '0007a', 'Latitude': '-30.53', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!). Source = Meteonorm 7, new period. CL = NatHERS Armidale. PHI January 2015', 'Location': 'Armidale', 'Longitude': '151.62', 'Region': 'New South Wales', 'Country': 'AU'},
    {'Dataset': 'AU0008a-Toowoomba', 'Comments': '2016 PHI', 'Code': '0008a', 'Latitude': '-27.54', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!). Temp = 2006-2015; Other derived from Meteonorm. ', 'Location': 'Toowoomba', 'Longitude': '151.91', 'Region': 'Queensland', 'Country': 'AU'},
    {'Dataset': 'AU0009a-Oakey', 'Comments': '2016 PHI', 'Code': '0009a', 'Latitude': '-27.4', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!). Temp = 1981-2010; Other derived from Meteonorm and NatHERS. ', 'Location': 'Oakey', 'Longitude': '151.74', 'Region': 'Queensland', 'Country': 'AU'},
    {'Dataset': u'BE0001c-Br\xc3\xbcssel (Ukkel)', 'Comments': '2014 PHI. Load data from IWEC. PassREg', 'Code': '0001c', 'Latitude': '50.78', 'Source': 'Source: Temperature & Dew Point from KMI (1981-2010). Radiation: MN7 Station (1968-2005).', 'Location': u'Br\xc3\xbcssel (Ukkel)', 'Longitude': '4.35', 'Region': '', 'Country': 'BE'},
    {'Dataset': 'BE0002b-Saint-Hubert', 'Comments': '', 'Code': '0002b', 'Latitude': '50.017', 'Source': '', 'Location': 'Saint-Hubert', 'Longitude': '5.317', 'Region': '', 'Country': 'BE'},
    {'Dataset': 'BE0003b-Oostende', 'Comments': '', 'Code': '0003b', 'Latitude': '51.217', 'Source': '', 'Location': 'Oostende', 'Longitude': '2.917', 'Region': '', 'Country': 'BE'},
    {'Dataset': 'BE0004b-Florennes', 'Comments': '', 'Code': '0004b', 'Latitude': '50.25', 'Source': '', 'Location': 'Florennes', 'Longitude': '4.7', 'Region': '', 'Country': 'BE'},
    {'Dataset': 'BE0005b-Elsenhorn', 'Comments': '', 'Code': '0005b', 'Latitude': '50.45', 'Source': '', 'Location': 'Elsenhorn', 'Longitude': '6.217', 'Region': '', 'Country': 'BE'},
    {'Dataset': 'BE0006b-Limbourg', 'Comments': '', 'Code': '0006b', 'Latitude': '50.617', 'Source': '', 'Location': 'Limbourg', 'Longitude': '5.933', 'Region': '', 'Country': 'BE'},
    {'Dataset': 'BG0001a-Varna', 'Comments': '2011 PHI. Vergleich mit EOSWEB.', 'Code': '0001a', 'Latitude': '43.21', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Varna', 'Longitude': '27.91', 'Region': 'Climate Zone 1', 'Country': 'BG'},
    {'Dataset': 'BG0002a-Shumen', 'Comments': '2011 PHI. Vergleich mit EOSWEB.', 'Code': '0002a', 'Latitude': '43.283', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Shumen', 'Longitude': '26.933', 'Region': 'Climate Zone 2', 'Country': 'BG'},
    {'Dataset': 'BG0003a-Ruse', 'Comments': '2011 PHI. Vergleich mit EOSWEB.', 'Code': '0003a', 'Latitude': '43.856', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Ruse', 'Longitude': '25.971', 'Region': 'Climate Zone 3', 'Country': 'BG'},
    {'Dataset': 'BG0004b-Veliko Tarnovo', 'Comments': '2011 PHI. Confirmed 2014 by comparison with data from national regulation. PassREg', 'Code': '0004b', 'Latitude': '43.086', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Veliko Tarnovo', 'Longitude': '25.656', 'Region': 'Climate Zone 4', 'Country': 'BG'},
    {'Dataset': 'BG0005b-Burgas', 'Comments': '2011 PHI. Confirmed 2014 by comparison with data from national regulation. PassREg', 'Code': '0005b', 'Latitude': '42.51', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Burgas', 'Longitude': '27.47', 'Region': 'Climate Zone 5', 'Country': 'BG'},
    {'Dataset': 'BG0006a-Plovdiv', 'Comments': '2011 PHI. Vergleich mit EOSWEB.', 'Code': '0006a', 'Latitude': '42.15', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Plovdiv', 'Longitude': '24.75', 'Region': 'Climate Zone 6', 'Country': 'BG'},
    {'Dataset': 'BG0007a-Sofia', 'Comments': '2011 PHI. Vergleich mit EOSWEB.', 'Code': '0007a', 'Latitude': '42.697', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Sofia', 'Longitude': '23.323', 'Region': 'Climate Zone 7', 'Country': 'BG'},
    {'Dataset': 'BG0008a-Haskovo', 'Comments': '2011 PHI. Vergleich mit EOSWEB.', 'Code': '0008a', 'Latitude': '41.933', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Haskovo', 'Longitude': '25.567', 'Region': 'Climate Zone 8', 'Country': 'BG'},
    {'Dataset': 'BG0009a-Blagoevgrad', 'Comments': '2011 PHI. Vergleich mit EOSWEB.', 'Code': '0009a', 'Latitude': '42.014', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Blagoevgrad', 'Longitude': '23.095', 'Region': 'Climate Zone 9', 'Country': 'BG'},
    {'Dataset': 'BR0001b-Brasilia', 'Comments': '2011 PHI', 'Code': '0001b', 'Latitude': '-15.78', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)  Source: Meteonorm V6. ', 'Location': 'Brasilia', 'Longitude': '-47.93', 'Region': '', 'Country': 'BR'},
    {'Dataset': 'BY0001a-Minsk', 'Comments': '', 'Code': '0001a', 'Latitude': '53.9', 'Source': '', 'Location': 'Minsk', 'Longitude': '27.5', 'Region': '', 'Country': 'BY'},
    {'Dataset': 'CA0001b-Toronto', 'Comments': 'CanPHI dataset', 'Code': '0001b', 'Latitude': '43.661', 'Source': '', 'Location': 'Toronto', 'Longitude': '-79.383', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': u'CA0002c-Montr\xc3\xa9al', 'Comments': u'2015 PHI. Erg\xc3\xa4nzt um 2 K\xc3\xbchllastdaten, sonst identisch zu Vorg\xc3\xa4nger.', 'Code': '0002c', 'Latitude': '45.51', 'Source': 'source: Meteonorm. Cooling load: PHI', 'Location': u'Montr\xc3\xa9al', 'Longitude': '-73.56', 'Region': u'Qu\xc3\xa9bec', 'Country': 'CA'},
    {'Dataset': 'CA0003d-Vancouver', 'Comments': '2015 PHI', 'Code': '0003d', 'Latitude': '49.2', 'Source': 'Temp: Environment Canada data (Climate Normal 1981-2010). Radiation & load data: Derived from Meteonorm & CWEC. ', 'Location': 'Vancouver', 'Longitude': '-123.18', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0004b-Ottawa', 'Comments': 'CanPHI dataset', 'Code': '0004b', 'Latitude': '45.412', 'Source': '', 'Location': 'Ottawa', 'Longitude': '-75.699', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0005b-Calgary', 'Comments': 'CanPHI dataset', 'Code': '0005b', 'Latitude': '51.046', 'Source': '', 'Location': 'Calgary', 'Longitude': '-114.061', 'Region': 'Alberta', 'Country': 'CA'},
    {'Dataset': 'CA0006b-Edmonton', 'Comments': 'CanPHI dataset', 'Code': '0006b', 'Latitude': '53.541', 'Source': '', 'Location': 'Edmonton', 'Longitude': '-113.494', 'Region': 'Alberta', 'Country': 'CA'},
    {'Dataset': 'CA0007b-Quebec', 'Comments': 'CanPHI dataset', 'Code': '0007b', 'Latitude': '46.816', 'Source': '', 'Location': 'Quebec', 'Longitude': '-71.224', 'Region': u'Qu\xc3\xa9bec', 'Country': 'CA'},
    {'Dataset': 'CA0008b-Winnipeg', 'Comments': 'CanPHI dataset', 'Code': '0008b', 'Latitude': '49.896', 'Source': '', 'Location': 'Winnipeg', 'Longitude': '-97.143', 'Region': 'Manitoba', 'Country': 'CA'},
    {'Dataset': 'CA0009b-St. Catherines', 'Comments': 'CanPHI dataset', 'Code': '0009b', 'Latitude': '43.183', 'Source': '', 'Location': 'St. Catherines', 'Longitude': '-79.233', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0010b-Halifax', 'Comments': 'CanPHI dataset', 'Code': '0010b', 'Latitude': '44.648', 'Source': '', 'Location': 'Halifax', 'Longitude': '-63.572', 'Region': 'Nova Scotia', 'Country': 'CA'},
    {'Dataset': 'CA0011b-Saskatoon', 'Comments': 'CanPHI dataset', 'Code': '0011b', 'Latitude': '52.129', 'Source': '', 'Location': 'Saskatoon', 'Longitude': '-106.662', 'Region': 'Saskatchewan', 'Country': 'CA'},
    {'Dataset': 'CA0012b-Regina', 'Comments': 'CanPHI dataset', 'Code': '0012b', 'Latitude': '50.447', 'Source': '', 'Location': 'Regina', 'Longitude': '-104.618', 'Region': 'Saskatchewan', 'Country': 'CA'},
    {'Dataset': 'CA0013b-Sherbrooke', 'Comments': 'CanPHI dataset', 'Code': '0013b', 'Latitude': '45.401', 'Source': '', 'Location': 'Sherbrooke', 'Longitude': '-71.888', 'Region': u'Qu\xc3\xa9bec', 'Country': 'CA'},
    {'Dataset': "CA0014b-St. John's", 'Comments': 'CanPHI dataset', 'Code': '0014b', 'Latitude': '47.567', 'Source': '', 'Location': "St. John's", 'Longitude': '-52.705', 'Region': 'Newfoundland', 'Country': 'CA'},
    {'Dataset': 'CA0015b-Kelowna', 'Comments': 'CanPHI dataset', 'Code': '0015b', 'Latitude': '49.882', 'Source': '', 'Location': 'Kelowna', 'Longitude': '-119.455', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0016b-Thunder Bay', 'Comments': 'CanPHI dataset', 'Code': '0016b', 'Latitude': '48.383', 'Source': '', 'Location': 'Thunder Bay', 'Longitude': '-89.25', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0017b-Saint John', 'Comments': 'CanPHI dataset', 'Code': '0017b', 'Latitude': '45.268', 'Source': '', 'Location': 'Saint John', 'Longitude': '-66.055', 'Region': 'New Brunswick', 'Country': 'CA'},
    {'Dataset': 'CA0018b-Prince George', 'Comments': 'CanPHI dataset', 'Code': '0018b', 'Latitude': '53.916', 'Source': '', 'Location': 'Prince George', 'Longitude': '-122.75', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0019b-Charlottetown', 'Comments': 'CanPHI dataset', 'Code': '0019b', 'Latitude': '46.233', 'Source': '', 'Location': 'Charlottetown', 'Longitude': '-63.133', 'Region': 'Prince Edward Island', 'Country': 'CA'},
    {'Dataset': 'CA0020b-Yellowknife', 'Comments': 'CanPHI dataset', 'Code': '0020b', 'Latitude': '62.444', 'Source': '', 'Location': 'Yellowknife', 'Longitude': '-114.396', 'Region': 'Northwest Territories', 'Country': 'CA'},
    {'Dataset': 'CA0021a-Smithers', 'Comments': '2014 PHI', 'Code': '0021a', 'Latitude': '54.82', 'Source': 'Source: CWEC', 'Location': 'Smithers', 'Longitude': '-127.18', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0022a-Whistler', 'Comments': 'CanPHI dataset', 'Code': '0022a', 'Latitude': '50.121', 'Source': '', 'Location': 'Whistler', 'Longitude': '-122.954', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0023b-Upper Squamish', 'Comments': '2015 PHI', 'Code': '0023b', 'Latitude': '49.9', 'Source': 'Temp: Environment Canada data (Climate Normal 1981-2010). Radiation & load data: Derived from Meteonorm.', 'Location': 'Upper Squamish', 'Longitude': '-123.28', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0024a-Cranbrook', 'Comments': 'CanPHI dataset', 'Code': '0024a', 'Latitude': '49.514', 'Source': '', 'Location': 'Cranbrook', 'Longitude': '-115.769', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0025a-Victoria', 'Comments': 'CanPHI dataset', 'Code': '0025a', 'Latitude': '48.42', 'Source': '', 'Location': 'Victoria', 'Longitude': '-123.37', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0026a-Fort St.John', 'Comments': 'CanPHI dataset', 'Code': '0026a', 'Latitude': '56.247', 'Source': '', 'Location': 'Fort St.John', 'Longitude': '-120.848', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0027a-Medicine Hat', 'Comments': 'CanPHI dataset', 'Code': '0027a', 'Latitude': '50.042', 'Source': '', 'Location': 'Medicine Hat', 'Longitude': '-110.678', 'Region': 'Alberta', 'Country': 'CA'},
    {'Dataset': 'CA0028a-Banff', 'Comments': 'CanPHI dataset', 'Code': '0028a', 'Latitude': '51.178', 'Source': '', 'Location': 'Banff', 'Longitude': '-115.572', 'Region': 'Alberta', 'Country': 'CA'},
    {'Dataset': 'CA0029a-Edson', 'Comments': 'CanPHI dataset', 'Code': '0029a', 'Latitude': '53.582', 'Source': '', 'Location': 'Edson', 'Longitude': '-116.434', 'Region': 'Alberta', 'Country': 'CA'},
    {'Dataset': 'CA0030a-Grand Prairie', 'Comments': 'CanPHI dataset', 'Code': '0030a', 'Latitude': '55.171', 'Source': '', 'Location': 'Grand Prairie', 'Longitude': '-118.795', 'Region': 'Alberta', 'Country': 'CA'},
    {'Dataset': 'CA0031a-Lethbridge', 'Comments': 'CanPHI dataset', 'Code': '0031a', 'Latitude': '49.694', 'Source': '', 'Location': 'Lethbridge', 'Longitude': '-112.833', 'Region': 'Alberta', 'Country': 'CA'},
    {'Dataset': 'CA0032a-Swift Current', 'Comments': 'CanPHI dataset', 'Code': '0032a', 'Latitude': '50.288', 'Source': '', 'Location': 'Swift Current', 'Longitude': '-107.794', 'Region': 'Saskatchewan', 'Country': 'CA'},
    {'Dataset': 'CA0033a-Yorkton', 'Comments': 'CanPHI dataset', 'Code': '0033a', 'Latitude': '51.214', 'Source': '', 'Location': 'Yorkton', 'Longitude': '-102.463', 'Region': 'Saskatchewan', 'Country': 'CA'},
    {'Dataset': 'CA0034a-Prince Albert', 'Comments': 'CanPHI dataset', 'Code': '0034a', 'Latitude': '53.2', 'Source': '', 'Location': 'Prince Albert', 'Longitude': '-105.75', 'Region': 'Saskatchewan', 'Country': 'CA'},
    {'Dataset': 'CA0035a-Brandon', 'Comments': 'CanPHI dataset', 'Code': '0035a', 'Latitude': '49.833', 'Source': '', 'Location': 'Brandon', 'Longitude': '-99.95', 'Region': 'Manitoba', 'Country': 'CA'},
    {'Dataset': 'CA0036a-Peterborough', 'Comments': 'CanPHI dataset', 'Code': '0036a', 'Latitude': '44.3', 'Source': '', 'Location': 'Peterborough', 'Longitude': '-78.317', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0037a-Windsor', 'Comments': 'CanPHI dataset', 'Code': '0037a', 'Latitude': '42.3', 'Source': '', 'Location': 'Windsor', 'Longitude': '-83.02', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0038a-London', 'Comments': 'CanPHI dataset', 'Code': '0038a', 'Latitude': '42.97', 'Source': '', 'Location': 'London', 'Longitude': '-81.25', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0039a-Kingston', 'Comments': 'CanPHI dataset', 'Code': '0039a', 'Latitude': '44.23', 'Source': '', 'Location': 'Kingston', 'Longitude': '-76.5', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0040a-Sudbury', 'Comments': 'CanPHI dataset', 'Code': '0040a', 'Latitude': '46.5', 'Source': '', 'Location': 'Sudbury', 'Longitude': '-81.02', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0041a-Sault Ste. Marie', 'Comments': 'CanPHI dataset', 'Code': '0041a', 'Latitude': '46.533', 'Source': '', 'Location': 'Sault Ste. Marie', 'Longitude': '-84.35', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0042a-Parry Sound', 'Comments': 'CanPHI dataset', 'Code': '0042a', 'Latitude': '45.333', 'Source': '', 'Location': 'Parry Sound', 'Longitude': '-80.033', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0043a-Saguenay', 'Comments': 'CanPHI dataset', 'Code': '0043a', 'Latitude': '48.417', 'Source': '', 'Location': 'Saguenay', 'Longitude': '-71.067', 'Region': u'Qu\xc3\xa9bec', 'Country': 'CA'},
    {'Dataset': 'CA0044a-Maniwaki', 'Comments': 'CanPHI dataset', 'Code': '0044a', 'Latitude': '46.383', 'Source': '', 'Location': 'Maniwaki', 'Longitude': '-75.983', 'Region': u'Qu\xc3\xa9bec', 'Country': 'CA'},
    {'Dataset': 'CA0045a-Baie Comeau', 'Comments': 'CanPHI dataset', 'Code': '0045a', 'Latitude': '49.217', 'Source': '', 'Location': 'Baie Comeau', 'Longitude': '-68.15', 'Region': u'Qu\xc3\xa9bec', 'Country': 'CA'},
    {'Dataset': 'CA0046a-Shawinigan', 'Comments': 'CanPHI dataset', 'Code': '0046a', 'Latitude': '46.55', 'Source': '', 'Location': 'Shawinigan', 'Longitude': '-72.733', 'Region': u'Qu\xc3\xa9bec', 'Country': 'CA'},
    {'Dataset': 'CA0047a-Fredericton', 'Comments': 'CanPHI dataset', 'Code': '0047a', 'Latitude': '45.95', 'Source': '', 'Location': 'Fredericton', 'Longitude': '-66.667', 'Region': 'New Brunswick', 'Country': 'CA'},
    {'Dataset': 'CA0048a-Moncton', 'Comments': 'CanPHI dataset', 'Code': '0048a', 'Latitude': '46.07', 'Source': '', 'Location': 'Moncton', 'Longitude': '-64.77', 'Region': 'New Brunswick', 'Country': 'CA'},
    {'Dataset': 'CA0049a-Bathurst', 'Comments': 'CanPHI dataset', 'Code': '0049a', 'Latitude': '47.62', 'Source': '', 'Location': 'Bathurst', 'Longitude': '-65.65', 'Region': 'New Brunswick', 'Country': 'CA'},
    {'Dataset': 'CA0050a-Truro', 'Comments': 'CanPHI dataset', 'Code': '0050a', 'Latitude': '45.365', 'Source': '', 'Location': 'Truro', 'Longitude': '-63.28', 'Region': 'Nova Scotia', 'Country': 'CA'},
    {'Dataset': 'CA0051a-Sydney', 'Comments': 'CanPHI dataset', 'Code': '0051a', 'Latitude': '46.138', 'Source': '', 'Location': 'Sydney', 'Longitude': '-60.183', 'Region': 'Nova Scotia', 'Country': 'CA'},
    {'Dataset': 'CA0052a-Kentville', 'Comments': 'CanPHI dataset', 'Code': '0052a', 'Latitude': '45.078', 'Source': '', 'Location': 'Kentville', 'Longitude': '-64.496', 'Region': 'Nova Scotia', 'Country': 'CA'},
    {'Dataset': 'CA0053a-Yarmouth', 'Comments': 'CanPHI dataset', 'Code': '0053a', 'Latitude': '43.836', 'Source': '', 'Location': 'Yarmouth', 'Longitude': '-66.118', 'Region': 'Nova Scotia', 'Country': 'CA'},
    {'Dataset': 'CA0054a-Corner Brook', 'Comments': 'CanPHI dataset', 'Code': '0054a', 'Latitude': '48.95', 'Source': '', 'Location': 'Corner Brook', 'Longitude': '-57.95', 'Region': 'Newfoundland', 'Country': 'CA'},
    {'Dataset': 'CA0055a-Whitehorse', 'Comments': 'CanPHI dataset', 'Code': '0055a', 'Latitude': '60.717', 'Source': '', 'Location': 'Whitehorse', 'Longitude': '-135.05', 'Region': 'Yukon', 'Country': 'CA'},
    {'Dataset': 'CA0056a-Dawson', 'Comments': 'CanPHI dataset', 'Code': '0056a', 'Latitude': '64.06', 'Source': '', 'Location': 'Dawson', 'Longitude': '-139.411', 'Region': 'Yukon', 'Country': 'CA'},
    {'Dataset': 'CA0057a-Campbell Island, Dryad Point', 'Comments': '2015 PHI', 'Code': '0057a', 'Latitude': '52.19', 'Source': 'Source: Environment Canada data (Canadian Climate Normals 1981-2010). Radiation & Load data: satellite data (Passipedia). T', 'Location': 'Campbell Island, Dryad Point', 'Longitude': '-128.11', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0058a-Nelson', 'Comments': u'2014 PHI. Keine K\xc3\xbchllast. Compared with satellite and Government of Canada normal data.', 'Code': '0058a', 'Latitude': '49.5', 'Source': 'Source:  Meteonorm', 'Location': 'Nelson', 'Longitude': '-117.3', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0059a-Blue River', 'Comments': '2015 PHI', 'Code': '0059a', 'Latitude': '52.13', 'Source': 'Temp: Environment Canada data (Climate Normal 1981-2010). Radiation & load data: based on MN7. ', 'Location': 'Blue River', 'Longitude': '-119.29', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0060a-Minden', 'Comments': '2015 PHI. ', 'Code': '0060a', 'Latitude': '44.93', 'Source': 'Source: Environment Canada data (Canadian Climate Normals 1981-2010). Radiation & Load data: Meteonorm / PHI.', 'Location': 'Minden', 'Longitude': '-78.72', 'Region': 'Ontario', 'Country': 'CA'},
    {'Dataset': 'CA0061a-Kamloops', 'Comments': '2015 PHI', 'Code': '0061a', 'Latitude': '50.7', 'Source': 'Temp: Environment Canada data (Climate Normal 1981-2010). Radiation & load data: based on MN7 / CWEC. ', 'Location': 'Kamloops', 'Longitude': '-120.44', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0062a-Merritt', 'Comments': '2015 PHI', 'Code': '0062a', 'Latitude': '50.11', 'Source': 'Temp: Environment Canada data (Climate Normal 1981-2010). Radiation & load data: approximation.', 'Location': 'Merritt', 'Longitude': '-120.8', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0063a-Abbotsford', 'Comments': '2015 PHI', 'Code': '0063a', 'Latitude': '49.03', 'Source': 'Temp: Environment Canada data (Climate Normal 1981-2010). Radiation & load data: Derived from Meteonorm. ', 'Location': 'Abbotsford', 'Longitude': '-122.36', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CA0064a-Hope', 'Comments': '2015 PHI', 'Code': '0064a', 'Latitude': '49.37', 'Source': 'Derived from Meteonorm. ', 'Location': 'Hope', 'Longitude': '-121.48', 'Region': 'British Columbia', 'Country': 'CA'},
    {'Dataset': 'CH0001a-Altdorf', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0001a', 'Latitude': '46.87', 'Source': '', 'Location': 'Altdorf', 'Longitude': '8.63', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0002a-Basel (Binningen)', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0002a', 'Latitude': '47.55', 'Source': '', 'Location': 'Basel (Binningen)', 'Longitude': '7.58', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0003a-Bern (Liebefeld)', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0003a', 'Latitude': '46.93', 'Source': '', 'Location': 'Bern (Liebefeld)', 'Longitude': '7.42', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0004a-Chur', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0004a', 'Latitude': '46.87', 'Source': '', 'Location': 'Chur', 'Longitude': '9.53', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0005a-Davos', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0005a', 'Latitude': '46.82', 'Source': '', 'Location': 'Davos', 'Longitude': '9.85', 'Region': '', 'Country': 'CH'},
    {'Dataset': u'CH0006a-Gen\xc3\xa8ve (Cointrin)', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0006a', 'Latitude': '46.25', 'Source': '', 'Location': u'Gen\xc3\xa8ve (Cointrin)', 'Longitude': '6.13', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0007a-Glarus', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0007a', 'Latitude': '47.03', 'Source': '', 'Location': 'Glarus', 'Longitude': '9.07', 'Region': '', 'Country': 'CH'},
    {'Dataset': u'CH0008a-G\xc3\xbcttingen', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0008a', 'Latitude': '47.6', 'Source': '', 'Location': u'G\xc3\xbcttingen', 'Longitude': '9.28', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0009a-Interlaken', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0009a', 'Latitude': '46.67', 'Source': '', 'Location': 'Interlaken', 'Longitude': '7.87', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0010a-La Chaux de Fonds', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0010a', 'Latitude': '47.08', 'Source': '', 'Location': 'La Chaux de Fonds', 'Longitude': '6.8', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0011a-Locarno', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0011a', 'Latitude': '46.17', 'Source': '', 'Location': 'Locarno', 'Longitude': '8.88', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0012a-Lugano', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0012a', 'Latitude': '46', 'Source': '', 'Location': 'Lugano', 'Longitude': '8.97', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0013a-Luzern', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0013a', 'Latitude': '47.03', 'Source': '', 'Location': 'Luzern', 'Longitude': '8.3', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0014a-Montana', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0014a', 'Latitude': '46.32', 'Source': '', 'Location': 'Montana', 'Longitude': '7.48', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0015a-Payerne', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0015a', 'Latitude': '46.82', 'Source': '', 'Location': 'Payerne', 'Longitude': '6.95', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0016a-Pully', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0016a', 'Latitude': '46.52', 'Source': '', 'Location': 'Pully', 'Longitude': '6.67', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0017a-St. Moritz', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0017a', 'Latitude': '46.53', 'Source': '', 'Location': 'St. Moritz', 'Longitude': '9.88', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0018a-Sion', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0018a', 'Latitude': '46.22', 'Source': '', 'Location': 'Sion', 'Longitude': '7.33', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0019a-St. Gallen', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0019a', 'Latitude': '47.43', 'Source': '', 'Location': 'St. Gallen', 'Longitude': '9.4', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CH0020a-Wynau', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0020a', 'Latitude': '47.25', 'Source': '', 'Location': 'Wynau', 'Longitude': '7.78', 'Region': '', 'Country': 'CH'},
    {'Dataset': u'CH0021a-Z\xc3\xbcrich', 'Comments': '2007 PHI. TRY konvertiert mit Meteonorm', 'Code': '0021a', 'Latitude': '47.43', 'Source': '', 'Location': u'Z\xc3\xbcrich', 'Longitude': '8.55', 'Region': '', 'Country': 'CH'},
    {'Dataset': 'CL0001a-Santiago de Chile', 'Comments': '', 'Code': '0001a', 'Latitude': '-33.383', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)  ', 'Location': 'Santiago de Chile', 'Longitude': '-70.783', 'Region': '', 'Country': 'CL'},
    {'Dataset': 'CN0001a-Shanghai', 'Comments': '2011 PHI', 'Code': '0001a', 'Latitude': '31.4', 'Source': 'source: Meteonorm V6', 'Location': 'Shanghai', 'Longitude': '121.4', 'Region': 'Shanghai', 'Country': 'CN'},
    {'Dataset': 'CN0002a-Beijing', 'Comments': '2011 PHI. Vergleich mit EOSWEB, SWERA, IWEC.', 'Code': '0002a', 'Latitude': '39.933', 'Source': 'source: Meteonorm V6', 'Location': 'Beijing', 'Longitude': '116.283', 'Region': 'Beijing', 'Country': 'CN'},
    {'Dataset': u'CN0003a-\xc3\x9cr\xc3\xbcmqi', 'Comments': '2010 PHI. Vergleich mit Daten von EOSWEB & China Meteorological Administration. ', 'Code': '0003a', 'Latitude': '43.8', 'Source': 'source: Meteonorm V6', 'Location': u'\xc3\x9cr\xc3\xbcmqi', 'Longitude': '87.58', 'Region': 'Xinjiang', 'Country': 'CN'},
    {'Dataset': 'CN0004a-Fuzhou', 'Comments': '2016 PHI', 'Code': '0004a', 'Latitude': '26.08', 'Source': 'Temp = 1981-2010, Other = Meteonorm. ', 'Location': 'Fuzhou', 'Longitude': '119.28', 'Region': 'Fujian', 'Country': 'CN'},
    {'Dataset': 'CN0005a-Harbin', 'Comments': '2016 PHI', 'Code': '0005a', 'Latitude': '45.75', 'Source': 'Temp = 1981-2010; Other derived from Meteonorm V7 & CSWD data. ', 'Location': 'Harbin', 'Longitude': '126.77', 'Region': 'Heilongjiang', 'Country': 'CN'},
    {'Dataset': 'CN0006a-Yanji', 'Comments': '2016 PHI', 'Code': '0006a', 'Latitude': '42.87', 'Source': 'Temp = 1981-2010; Other = derived from Meteonorm', 'Location': 'Yanji', 'Longitude': '129.5', 'Region': 'Jilin', 'Country': 'CN'},
    {'Dataset': 'CN0007a-Songjianghezhen', 'Comments': '2016 PHI', 'Code': '0007a', 'Latitude': '42.18', 'Source': 'Derived from Meteonorm & CSWD. ', 'Location': 'Songjianghezhen', 'Longitude': '127.48', 'Region': 'Jilin', 'Country': 'CN'},
    {'Dataset': 'CN0008a-Guangzhou', 'Comments': '2016 PHI', 'Code': '0008a', 'Latitude': '23.17', 'Source': 'Derived from CSWD and Meteonorm V7. ', 'Location': 'Guangzhou', 'Longitude': '113.33', 'Region': 'Guangdong', 'Country': 'CN'},
    {'Dataset': 'CN0009a-Chengdu', 'Comments': '2016 PHI', 'Code': '0009a', 'Latitude': '30.67', 'Source': 'Derived from Meteonorm', 'Location': 'Chengdu', 'Longitude': '104.02', 'Region': 'Sichuan', 'Country': 'CN'},
    {'Dataset': 'CN0010a-Lhasa', 'Comments': '2016 PHI', 'Code': '0010a', 'Latitude': '29.67', 'Source': 'Temp = 1981-2010. Other derived from Meteonorm & CSWD ', 'Location': 'Lhasa', 'Longitude': '91.13', 'Region': 'Tibet', 'Country': 'CN'},
    {'Dataset': 'CN0011a-Kunming', 'Comments': '2016 PHI', 'Code': '0011a', 'Latitude': '25.02', 'Source': 'Temp = 1971-2000; Other derived from Meteonorm (version7) and CSWD data', 'Location': 'Kunming', 'Longitude': '102.68', 'Region': 'Yunnan', 'Country': 'CN'},
    {'Dataset': 'CN0012a-Qionghai', 'Comments': '2016 PHI', 'Code': '0012a', 'Latitude': '19.23', 'Source': 'Temp = 1981-2010; Other derived from CSWD', 'Location': 'Qionghai', 'Longitude': '110.47', 'Region': 'Hainan', 'Country': 'CN'},
    {'Dataset': 'CN0013a-Tianjin', 'Comments': '2016 PHI', 'Code': '0013a', 'Latitude': '39.1', 'Source': 'Temp = 1981-2010;  Other derived from CSWD & Meteonorm V7. ', 'Location': 'Tianjin', 'Longitude': '117.17', 'Region': 'Tianjin', 'Country': 'CN'},
    {'Dataset': 'CZ0001a-Praha', 'Comments': '', 'Code': '0001a', 'Latitude': '50.1', 'Source': '', 'Location': 'Praha', 'Longitude': '14.43', 'Region': '', 'Country': 'CZ'},
    {'Dataset': 'CZ0002a-Brno', 'Comments': '', 'Code': '0002a', 'Latitude': '49.22', 'Source': '', 'Location': 'Brno', 'Longitude': '16.7', 'Region': '', 'Country': 'CZ'},
    {'Dataset': u'CZ0003a-\xc4\x8cesk\xc3\xa9 Bud\xc4\x9bjovice', 'Comments': '', 'Code': '0003a', 'Latitude': '48.9', 'Source': '', 'Location': u'\xc4\x8cesk\xc3\xa9 Bud\xc4\x9bjovice', 'Longitude': '14.5', 'Region': '', 'Country': 'CZ'},
    {'Dataset': u'CZ0004a-Hradec Kr\xc3\xa1lov\xc3\xa9', 'Comments': '', 'Code': '0004a', 'Latitude': '50.183', 'Source': '', 'Location': u'Hradec Kr\xc3\xa1lov\xc3\xa9', 'Longitude': '15.833', 'Region': '', 'Country': 'CZ'},
    {'Dataset': 'CZ0005a-Jihlava', 'Comments': '', 'Code': '0005a', 'Latitude': '49.24', 'Source': '', 'Location': 'Jihlava', 'Longitude': '15.5', 'Region': '', 'Country': 'CZ'},
    {'Dataset': 'CZ0006a-Karlovy Vary', 'Comments': '', 'Code': '0006a', 'Latitude': '50.2', 'Source': '', 'Location': 'Karlovy Vary', 'Longitude': '12.9', 'Region': '', 'Country': 'CZ'},
    {'Dataset': 'CZ0007a-Liberec', 'Comments': '', 'Code': '0007a', 'Latitude': '50.8', 'Source': '', 'Location': 'Liberec', 'Longitude': '15.08', 'Region': '', 'Country': 'CZ'},
    {'Dataset': 'CZ0008a-Olomouc', 'Comments': '', 'Code': '0008a', 'Latitude': '49.63', 'Source': '', 'Location': 'Olomouc', 'Longitude': '17.25', 'Region': '', 'Country': 'CZ'},
    {'Dataset': 'CZ0009a-Ostrava', 'Comments': '', 'Code': '0009a', 'Latitude': '49.83', 'Source': '', 'Location': 'Ostrava', 'Longitude': '18.25', 'Region': '', 'Country': 'CZ'},
    {'Dataset': u'CZ0010a-Plze\xc5\x88', 'Comments': '', 'Code': '0010a', 'Latitude': '49.75', 'Source': '', 'Location': u'Plze\xc5\x88', 'Longitude': '13.42', 'Region': '', 'Country': 'CZ'},
    {'Dataset': 'CZ0012a-Znojmo', 'Comments': '', 'Code': '0012a', 'Latitude': '48.88', 'Source': '', 'Location': 'Znojmo', 'Longitude': '16.08', 'Region': '', 'Country': 'CZ'},
    {'Dataset': 'DE0001a-Norderney', 'Comments': 'Niedersachsen', 'Code': '0001a', 'Latitude': '53.71', 'Source': 'source: DIN 4108-6:2003. Klimaregion 1', 'Location': 'Norderney', 'Longitude': '7.15', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0002a-Husum', 'Comments': 'Schleswig-Holstein', 'Code': '0002a', 'Latitude': '54.48', 'Source': 'source: DIN 4108-6:2003. Klimaregion 1', 'Location': 'Husum', 'Longitude': '9.06', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0003a-Hamburg', 'Comments': 'Hamburg', 'Code': '0003a', 'Latitude': '53.64', 'Source': 'source: DIN 4108-6:2003. Klimaregion 2', 'Location': 'Hamburg', 'Longitude': '9.99', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0004a-Hannover', 'Comments': 'Niedersachsen', 'Code': '0004a', 'Latitude': '52.47', 'Source': 'source: DIN 4108-6:2003. Klimaregion 2', 'Location': 'Hannover', 'Longitude': '9.68', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0005a-Kiel', 'Comments': 'Schleswig-Holstein', 'Code': '0005a', 'Latitude': '54.34', 'Source': 'source: DIN 4108-6:2003. Klimaregion 2', 'Location': 'Kiel', 'Longitude': '10.09', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0006a-Arkona', 'Comments': 'Mecklenburg-Vorpommern', 'Code': '0006a', 'Latitude': '54.68', 'Source': 'source: DIN 4108-6:2003. Klimaregion 3', 'Location': 'Arkona', 'Longitude': '13.44', 'Region': '', 'Country': 'DE'},
    {'Dataset': u'DE0007a-Warnem\xc3\xbcnde', 'Comments': 'Mecklenburg-Vorpommern', 'Code': '0007a', 'Latitude': '54.18', 'Source': 'source: DIN 4108-6:2003. Klimaregion 3', 'Location': u'Warnem\xc3\xbcnde', 'Longitude': '12.08', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0008a-Potsdam', 'Comments': 'Brandenburg', 'Code': '0008a', 'Latitude': '52.38', 'Source': 'source: DIN 4108-6:2003. Klimaregion 4', 'Location': 'Potsdam', 'Longitude': '13.06', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0009a-Schwerin', 'Comments': 'Mecklenburg-Vorpommern', 'Code': '0009a', 'Latitude': '53.64', 'Source': 'source: DIN 4108-6:2003. Klimaregion 4', 'Location': 'Schwerin', 'Longitude': '11.39', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0010a-Teterow', 'Comments': 'Mecklenburg-Vorpommern', 'Code': '0010a', 'Latitude': '53.76', 'Source': 'source: DIN 4108-6:2003. Klimaregion 4', 'Location': 'Teterow', 'Longitude': '12.56', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0011a-Braunschweig', 'Comments': 'Niedersachsen', 'Code': '0011a', 'Latitude': '52.29', 'Source': 'source: DIN 4108-6:2003. Klimaregion 5', 'Location': 'Braunschweig', 'Longitude': '10.45', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0012a-Dresden', 'Comments': 'Sachsen', 'Code': '0012a', 'Latitude': '51.02', 'Source': 'source: DIN 4108-6:2003. Klimaregion 5', 'Location': 'Dresden', 'Longitude': '13.78', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0013a-Wittenberg', 'Comments': 'Sachsen-Anhalt', 'Code': '0013a', 'Latitude': '51.89', 'Source': 'source: DIN 4108-6:2003. Klimaregion 5', 'Location': 'Wittenberg', 'Longitude': '12.65', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0014a-Erfurt', 'Comments': u'Th\xc3\xbcringen', 'Code': '0014a', 'Latitude': '50.98', 'Source': 'source: DIN 4108-6:2003. Klimaregion 6', 'Location': 'Erfurt', 'Longitude': '10.96', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0015a-Harzgerode', 'Comments': 'Sachsen-Anhalt', 'Code': '0015a', 'Latitude': '51.65', 'Source': 'source: DIN 4108-6:2003. Klimaregion 6', 'Location': 'Harzgerode', 'Longitude': '11.14', 'Region': '', 'Country': 'DE'},
    {'Dataset': u'DE0016a-L\xc3\xbcdenscheid', 'Comments': 'Nordrhein-Westfalen', 'Code': '0016a', 'Latitude': '51.25', 'Source': 'source: DIN 4108-6:2003. Klimaregion 6', 'Location': u'L\xc3\xbcdenscheid', 'Longitude': '7.64', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0017a-Essen', 'Comments': 'Nordrhein-Westfalen', 'Code': '0017a', 'Latitude': '51.41', 'Source': 'source: DIN 4108-6:2003. Klimaregion 7', 'Location': 'Essen', 'Longitude': '6.97', 'Region': '', 'Country': 'DE'},
    {'Dataset': u'DE0018a-K\xc3\xb6ln', 'Comments': 'Nordrhein-Westfalen', 'Code': '0018a', 'Latitude': '50.87', 'Source': 'source: DIN 4108-6:2003. Klimaregion 7', 'Location': u'K\xc3\xb6ln', 'Longitude': '7.16', 'Region': '', 'Country': 'DE'},
    {'Dataset': u'DE0019a-M\xc3\xbcnster', 'Comments': 'Nordrhein-Westfalen', 'Code': '0019a', 'Latitude': '51.95', 'Source': 'source: DIN 4108-6:2003. Klimaregion 7', 'Location': u'M\xc3\xbcnster', 'Longitude': '7.59', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0020a-Geisenheim', 'Comments': 'Hessen', 'Code': '0020a', 'Latitude': '49.99', 'Source': 'source: DIN 4108-6:2003. Klimaregion 8', 'Location': 'Geisenheim', 'Longitude': '7.95', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0021a-Kassel', 'Comments': 'Hessen', 'Code': '0021a', 'Latitude': '51.3', 'Source': 'source: DIN 4108-6:2003. Klimaregion 8', 'Location': 'Kassel', 'Longitude': '9.44', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0022a-Trier', 'Comments': 'Rheinland-Pfalz', 'Code': '0022a', 'Latitude': '49.75', 'Source': 'source: DIN 4108-6:2003. Klimaregion 8', 'Location': 'Trier', 'Longitude': '6.65', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0023a-Chemnitz', 'Comments': 'Sachsen', 'Code': '0023a', 'Latitude': '50.79', 'Source': 'source: DIN 4108-6:2003. Klimaregion 9', 'Location': 'Chemnitz', 'Longitude': '12.87', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0024a-Leipzig', 'Comments': 'Sachsen', 'Code': '0024a', 'Latitude': '51.39', 'Source': 'source: DIN 4108-6:2003. Klimaregion 9', 'Location': 'Leipzig', 'Longitude': '12.4', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0025a-Cham', 'Comments': 'Bayern', 'Code': '0025a', 'Latitude': '49.24', 'Source': 'source: DIN 4108-6:2003. Klimaregion 10', 'Location': 'Cham', 'Longitude': '12.62', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0026a-Hof', 'Comments': 'Bayern', 'Code': '0026a', 'Latitude': '50.31', 'Source': 'source: DIN 4108-6:2003. Klimaregion 10', 'Location': 'Hof', 'Longitude': '11.88', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0027a-Freudenstadt', 'Comments': u'Baden-W\xc3\xbcrttemberg', 'Code': '0027a', 'Latitude': '48.45', 'Source': 'source: DIN 4108-6:2003. Klimaregion 11', 'Location': 'Freudenstadt', 'Longitude': '8.41', 'Region': '', 'Country': 'DE'},
    {'Dataset': u'DE0028a-N\xc3\xbcrnberg', 'Comments': 'Bayern', 'Code': '0028a', 'Latitude': '49.5', 'Source': 'source: DIN 4108-6:2003. Klimaregion 11', 'Location': u'N\xc3\xbcrnberg', 'Longitude': '11.06', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0029a-Stuttgart', 'Comments': u'Baden-W\xc3\xbcrttemberg', 'Code': '0029a', 'Latitude': '48.77', 'Source': 'source: DIN 4108-6:2003. Klimaregion 11', 'Location': 'Stuttgart', 'Longitude': '9.18', 'Region': '', 'Country': 'DE'},
    {'Dataset': u'DE0030a-W\xc3\xbcrzburg', 'Comments': 'Bayern', 'Code': '0030a', 'Latitude': '49.77', 'Source': 'source: DIN 4108-6:2003. Klimaregion 11', 'Location': u'W\xc3\xbcrzburg', 'Longitude': '9.96', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0031a-Frankfurt am Main', 'Comments': 'Hessen', 'Code': '0031a', 'Latitude': '50.15', 'Source': 'source: DIN 4108-6:2003. Klimaregion 12', 'Location': 'Frankfurt am Main', 'Longitude': '8.68', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0032a-Mannheim', 'Comments': u'Baden-W\xc3\xbcrttemberg', 'Code': '0032a', 'Latitude': '49.51', 'Source': 'source: DIN 4108-6:2003. Klimaregion 12', 'Location': 'Mannheim', 'Longitude': '8.56', 'Region': '', 'Country': 'DE'},
    {'Dataset': u'DE0033a-Saarbr\xc3\xbccken', 'Comments': 'Saarland', 'Code': '0033a', 'Latitude': '49.21', 'Source': 'source: DIN 4108-6:2003. Klimaregion 12', 'Location': u'Saarbr\xc3\xbccken', 'Longitude': '7.11', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0034a-Freiburg', 'Comments': u'Baden-W\xc3\xbcrttemberg', 'Code': '0034a', 'Latitude': '48.02', 'Source': 'source: DIN 4108-6:2003. Klimaregion 13', 'Location': 'Freiburg', 'Longitude': '7.84', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0035a-Konstanz', 'Comments': u'Baden-W\xc3\xbcrttemberg', 'Code': '0035a', 'Latitude': '47.68', 'Source': 'source: DIN 4108-6:2003. Klimaregion 13', 'Location': 'Konstanz', 'Longitude': '9.19', 'Region': '', 'Country': 'DE'},
    {'Dataset': u'DE0036a-M\xc3\xbcnchen', 'Comments': 'Bayern', 'Code': '0036a', 'Latitude': '48.16', 'Source': 'source: DIN 4108-6:2003. Klimaregion 14', 'Location': u'M\xc3\xbcnchen', 'Longitude': '11.54', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0037a-Passau', 'Comments': 'Bayern', 'Code': '0037a', 'Latitude': '48.58', 'Source': 'source: DIN 4108-6:2003. Klimaregion 14', 'Location': 'Passau', 'Longitude': '13.42', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0038a-Garmisch-Partenkirchen', 'Comments': 'Bayern', 'Code': '0038a', 'Latitude': '47.48', 'Source': 'source: DIN 4108-6:2003. Klimaregion 15', 'Location': 'Garmisch-Partenkirchen', 'Longitude': '11.06', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE0039a-Oberstdorf', 'Comments': 'Bayern', 'Code': '0039a', 'Latitude': '47.4', 'Source': 'source: DIN 4108-6:2003. Klimaregion 15', 'Location': 'Oberstdorf', 'Longitude': '10.28', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE-9999-PHPP-Standard', 'Comments': '', 'Code': '-9999', 'Latitude': '51.301', 'Source': 'Representative of typical climate conditions in Central Europe. This dataset can be used for an assessment independent of the location. ', 'Location': 'PHPP-Standard', 'Longitude': '9.44', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE------Referenzklima (DIN 4108-6:2003)', 'Comments': u'Im PHPP auff\xc3\xbchren oder nicht (weil veraltet)?', 'Code': '', 'Latitude': '50', 'Source': 'Referenzklima Deutschland aus DIN 4108-6. ', 'Location': 'Referenzklima (DIN 4108-6:2003)', 'Longitude': '10', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DE------Referenzklima - EnEV 2014', 'Comments': '', 'Code': '', 'Latitude': '52.38', 'Source': u'Referenzklima f\xc3\xbcr Deutschland nach DIN V 18599-10:2011-12 (TRY Region 4, Potsdam)', 'Location': 'Referenzklima - EnEV 2014', 'Longitude': '13.07', 'Region': '', 'Country': 'DE'},
    {'Dataset': 'DK0001a-Kopenhagen', 'Comments': '', 'Code': '0001a', 'Latitude': '55.72', 'Source': '', 'Location': 'Kopenhagen', 'Longitude': '12.57', 'Region': '', 'Country': 'DK'},
    {'Dataset': 'EE0001a-Toravere', 'Comments': '2011 PHI. Vergleich mit EOSWEB Daten. ', 'Code': '0001a', 'Latitude': '58.45', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Toravere', 'Longitude': '26.783', 'Region': '', 'Country': 'EE'},
    {'Dataset': 'ES0001b-Madrid', 'Comments': '2015 PHI. (LastermittlungPHI-141119)', 'Code': '0001b', 'Latitude': '40.41', 'Source': 'Source: CTE. ', 'Location': 'Madrid', 'Longitude': '-3.68', 'Region': 'D3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0002c-Barcelona', 'Comments': '2016 PHI', 'Code': '0002c', 'Latitude': '41.38', 'Source': 'Temp derived from on 1981-2010 data. Other from CTE & Meteonorm. ', 'Location': 'Barcelona', 'Longitude': '2.13', 'Region': 'B2 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0003b-Sevilla', 'Comments': '2015 PHI', 'Code': '0003b', 'Latitude': '37.42', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Sevilla', 'Longitude': '-5.88', 'Region': 'B4 (CTE)', 'Country': 'ES'},
    {'Dataset': u'ES0004a-L\xc3\xa9rida', 'Comments': '2011 PHI', 'Code': '0004a', 'Latitude': '41.63', 'Source': '', 'Location': u'L\xc3\xa9rida', 'Longitude': '0.6', 'Region': 'D3 (CTE)', 'Country': 'ES'},
    {'Dataset': u'ES0005a-M\xc3\xa1laga', 'Comments': u'2011 PHI.   EOSWEB Heiz- & K\xc3\xbchllastdaten.', 'Code': '0005a', 'Latitude': '36.67', 'Source': 'Source: CTE, Meteonorm V6.', 'Location': u'M\xc3\xa1laga', 'Longitude': '-4.48', 'Region': 'A3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0006c-Bilbao', 'Comments': '2015 PHI', 'Code': '0006c', 'Latitude': '43.3', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Bilbao', 'Longitude': '-2.91', 'Region': 'C1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0007b-Santiago de Compostela', 'Comments': '2016 PHI', 'Code': '0007b', 'Latitude': '42.89', 'Source': 'Temp = 1981-2010; Other derived from Meteonorm. ', 'Location': 'Santiago de Compostela', 'Longitude': '-8.41', 'Region': 'C1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0008b-Albacete', 'Comments': '2015 PHI', 'Code': '0008b', 'Latitude': '38.95', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Albacete', 'Longitude': '-1.86', 'Region': 'D3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0009b-Alicante', 'Comments': '2015 PHI', 'Code': '0009b', 'Latitude': '38.37', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Alicante', 'Longitude': '-0.49', 'Region': 'B4 (CTE)', 'Country': 'ES'},
    {'Dataset': u'ES0010b-Almer\xc3\xada', 'Comments': '2015 PHI', 'Code': '0010b', 'Latitude': '36.85', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': u'Almer\xc3\xada', 'Longitude': '-2.36', 'Region': 'A4 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0011b-Badajoz', 'Comments': '2015 PHI', 'Code': '0011b', 'Latitude': '38.88', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Badajoz', 'Longitude': '-6.81', 'Region': 'C4 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0013b-Burgos', 'Comments': '2015 PHI', 'Code': '0013b', 'Latitude': '42.36', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Burgos', 'Longitude': '-3.62', 'Region': 'E1 (CTE)', 'Country': 'ES'},
    {'Dataset': u'ES0014b-C\xc3\xa1diz', 'Comments': '2015 PHI', 'Code': '0014b', 'Latitude': '36.5', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': u'C\xc3\xa1diz', 'Longitude': '-6.26', 'Region': 'A3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0015b-Granada', 'Comments': '2015 PHI', 'Code': '0015b', 'Latitude': '37.14', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Granada', 'Longitude': '-3.63', 'Region': 'C3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0016b-Las Palmas de Gran Canaria', 'Comments': '2015 PHI', 'Code': '0016b', 'Latitude': '27.92', 'Source': 'Temp = 1981-2010; Other monthly = Satellite & CTE; Load data derived by PHI. ', 'Location': 'Las Palmas de Gran Canaria', 'Longitude': '-15.39', 'Region': 'A3 (CTE)', 'Country': 'ES'},
    {'Dataset': u'ES0017b-Le\xc3\xb3n', 'Comments': '2015 PHI', 'Code': '0017b', 'Latitude': '42.59', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': u'Le\xc3\xb3n', 'Longitude': '-5.65', 'Region': 'E1 (CTE)', 'Country': 'ES'},
    {'Dataset': u'ES0018b-Logro\xc3\xb1o', 'Comments': '2014 PHI', 'Code': '0018b', 'Latitude': '42.45', 'Source': 'Source: CTE. Radiation = satellite data.', 'Location': u'Logro\xc3\xb1o', 'Longitude': '-2.33', 'Region': 'D2 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0019b-Murcia', 'Comments': '2015 PHI', 'Code': '0019b', 'Latitude': '38', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Murcia', 'Longitude': '-1.17', 'Region': 'B3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0020b-Ourense', 'Comments': '2015 PHI', 'Code': '0020b', 'Latitude': '42.33', 'Source': 'Temp = 1981-2010; Other monthly = satellite data; Load data derived by PHI. ', 'Location': 'Ourense', 'Longitude': '-7.86', 'Region': 'C2 (CTE)', 'Country': 'ES'},
    {'Dataset': u'ES0021b-Ovi\xc3\xa9do', 'Comments': '2015 PHI', 'Code': '0021b', 'Latitude': '43.35', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': u'Ovi\xc3\xa9do', 'Longitude': '-5.87', 'Region': 'C1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0022b-Palma de Mallorca', 'Comments': '2015 PHI', 'Code': '0022b', 'Latitude': '39.55', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Palma de Mallorca', 'Longitude': '2.63', 'Region': 'B3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0023b-Pamplona', 'Comments': '2015 PHI', 'Code': '0023b', 'Latitude': '42.78', 'Source': 'Source: CTE. ', 'Location': 'Pamplona', 'Longitude': '-1.65', 'Region': 'D1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0024b-Pontevedra', 'Comments': '2015 PHI', 'Code': '0024b', 'Latitude': '42.44', 'Source': 'Temp = 1981-2010; Monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Pontevedra', 'Longitude': '-8.62', 'Region': 'C1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0025b-Salamanca', 'Comments': '2015 PHI', 'Code': '0025b', 'Latitude': '40.96', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Salamanca', 'Longitude': '-5.5', 'Region': 'D2 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0026b-Toledo', 'Comments': '2015 PHI', 'Code': '0026b', 'Latitude': '39.88', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Toledo', 'Longitude': '-4.05', 'Region': 'C4 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0027b-Valencia', 'Comments': '2015 PHI', 'Code': '0027b', 'Latitude': '39.49', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Valencia', 'Longitude': '-0.47', 'Region': 'B3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0028b-Valladolid', 'Comments': '2015 PHI', 'Code': '0028b', 'Latitude': '41.64', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Valladolid', 'Longitude': '-4.75', 'Region': 'D2 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0029b-Vitoria-Gasteiz', 'Comments': '2015 PHI', 'Code': '0029b', 'Latitude': '42.88', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Vitoria-Gasteiz', 'Longitude': '-2.74', 'Region': 'D1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0030b-Zaragoza', 'Comments': '2015 PHI', 'Code': '0030b', 'Latitude': '41.66', 'Source': 'Temp: AEMET 1981-2010 climate normals. Based on CTE data', 'Location': 'Zaragoza', 'Longitude': '-1', 'Region': 'D3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0031a-Gerona', 'Comments': '2014 PHI', 'Code': '0031a', 'Latitude': '41.9', 'Source': 'Source: CTE. ', 'Location': 'Gerona', 'Longitude': '1.25', 'Region': 'C2 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0032a-Santander', 'Comments': '2015 PHI', 'Code': '0032a', 'Latitude': '43.46', 'Source': 'Source: CTE. ', 'Location': 'Santander', 'Longitude': '-3.82', 'Region': 'C1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0033a-Segovia', 'Comments': '2015 PHI', 'Code': '0033a', 'Latitude': '40.95', 'Source': 'Source: CTE. ', 'Location': 'Segovia', 'Longitude': '-4.13', 'Region': 'D2 (CTE)', 'Country': 'ES'},
    {'Dataset': u'ES0034a-\xc3\x81vilaz', 'Comments': '2015 PHI', 'Code': '0034a', 'Latitude': '40.66', 'Source': 'Source: CTE. ', 'Location': u'\xc3\x81vilaz', 'Longitude': '-4.7', 'Region': 'E1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0035a-Santa Cruz', 'Comments': '2015 PHI', 'Code': '0035a', 'Latitude': '28.46', 'Source': 'Source: CTE. ', 'Location': 'Santa Cruz', 'Longitude': '-16.25', 'Region': 'A3 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0036a-Huesca', 'Comments': '2015 PHI', 'Code': '0036a', 'Latitude': '42.08', 'Source': 'Temp: AEMET 1981-2010 climate normals. Based on CTE data.', 'Location': 'Huesca', 'Longitude': '0.33', 'Region': 'D2 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0037a-A Coruna', 'Comments': '2016 PHI', 'Code': '0037a', 'Latitude': '43.37', 'Source': 'Temp = 1981-2010; Other derived from Meteonorm. ', 'Location': 'A Coruna', 'Longitude': '-8.42', 'Region': 'C1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'ES0038a-Lugo', 'Comments': '2016 PHI', 'Code': '0038a', 'Latitude': '43.11', 'Source': 'Temp = 1981-2010; Other derived from Meteonorm. ', 'Location': 'Lugo', 'Longitude': '-7.46', 'Region': 'D1 (CTE)', 'Country': 'ES'},
    {'Dataset': u'ES0039a-Ir\xc3\xban', 'Comments': '2016 PHI', 'Code': '0039a', 'Latitude': '43.35', 'Source': 'Derived from Meteonorm & 1981-2010 Climate Normals. ', 'Location': u'Ir\xc3\xban', 'Longitude': '-1.8', 'Region': 'C1 (CTE)', 'Country': 'ES'},
    {'Dataset': 'FI0001a-Helsinki', 'Comments': 'IWEC? ', 'Code': '0001a', 'Latitude': '60.22', 'Source': '', 'Location': 'Helsinki', 'Longitude': '25', 'Region': '', 'Country': 'FI'},
    {'Dataset': 'FI0002a-Tampere', 'Comments': '', 'Code': '0002a', 'Latitude': '61.5', 'Source': '', 'Location': 'Tampere', 'Longitude': '23.75', 'Region': '', 'Country': 'FI'},
    {'Dataset': 'FR0001a-Paris', 'Comments': '', 'Code': '0001a', 'Latitude': '48.87', 'Source': '', 'Location': 'Paris', 'Longitude': '2.33', 'Region': u'\xc3\x8ele-de-France', 'Country': 'FR'},
    {'Dataset': 'FR0002a-Nantes', 'Comments': '', 'Code': '0002a', 'Latitude': '47.23', 'Source': '', 'Location': 'Nantes', 'Longitude': '-1.58', 'Region': 'Pays de la Loire', 'Country': 'FR'},
    {'Dataset': 'FR0003a-Dijon', 'Comments': '', 'Code': '0003a', 'Latitude': '47.33', 'Source': '', 'Location': 'Dijon', 'Longitude': '5.03', 'Region': u'Bourgogne-Franche-Comt\xc3\xa9', 'Country': 'FR'},
    {'Dataset': 'FR0004a-Lyon', 'Comments': '', 'Code': '0004a', 'Latitude': '45.77', 'Source': '', 'Location': 'Lyon', 'Longitude': '4.83', 'Region': u'Auvergne-Rh\xc3\xb4ne-Alpes', 'Country': 'FR'},
    {'Dataset': 'FR0005b-Bordeaux', 'Comments': '2014 PHI. PassREg. Compared with EOSWEB, IWEC, WWR data. ', 'Code': '0005b', 'Latitude': '44.83', 'Source': 'Source: Meteonorm V7. ', 'Location': 'Bordeaux', 'Longitude': '-0.57', 'Region': 'Aquitaine-Limousin-Pitou-Charentes', 'Country': 'FR'},
    {'Dataset': 'FR0006a-Marseille', 'Comments': '', 'Code': '0006a', 'Latitude': '43.3', 'Source': '', 'Location': 'Marseille', 'Longitude': '5.37', 'Region': u'Provence-Alpes-C\xc3\xb4te d\xe2\x80\x99Azur', 'Country': 'FR'},
    {'Dataset': 'FR0007b-Brest', 'Comments': '', 'Code': '0007b', 'Latitude': '48.45', 'Source': '', 'Location': 'Brest', 'Longitude': '-4.42', 'Region': 'Bretagne', 'Country': 'FR'},
    {'Dataset': 'FR0008a-Clermont-Ferrand', 'Comments': '', 'Code': '0008a', 'Latitude': '45.78', 'Source': '', 'Location': 'Clermont-Ferrand', 'Longitude': '3.17', 'Region': u'Auvergne-Rh\xc3\xb4ne-Alpes', 'Country': 'FR'},
    {'Dataset': 'FR0009a-Montpellier', 'Comments': '', 'Code': '0009a', 'Latitude': '43.58', 'Source': '', 'Location': 'Montpellier', 'Longitude': '3.97', 'Region': u'Languedoc-Roussillon-Midi-Pyr\xc3\xa9n\xc3\xa9es', 'Country': 'FR'},
    {'Dataset': 'FR0010a-Nancy', 'Comments': '', 'Code': '0010a', 'Latitude': '48.68', 'Source': '', 'Location': 'Nancy', 'Longitude': '6.22', 'Region': 'Alsace-Champagne-Ardenne-Lorraine', 'Country': 'FR'},
    {'Dataset': 'FR0011a-Nice', 'Comments': '', 'Code': '0011a', 'Latitude': '43.65', 'Source': '', 'Location': 'Nice', 'Longitude': '7.2', 'Region': u'Provence-Alpes-C\xc3\xb4te d\xe2\x80\x99Azur', 'Country': 'FR'},
    {'Dataset': 'FR0012a-Strasbourg', 'Comments': '', 'Code': '0012a', 'Latitude': '48.55', 'Source': '', 'Location': 'Strasbourg', 'Longitude': '7.63', 'Region': 'Alsace-Champagne-Ardenne-Lorraine', 'Country': 'FR'},
    {'Dataset': 'FR0013a-Rennes', 'Comments': '', 'Code': '0013a', 'Latitude': '48.12', 'Source': '', 'Location': 'Rennes', 'Longitude': '-1.68', 'Region': 'Bretagne', 'Country': 'FR'},
    {'Dataset': u'FR0014a-M\xc3\xa2con', 'Comments': '', 'Code': '0014a', 'Latitude': '46.3', 'Source': '', 'Location': u'M\xc3\xa2con', 'Longitude': '4.83', 'Region': u'Bourgogne-Franche-Comt\xc3\xa9', 'Country': 'FR'},
    {'Dataset': 'FR0015a-La Rochelle', 'Comments': '', 'Code': '0015a', 'Latitude': '46.17', 'Source': '', 'Location': 'La Rochelle', 'Longitude': '-1.15', 'Region': 'Aquitaine-Limousin-Pitou-Charentes', 'Country': 'FR'},
    {'Dataset': 'FR0016a-Carpentras', 'Comments': '', 'Code': '0016a', 'Latitude': '44.05', 'Source': '', 'Location': 'Carpentras', 'Longitude': '5.05', 'Region': u'Provence-Alpes-C\xc3\xb4te d\xe2\x80\x99Azur', 'Country': 'FR'},
    {'Dataset': 'FR0017a-Agen', 'Comments': '', 'Code': '0017a', 'Latitude': '44.2', 'Source': '', 'Location': 'Agen', 'Longitude': '0.62', 'Region': 'Aquitaine-Limousin-Pitou-Charentes', 'Country': 'FR'},
    {'Dataset': 'FR0018a-Reims', 'Comments': '2015 PHI', 'Code': '0018a', 'Latitude': '49.3', 'Source': 'Temp = 1981-2010; Other derived from Meteonorm.', 'Location': 'Reims', 'Longitude': '4.03', 'Region': 'Alsace-Champagne-Ardenne-Lorraine', 'Country': 'FR'},
    {'Dataset': 'FR0019a-Lille', 'Comments': '2015 PHI', 'Code': '0019a', 'Latitude': '50.56', 'Source': 'Temp = 1981-2010; Other derived from Meteonorm V7. ', 'Location': 'Lille', 'Longitude': '3.09', 'Region': 'Nord-Pas-de-Calais-Picardie', 'Country': 'FR'},
    {'Dataset': 'FR0020a-Abbeville', 'Comments': '2015 PHI', 'Code': '0020a', 'Latitude': '50.14', 'Source': 'Temp = 1981-2010; Other derived from Meteonorm V7. ', 'Location': 'Abbeville', 'Longitude': '1.83', 'Region': 'Nord-Pas-de-Calais-Picardie', 'Country': 'FR'},
    {'Dataset': 'GB0001a-London (Central)', 'Comments': '2011 PHI & BRE.', 'Code': '0001a', 'Latitude': '51.517', 'Source': 'Source: Meteonorm V6. ', 'Location': 'London (Central)', 'Longitude': '-0.111', 'Region': 'Zone 01 - London ', 'Country': 'GB'},
    {'Dataset': 'GB0002a-Silsoe', 'Comments': '2011 PHI & BRE.', 'Code': '0002a', 'Latitude': '52.017', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Silsoe', 'Longitude': '-0.417', 'Region': 'Zone 02 - Thames Valley', 'Country': 'GB'},
    {'Dataset': 'GB0003a-London Gatwick', 'Comments': '2011 PHI & BRE.', 'Code': '0003a', 'Latitude': '51.15', 'Source': 'Source: Meteonorm V6. ', 'Location': 'London Gatwick', 'Longitude': '-0.183', 'Region': 'Zone 03 - South East England', 'Country': 'GB'},
    {'Dataset': 'GB0004a-Efford', 'Comments': '2011 PHI & BRE.', 'Code': '0004a', 'Latitude': '50.733', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Efford', 'Longitude': '-1.567', 'Region': 'Zone 04 - South England', 'Country': 'GB'},
    {'Dataset': 'GB0005a-Exeter', 'Comments': '2011 PHI & BRE.', 'Code': '0005a', 'Latitude': '50.73', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Exeter', 'Longitude': '-3.41', 'Region': 'Zone 05 - South West', 'Country': 'GB'},
    {'Dataset': 'GB0006a-Lyneham', 'Comments': '2011 PHI & BRE.', 'Code': '0006a', 'Latitude': '51.5', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Lyneham', 'Longitude': '-1.983', 'Region': 'Zone 06 - Severn', 'Country': 'GB'},
    {'Dataset': 'GB0007a-Sutton Bonnington', 'Comments': '2011 PHI & BRE.', 'Code': '0007a', 'Latitude': '52.833', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Sutton Bonnington', 'Longitude': '-1.25', 'Region': 'Zone 07 - Midlands', 'Country': 'GB'},
    {'Dataset': 'GB0008a-Fairfield', 'Comments': '2011 PHI & BRE.', 'Code': '0008a', 'Latitude': '53.8', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Fairfield', 'Longitude': '-2.883', 'Region': 'Zone 08 - West Pennines', 'Country': 'GB'},
    {'Dataset': 'GB0009a-Carlise', 'Comments': '2011 PHI & BRE.', 'Code': '0009a', 'Latitude': '54.88', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Carlise', 'Longitude': '-2.93', 'Region': 'Zone 09 - NW England / SW Scotland', 'Country': 'GB'},
    {'Dataset': 'GB0010a-Eskdalemuir', 'Comments': '2011 PHI & BRE.', 'Code': '0010a', 'Latitude': '55.317', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Eskdalemuir', 'Longitude': '-3.2', 'Region': 'Zone 10 - Borders', 'Country': 'GB'},
    {'Dataset': 'GB0011a-Leeming', 'Comments': '2011 PHI & BRE.', 'Code': '0011a', 'Latitude': '54.3', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Leeming', 'Longitude': '-1.533', 'Region': 'Zone 11 - North East', 'Country': 'GB'},
    {'Dataset': 'GB0012a-Waddington', 'Comments': '2011 PHI & BRE.', 'Code': '0012a', 'Latitude': '53.167', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Waddington', 'Longitude': '-0.517', 'Region': 'Zone 12 - East Pennines', 'Country': 'GB'},
    {'Dataset': 'GB0013a-Hemsby', 'Comments': '2011 PHI & BRE.', 'Code': '0013a', 'Latitude': '52.683', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Hemsby', 'Longitude': '1.683', 'Region': 'Zone 13 - East Anglia', 'Country': 'GB'},
    {'Dataset': 'GB0014a-Sennybridge', 'Comments': '2011 PHI & BRE.', 'Code': '0014a', 'Latitude': '52.06', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Sennybridge', 'Longitude': '-3.61', 'Region': 'Zone 14 - Wales', 'Country': 'GB'},
    {'Dataset': 'GB0015b-Glasgow Airport', 'Comments': 'PHI & BRE. Corrected 2015', 'Code': '0015b', 'Latitude': '55.86', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Glasgow Airport', 'Longitude': '-4.43', 'Region': 'Zone 15 - West Scotland', 'Country': 'GB'},
    {'Dataset': 'GB0016a-Dundee', 'Comments': '2011 PHI & BRE.', 'Code': '0016a', 'Latitude': '56.45', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Dundee', 'Longitude': '-3.067', 'Region': 'Zone 16 - East Scotland', 'Country': 'GB'},
    {'Dataset': 'GB0017a-Aberdeen', 'Comments': '2011 PHI & BRE.', 'Code': '0017a', 'Latitude': '57.167', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Aberdeen', 'Longitude': '-2.083', 'Region': 'Zone 17 - North East Scotland', 'Country': 'GB'},
    {'Dataset': 'GB0018a-Aviemore', 'Comments': '2011 PHI & BRE.', 'Code': '0018a', 'Latitude': '57.2', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Aviemore', 'Longitude': '-3.83', 'Region': 'Zone 18 - Highlands', 'Country': 'GB'},
    {'Dataset': 'GB0019a-Stornoway', 'Comments': '2011 PHI & BRE.', 'Code': '0019a', 'Latitude': '58.217', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Stornoway', 'Longitude': '-6.317', 'Region': 'Zone 19 - Western Isles', 'Country': 'GB'},
    {'Dataset': 'GB0020a-Kirkwall Airport', 'Comments': '2011 PHI & BRE.', 'Code': '0020a', 'Latitude': '58.95', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Kirkwall Airport', 'Longitude': '-2.9', 'Region': 'Zone 20 - Orkney', 'Country': 'GB'},
    {'Dataset': 'GB0021a-Lerwick', 'Comments': '2011 PHI & BRE.', 'Code': '0021a', 'Latitude': '60.133', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Lerwick', 'Longitude': '-1.183', 'Region': 'Zone 21 - Shetland', 'Country': 'GB'},
    {'Dataset': 'GB0022a-Belfast-Aldergrove', 'Comments': '2011 PHI & BRE.', 'Code': '0022a', 'Latitude': '54.65', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Belfast-Aldergrove', 'Longitude': '-6.217', 'Region': 'Zone 22 - Northern Ireland', 'Country': 'GB'},
    {'Dataset': 'GR0001a-Volos', 'Comments': '2011 PHI. Vergleich mit IWEC Daten', 'Code': '0001a', 'Latitude': '39.34', 'Source': 'Source: Meteonorm V6.', 'Location': 'Volos', 'Longitude': '23.01', 'Region': '', 'Country': 'GR'},
    {'Dataset': 'GR0002b-Athen', 'Comments': '2015 PHI', 'Code': '0002b', 'Latitude': '37.9', 'Source': 'Source: Meteonorm V7 (Hellenkion, new period). Load data by PHI. ', 'Location': 'Athen', 'Longitude': '23.73', 'Region': '', 'Country': 'GR'},
    {'Dataset': 'HR0001b-Zagreb', 'Comments': '2014 PHI. PassREg. ', 'Code': '0001b', 'Latitude': '45.82', 'Source': 'Source: DHMZ Climate Atlas & Meteonorm V7. ', 'Location': 'Zagreb', 'Longitude': '16.37', 'Region': '', 'Country': 'HR'},
    {'Dataset': 'HU0001a-Budapest', 'Comments': '', 'Code': '0001a', 'Latitude': '47.5', 'Source': '', 'Location': 'Budapest', 'Longitude': '19.05', 'Region': '', 'Country': 'HU'},
    {'Dataset': 'ID0001a-Jakarta', 'Comments': '2015 PHI. ', 'Code': '0001a', 'Latitude': '-5.63', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see handbook)! Source: BerkeleyEarth database (1981-2010 raw data) & satellite data.', 'Location': 'Jakarta', 'Longitude': '106.55', 'Region': '', 'Country': 'ID'},
    {'Dataset': 'IE0001a-Dublin', 'Comments': '2012 PHI. ', 'Code': '0001a', 'Latitude': '53.33', 'Source': 'Source: Meteonorm V6 & satellite data. ', 'Location': 'Dublin', 'Longitude': '-6.25', 'Region': '', 'Country': 'IE'},
    {'Dataset': 'IE0002a-Birr', 'Comments': '', 'Code': '0002a', 'Latitude': '53.083', 'Source': '', 'Location': 'Birr', 'Longitude': '-7.9', 'Region': '', 'Country': 'IE'},
    {'Dataset': 'IE0003a-Cork', 'Comments': '2014 PHI. ', 'Code': '0003a', 'Latitude': '51.85', 'Source': 'Source: 1981-2010 Climate Normals (met.ie). Other data = MN7 Cork Airport. CL = Passipedia. ', 'Location': 'Cork', 'Longitude': '-8.48', 'Region': '', 'Country': 'IE'},
    {'Dataset': 'IE0004a-Belmullet', 'Comments': '2015 PHI. IWEC', 'Code': '0004a', 'Latitude': '54.23', 'Source': '', 'Location': 'Belmullet', 'Longitude': '-10', 'Region': '', 'Country': 'IE'},
    {'Dataset': 'IS0001a-Reykjavik', 'Comments': '', 'Code': '0001a', 'Latitude': '64.13', 'Source': '', 'Location': 'Reykjavik', 'Longitude': '-20.07', 'Region': '', 'Country': 'IS'},
    {'Dataset': "IT0001a-L'Aquila", 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0001a', 'Latitude': '42.135', 'Source': '', 'Location': "L'Aquila", 'Longitude': '13.621', 'Region': 'Abruzzo', 'Country': 'IT'},
    {'Dataset': 'IT0002a-Potenza', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0002a', 'Latitude': '40.636', 'Source': '', 'Location': 'Potenza', 'Longitude': '15.813', 'Region': 'Basilicata', 'Country': 'IT'},
    {'Dataset': 'IT0003a-Catanzaro', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0003a', 'Latitude': '38.9', 'Source': '', 'Location': 'Catanzaro', 'Longitude': '16.6', 'Region': 'Calabria', 'Country': 'IT'},
    {'Dataset': 'IT0004a-Napoli', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0004a', 'Latitude': '40.837', 'Source': '', 'Location': 'Napoli', 'Longitude': '14.252', 'Region': 'Campania', 'Country': 'IT'},
    {'Dataset': 'IT0005a-Bologna', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0005a', 'Latitude': '44.533', 'Source': '', 'Location': 'Bologna', 'Longitude': '11.3', 'Region': 'Emilia-Romagna', 'Country': 'IT'},
    {'Dataset': 'IT0006a-Trieste', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0006a', 'Latitude': '45.65', 'Source': '', 'Location': 'Trieste', 'Longitude': '13.78', 'Region': 'Friuli Venezia Giulia', 'Country': 'IT'},
    {'Dataset': 'IT0007a-Roma (Pratica di Mare)', 'Comments': u'2005 PHI. Pr\xc3\xbcfen!!', 'Code': '0007a', 'Latitude': '41.88', 'Source': 'Source: Meteonorm. Load data based on IGDG. ', 'Location': 'Roma (Pratica di Mare)', 'Longitude': '12.5', 'Region': 'Lazio', 'Country': 'IT'},
    {'Dataset': 'IT0008a-Genova ', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0008a', 'Latitude': '44.4', 'Source': '', 'Location': 'Genova ', 'Longitude': '8.93', 'Region': 'Liguria', 'Country': 'IT'},
    {'Dataset': 'IT0009a-Brescia', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0009a', 'Latitude': '45.55', 'Source': '', 'Location': 'Brescia', 'Longitude': '10.22', 'Region': 'Lombardia', 'Country': 'IT'},
    {'Dataset': 'IT0010b-Milano', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0010b', 'Latitude': '45.47', 'Source': '1996-2005', 'Location': 'Milano', 'Longitude': '9.2', 'Region': 'Lombardia', 'Country': 'IT'},
    {'Dataset': 'IT0011b-Ancona', 'Comments': '2014 PHI. PassREg', 'Code': '0011b', 'Latitude': '43.62', 'Source': '', 'Location': 'Ancona', 'Longitude': '13.52', 'Region': 'Marche', 'Country': 'IT'},
    {'Dataset': 'IT0012a-Campobasso', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0012a', 'Latitude': '41.556', 'Source': '', 'Location': 'Campobasso', 'Longitude': '14.659', 'Region': 'Molise', 'Country': 'IT'},
    {'Dataset': 'IT0013b-Torino', 'Comments': '2015 PHI & ZEPHIR', 'Code': '0013b', 'Latitude': '45.19', 'Source': 'Verschiedene sourcen; Vergleich & Bearbeitung durch ZEPHIR Passivhaus Italia & PHI', 'Location': 'Torino', 'Longitude': '7.65', 'Region': 'Piemonte', 'Country': 'IT'},
    {'Dataset': 'IT0014a-Bari', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0014a', 'Latitude': '41.117', 'Source': '', 'Location': 'Bari', 'Longitude': '16.867', 'Region': 'Puglia', 'Country': 'IT'},
    {'Dataset': 'IT0015a-Cagliari', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0015a', 'Latitude': '39.22', 'Source': '', 'Location': 'Cagliari', 'Longitude': '9.13', 'Region': 'Sardegna', 'Country': 'IT'},
    {'Dataset': 'IT0016b-Palermo  (Punta Raisi)', 'Comments': u'2005 PHI. Pr\xc3\xbcfen!!', 'Code': '0016b', 'Latitude': '38.1', 'Source': 'Source: Meteonorm. Load data based on IGDG. ', 'Location': 'Palermo  (Punta Raisi)', 'Longitude': '13.38', 'Region': 'Sicilia', 'Country': 'IT'},
    {'Dataset': 'IT0017a-Firenze', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0017a', 'Latitude': '43.778', 'Source': '', 'Location': 'Firenze', 'Longitude': '11.254', 'Region': 'Toscana', 'Country': 'IT'},
    {'Dataset': 'IT0018a-Pisa', 'Comments': u'2005 PHI. Pr\xc3\xbcfen!!', 'Code': '0018a', 'Latitude': '43.67', 'Source': 'Source: Meteonorm. Load data based on IGDG. ', 'Location': 'Pisa', 'Longitude': '10.38', 'Region': 'Toscana', 'Country': 'IT'},
    {'Dataset': 'IT0019a-Venezia', 'Comments': u'2005 PHI. Pr\xc3\xbcfen!!', 'Code': '0019a', 'Latitude': '45.49', 'Source': 'Source: Meteonorm. Load data based on IGDG. ', 'Location': 'Venezia', 'Longitude': '12.33', 'Region': 'Veneto', 'Country': 'IT'},
    {'Dataset': 'IT0020a-Bolzano', 'Comments': u'2005 PHI. Pr\xc3\xbcfen!!', 'Code': '0020a', 'Latitude': '46.46', 'Source': 'Source: Meteonorm. Load data based on IGDG. ', 'Location': 'Bolzano', 'Longitude': '11.33', 'Region': 'Trentino - Alto Adige', 'Country': 'IT'},
    {'Dataset': 'IT0021a-Trento', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0021a', 'Latitude': '46.069', 'Source': '', 'Location': 'Trento', 'Longitude': '11.12', 'Region': 'Trentino - Alto Adige', 'Country': 'IT'},
    {'Dataset': 'IT0022a-Perugia', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0022a', 'Latitude': '43.98', 'Source': '', 'Location': 'Perugia', 'Longitude': '12.65', 'Region': 'Umbria', 'Country': 'IT'},
    {'Dataset': 'IT0023a-Aosta', 'Comments': u'pr\xc3\xbcfen!!', 'Code': '0023a', 'Latitude': '45.735', 'Source': '', 'Location': 'Aosta', 'Longitude': '7.309', 'Region': "Valle d'Aosta", 'Country': 'IT'},
    {'Dataset': 'IT0024a-Catania', 'Comments': '2014 PHI. PassREg', 'Code': '0024a', 'Latitude': '37.4', 'Source': 'Source: ClimateAtlas 1971-2000 & ground measured data 2009-2013.', 'Location': 'Catania', 'Longitude': '14.91', 'Region': 'Sicilia', 'Country': 'IT'},
    {'Dataset': 'IT0025a-Cervia', 'Comments': '2014 PHI. PassREg. Compared with satellite, UNI, ground measured, ClimateAtlas. ', 'Code': '0025a', 'Latitude': '44.21', 'Source': 'Source: Meteonorm V7 Station Cervia ("new" period). Load data from CTI TRY. ', 'Location': 'Cervia', 'Longitude': '12.3', 'Region': 'Emilia-Romagna', 'Country': 'IT'},
    {'Dataset': 'IT0026a-Verona / Valeggio', 'Comments': '2015 PHI & ZEPHIR', 'Code': '0026a', 'Latitude': '45.38', 'Source': 'Verschiedene sourcen; Vergleich & Bearbeitung durch ZEPHIR Passivhaus Italia & PHI', 'Location': 'Verona / Valeggio', 'Longitude': '10.87', 'Region': 'Veneto', 'Country': 'IT'},
    {'Dataset': 'IT0027a-Bergamo', 'Comments': '2015 PHI & ZEPHIR', 'Code': '0027a', 'Latitude': '45.66', 'Source': 'Verschiedene sourcen; Vergleich & Bearbeitung durch ZEPHIR Passivhaus Italia & PHI', 'Location': 'Bergamo', 'Longitude': '9.66', 'Region': 'Lombardia', 'Country': 'IT'},
    {'Dataset': 'JP0001a-Tokyo', 'Comments': '2009 PHI. Vergleich mit IWEC, EOSWEB, NOAA.', 'Code': '0001a', 'Latitude': '35.683', 'Source': 'Source: Meteonorm. ', 'Location': 'Tokyo', 'Longitude': '139.767', 'Region': '', 'Country': 'JP'},
    {'Dataset': 'JP0002a-Sapporo', 'Comments': '2009 PHI. Vergleich mit IWEC.', 'Code': '0002a', 'Latitude': '43.08', 'Source': 'Source: Meteonorm. ', 'Location': 'Sapporo', 'Longitude': '141.35', 'Region': '', 'Country': 'JP'},
    {'Dataset': 'KP0001a-Hyesan', 'Comments': '2016 PHI', 'Code': '0001a', 'Latitude': '41.4', 'Source': 'Temp = 1981-2010; Other = derived from Meteonorm', 'Location': 'Hyesan', 'Longitude': '128.17', 'Region': '', 'Country': 'KP'},
    {'Dataset': 'KR0001a-Seoul', 'Comments': 'confirmed 2015 PHI', 'Code': '0001a', 'Latitude': '37.5', 'Source': '', 'Location': 'Seoul', 'Longitude': '127', 'Region': '', 'Country': 'KR'},
    {'Dataset': 'KR0002a-Jeonju', 'Comments': '2013 PHI. MN Station: Chonchu (Jeonju) KS, Perez, neuePeriode. KL: EOSWEB. Vergleich mit EOSWEB & Korean Meteorology Organisation (1981-2012).', 'Code': '0002a', 'Latitude': '35.82', 'Source': 'Source: Meteonorm V6.1 & satellite data', 'Location': 'Jeonju', 'Longitude': '127.15', 'Region': '', 'Country': 'KR'},
    {'Dataset': 'KR0003a-Cheongju', 'Comments': '2011 PHI', 'Code': '0003a', 'Latitude': '36.65', 'Source': 'Source: Meteonorm V6.1 & satellite data', 'Location': 'Cheongju', 'Longitude': '127.45', 'Region': '', 'Country': 'KR'},
    {'Dataset': 'KZ0001a-Almaty', 'Comments': '2011 PHI', 'Code': '0001a', 'Latitude': '43.32', 'Source': 'Source: Meteonorm.', 'Location': 'Almaty', 'Longitude': '76.92', 'Region': '', 'Country': 'KZ'},
    {'Dataset': 'LT0001a-Vilnius', 'Comments': '', 'Code': '0001a', 'Latitude': '54.68333333', 'Source': '', 'Location': 'Vilnius', 'Longitude': '25.26666667', 'Region': '', 'Country': 'LT'},
    {'Dataset': 'LT0002a-Panara', 'Comments': '2012 PHI. ', 'Code': '0002a', 'Latitude': '54.09', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Panara', 'Longitude': '24.11', 'Region': '', 'Country': 'LT'},
    {'Dataset': 'LU0001a-Luxembourg', 'Comments': '', 'Code': '0001a', 'Latitude': '49.612', 'Source': '', 'Location': 'Luxembourg', 'Longitude': '6.13', 'Region': '', 'Country': 'LU'},
    {'Dataset': 'LV0001a-Riga', 'Comments': '2014 PHI. PassREg. New period. Compared with WMO, EOSWEB, LBN 003-01.', 'Code': '0001a', 'Latitude': '56.97', 'Source': 'Source: Meteonorm7, Station: Riga. ', 'Location': 'Riga', 'Longitude': '24.05', 'Region': '', 'Country': 'LV'},
    {'Dataset': 'LV0002a-Dougavpils', 'Comments': '2014 PHI. PassREg. Compared with WMO, EOSWEB, LBN 003-01. PHI May 2014', 'Code': '0002a', 'Latitude': '55.87', 'Source': 'Source: Meteonorm7, Station: Dougavpils.', 'Location': 'Dougavpils', 'Longitude': '26.62', 'Region': '', 'Country': 'LV'},
    {'Dataset': 'MX0001b-Puebla, Puebla', 'Comments': '2012 PHI. ', 'Code': '0001b', 'Latitude': '19.05', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Puebla, Puebla', 'Longitude': '-98.2', 'Region': u'6 - Templado subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0002b-Jalisco, Guadalajara', 'Comments': '2012 PHI. ', 'Code': '0002b', 'Latitude': '20.7', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Jalisco, Guadalajara', 'Longitude': '-103.3', 'Region': u'6 - Templado subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0003b-Sonora, Hermosillo', 'Comments': '2012 PHI. ', 'Code': '0003b', 'Latitude': '29.1', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Sonora, Hermosillo', 'Longitude': '-111', 'Region': '3 - Muy seco', 'Country': 'MX'},
    {'Dataset': u'MX0004b-Quintana Roo, Canc\xc3\xban', 'Comments': '2012 PHI. ', 'Code': '0004b', 'Latitude': '21.1', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'Quintana Roo, Canc\xc3\xban', 'Longitude': '-86.8', 'Region': u'2 - C\xc3\xa1lido subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0005a-Aguascalientes, Aguascalientes', 'Comments': '2012 PHI. ', 'Code': '0005a', 'Latitude': '21.88', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Aguascalientes, Aguascalientes', 'Longitude': '-102.3', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': u'MX0006a-Distrito Federal, M\xc3\xa9xico D.F.', 'Comments': '2012 PHI. ', 'Code': '0006a', 'Latitude': '19.47', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'Distrito Federal, M\xc3\xa9xico D.F.', 'Longitude': '-99.08', 'Region': u'6 - Templado subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': u'MX0007a-Nuevo Le\xc3\xb3n, Monterrey', 'Comments': '2012 PHI. ', 'Code': '0007a', 'Latitude': '25.67', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'Nuevo Le\xc3\xb3n, Monterrey', 'Longitude': '-100.31', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': 'MX0008a-Oaxaca, Oaxaca', 'Comments': '2012 PHI. ', 'Code': '0008a', 'Latitude': '17.08', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Oaxaca, Oaxaca', 'Longitude': '-96.71', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': 'MX0009a-Baja California, Tijuana', 'Comments': '2012 PHI. ', 'Code': '0009a', 'Latitude': '32.44', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Baja California, Tijuana', 'Longitude': '-116.91', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': 'MX0010a-Veracruz, Xalapa', 'Comments': '2012 PHI. ', 'Code': '0010a', 'Latitude': '19.54166667', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Veracruz, Xalapa', 'Longitude': '-96.91383333', 'Region': u'1 - C\xc3\xa1lido h\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0011a-Chihuahua, Chihuahua', 'Comments': '2012 PHI. ', 'Code': '0011a', 'Latitude': '28.4', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Chihuahua, Chihuahua', 'Longitude': '-106.12', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': 'MX0012a-Chihuahua, Juarez', 'Comments': '2012 PHI. ', 'Code': '0012a', 'Latitude': '31.63', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Chihuahua, Juarez', 'Longitude': '-106.43', 'Region': '3 - Muy seco', 'Country': 'MX'},
    {'Dataset': u'MX0013a-Quer\xc3\xa9taro, Quer\xc3\xa9taro', 'Comments': '2012 PHI. ', 'Code': '0013a', 'Latitude': '20.57', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'Quer\xc3\xa9taro, Quer\xc3\xa9taro', 'Longitude': '-100.37', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': u'MX0014a-San Luis Potos\xc3\xad, San Luis Potos\xc3\xad', 'Comments': '2012 PHI. ', 'Code': '0014a', 'Latitude': '22.15', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'San Luis Potos\xc3\xad, San Luis Potos\xc3\xad', 'Longitude': '-101', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': u'MX0015a-M\xc3\xa9xico, Toluca', 'Comments': '2012 PHI. ', 'Code': '0015a', 'Latitude': '19.3', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'M\xc3\xa9xico, Toluca', 'Longitude': '-99.7', 'Region': u'6 - Templado subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': u'MX0016a-Guanajuato, Le\xc3\xb3n', 'Comments': '2012 PHI. ', 'Code': '0016a', 'Latitude': '21.1', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'Guanajuato, Le\xc3\xb3n', 'Longitude': '-101.7', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': 'MX0017a-Guerrero, Acapulco', 'Comments': '2012 PHI. ', 'Code': '0017a', 'Latitude': '16.8', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Guerrero, Acapulco', 'Longitude': '-99.9', 'Region': u'2 - C\xc3\xa1lido subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0018a-Campeche, Campeche', 'Comments': '2012 PHI. ', 'Code': '0018a', 'Latitude': '19.85', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Campeche, Campeche', 'Longitude': '-90.55', 'Region': u'2 - C\xc3\xa1lido subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0019a-Tamaulipas, Ciudad Victoria', 'Comments': '2012 PHI. ', 'Code': '0019a', 'Latitude': '23.73', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Tamaulipas, Ciudad Victoria', 'Longitude': '-99.17', 'Region': u'6 - Templado subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': u'MX0020a-Sinaloa, Culiac\xc3\xa1n', 'Comments': '2012 PHI. ', 'Code': '0020a', 'Latitude': '24.79', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'Sinaloa, Culiac\xc3\xa1n', 'Longitude': '-107.4', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': 'MX0021a-Durango, Durango', 'Comments': '2012 PHI. Koordinaten korrigiert', 'Code': '0021a', 'Latitude': '24.09', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Durango, Durango', 'Longitude': '-104.6', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': 'MX0022a-Baja California Sur, La Paz', 'Comments': '2012 PHI. ', 'Code': '0022a', 'Latitude': '24.06', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Baja California Sur, La Paz', 'Longitude': '-110.36', 'Region': '3 - Muy seco', 'Country': 'MX'},
    {'Dataset': 'MX0023a-Tamaulipas, Matamoros', 'Comments': '2012 PHI. ', 'Code': '0023a', 'Latitude': '25.76', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Tamaulipas, Matamoros', 'Longitude': '-97.53', 'Region': u'2 - C\xc3\xa1lido subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0024a-Sinaloa, Mazatlan', 'Comments': '2012 PHI. ', 'Code': '0024a', 'Latitude': '23.22', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Sinaloa, Mazatlan', 'Longitude': '-106.41', 'Region': u'2 - C\xc3\xa1lido subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0025a-Baja California, Mexicali', 'Comments': '2012 PHI. ', 'Code': '0025a', 'Latitude': '32.66', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Baja California, Mexicali', 'Longitude': '-115.47', 'Region': '3 - Muy seco', 'Country': 'MX'},
    {'Dataset': 'MX0026a-Tamaulipas, Nuevo Laredo', 'Comments': '2012 PHI. ', 'Code': '0026a', 'Latitude': '27.55', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Tamaulipas, Nuevo Laredo', 'Longitude': '-99.46', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': 'MX0027a-Coahuila, Saltillo', 'Comments': '2012 PHI. ', 'Code': '0027a', 'Latitude': '25.38', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Coahuila, Saltillo', 'Longitude': '-101.17', 'Region': '4 - Seco y semiseco', 'Country': 'MX'},
    {'Dataset': 'MX0028a-Tamaulipas, Tampico', 'Comments': '2012 PHI. ', 'Code': '0028a', 'Latitude': '22.2', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Tamaulipas, Tampico', 'Longitude': '-97.86', 'Region': u'2 - C\xc3\xa1lido subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': u'MX0029a-Puebla, Teziutl\xc3\xa1n', 'Comments': '2012 PHI. ', 'Code': '0029a', 'Latitude': '19.82', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'Puebla, Teziutl\xc3\xa1n', 'Longitude': '-97.36', 'Region': u'5 - Templado h\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': u'MX0030b-Coahuila, Torre\xc3\xb3n', 'Comments': '2015 PHI. DEEVi 2', 'Code': '0030b', 'Latitude': '25.54', 'Source': 'Temp & humidity = 1981-2010 SMN; Other = Satellite data. ', 'Location': u'Coahuila, Torre\xc3\xb3n', 'Longitude': '-103.47', 'Region': '3 - Muy seco', 'Country': 'MX'},
    {'Dataset': 'MX0031a-Chiapas, Tuxtla', 'Comments': '2012 PHI. ', 'Code': '0031a', 'Latitude': '16.75', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Chiapas, Tuxtla', 'Longitude': '-93.13', 'Region': u'2 - C\xc3\xa1lido subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': u'MX0032a-Michoac\xc3\xa1n, Uruapan', 'Comments': u'2012 PHI. (2015 L\xc3\xa4ngengrad korrigiert)', 'Code': '0032a', 'Latitude': '19.4', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': u'Michoac\xc3\xa1n, Uruapan', 'Longitude': '-102.03', 'Region': u'1 - C\xc3\xa1lido h\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0033a-Veracruz, Veracruz', 'Comments': '2012 PHI. ', 'Code': '0033a', 'Latitude': '19.16', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Veracruz, Veracruz', 'Longitude': '-96.14', 'Region': u'2 - C\xc3\xa1lido subh\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'MX0034a-Tabasco, Villahermosa', 'Comments': '2012 PHI. ', 'Code': '0034a', 'Latitude': '18', 'Source': 'Source: Satellite data / SMN / Meteonorm.', 'Location': 'Tabasco, Villahermosa', 'Longitude': '-92.93', 'Region': u'1 - C\xc3\xa1lido h\xc3\xbamedo', 'Country': 'MX'},
    {'Dataset': 'NL0001b-Amsterdam (Schiphol)', 'Comments': '2015 PHI', 'Code': '0001b', 'Latitude': '52.3', 'Source': 'Source: KNMI (1980-2009), supplemented with MN7 data for radiation.', 'Location': 'Amsterdam (Schiphol)', 'Longitude': '4.77', 'Region': '', 'Country': 'NL'},
    {'Dataset': 'NL0002b-Groningen (Eelde)', 'Comments': '2015 PHI', 'Code': '0002b', 'Latitude': '53.13', 'Source': 'Source: KNMI (1980-2009), supplemented with MN7 data for radiation.', 'Location': 'Groningen (Eelde)', 'Longitude': '6.59', 'Region': '', 'Country': 'NL'},
    {'Dataset': 'NL0003c-De Bilt', 'Comments': '2015 PHI', 'Code': '0003c', 'Latitude': '52.1', 'Source': 'Source: KNMI (1980-2009), supplemented with MN7 data for radiation.', 'Location': 'De Bilt', 'Longitude': '5.18', 'Region': '', 'Country': 'NL'},
    {'Dataset': 'NL0004b-De Kooy', 'Comments': '2015 PHI', 'Code': '0004b', 'Latitude': '52.92', 'Source': 'Source: KNMI (1980-2009), supplemented with MN7 data for radiation.', 'Location': 'De Kooy', 'Longitude': '4.79', 'Region': '', 'Country': 'NL'},
    {'Dataset': 'NL0005b-Vlissingen', 'Comments': '2015 PHI', 'Code': '0005b', 'Latitude': '51.44', 'Source': 'Source: KNMI (1980-2009), supplemented with MN7 data for radiation.', 'Location': 'Vlissingen', 'Longitude': '3.6', 'Region': '', 'Country': 'NL'},
    {'Dataset': 'NL0006b-Twente', 'Comments': '2015 PHI', 'Code': '0006b', 'Latitude': '52.27', 'Source': 'Source: KNMI (1980-2009), supplemented with MN7 data for radiation.', 'Location': 'Twente', 'Longitude': '6.9', 'Region': '', 'Country': 'NL'},
    {'Dataset': 'NL0007a-Eindhoven', 'Comments': '2015 PHI', 'Code': '0007a', 'Latitude': '51.45', 'Source': 'Source: KNMI (1980-2009), supplemented with MN7 data for radiation.', 'Location': 'Eindhoven', 'Longitude': '5.41', 'Region': '', 'Country': 'NL'},
    {'Dataset': 'NL0008a-Leiden (Valkenburg)', 'Comments': '2015 PHI', 'Code': '0008a', 'Latitude': '52.17', 'Source': 'Source: KNMI (1980-2009), supplemented with MN7 data for radiation.', 'Location': 'Leiden (Valkenburg)', 'Longitude': '4.42', 'Region': '', 'Country': 'NL'},
    {'Dataset': 'NL0009a-Maastricht', 'Comments': '2015 PHI', 'Code': '0009a', 'Latitude': '50.91', 'Source': 'Source: KNMI (1980-2009), supplemented with MN7 data for radiation.', 'Location': 'Maastricht', 'Longitude': '5.77', 'Region': '', 'Country': 'NL'},
    {'Dataset': 'NO0001a-Oslo', 'Comments': '', 'Code': '0001a', 'Latitude': '59.93', 'Source': '', 'Location': 'Oslo', 'Longitude': '10.75', 'Region': '', 'Country': 'NO'},
    {'Dataset': 'NO0002a-Bergen', 'Comments': 'PEP project', 'Code': '0002a', 'Latitude': '60.38', 'Source': '', 'Location': 'Bergen', 'Longitude': '5.33', 'Region': '', 'Country': 'NO'},
    {'Dataset': 'NO0003a-Trondheim', 'Comments': '', 'Code': '0003a', 'Latitude': '63.42', 'Source': '', 'Location': 'Trondheim', 'Longitude': '10.4', 'Region': '', 'Country': 'NO'},
    {'Dataset': 'NZ0001a-Auckland', 'Comments': 'PHPP 8 ', 'Code': '0001a', 'Latitude': '-37', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'Auckland', 'Longitude': '174.8', 'Region': 'AK (Auckland) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0002a-Wellington', 'Comments': 'PHPP 8 ', 'Code': '0002a', 'Latitude': '-41.4', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'Wellington', 'Longitude': '174.9', 'Region': 'WN (Wellington) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0003a-Christchurch', 'Comments': 'PHPP 8 ', 'Code': '0003a', 'Latitude': '-43.5', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'Christchurch', 'Longitude': '172.6', 'Region': 'CC (Christchurch) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0004a-Masterton', 'Comments': '2012 PHI. ', 'Code': '0004a', 'Latitude': '-41.02', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'Masterton', 'Longitude': '175.62', 'Region': 'WI (Wairarapa) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0005a-New Plymouth', 'Comments': '2012 PHI. ', 'Code': '0005a', 'Latitude': '-39.01', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'New Plymouth', 'Longitude': '174.18', 'Region': 'NP (New Plymouth) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0006b-Queenstown', 'Comments': '2015 PHI', 'Code': '0006b', 'Latitude': '-45.02', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'Queenstown', 'Longitude': '168.74', 'Region': 'QL (Queenstown-Lakes) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0007a-Tauranga', 'Comments': '2012 PHI. ', 'Code': '0007a', 'Latitude': '-37.67', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'Tauranga', 'Longitude': '176.2', 'Region': 'BP (Bay of Plenty) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0008a-Turangi', 'Comments': '2012 PHI. ', 'Code': '0008a', 'Latitude': '-38.99', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'Turangi', 'Longitude': '175.81', 'Region': 'TP (Taupo) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0009a-Paraparaumu', 'Comments': '2014 PHI. Compared with MN7, IWEC, satellite data.', 'Code': '0009a', 'Latitude': '-40.91', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'Paraparaumu', 'Longitude': '174.98', 'Region': 'MW (Manawatu) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0010a-Hamilton / Ruakura', 'Comments': '2014 PHI. Compard with Meteonorm & satellite data', 'Code': '0010a', 'Latitude': '-37.78', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!)   Source: NIWA, TMY2. ', 'Location': 'Hamilton / Ruakura', 'Longitude': '175.31', 'Region': 'HN (Hamilton) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0011a-Napier', 'Comments': '2015 PHI', 'Code': '0011a', 'Latitude': '-39.45', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!). Source: Meteonorm7. Load data derived from NIWA data. ', 'Location': 'Napier', 'Longitude': '176.85', 'Region': 'EC (East Coast) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0012a-Nelson ', 'Comments': '2015 PHI', 'Code': '0012a', 'Latitude': '-41.3', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!). Source: Meteonorm7. Load data derived from NIWA data. ', 'Location': 'Nelson ', 'Longitude': '173.23', 'Region': 'NM (Nelson Marlborough) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'NZ0013a-Dunedin', 'Comments': '2015 PHI', 'Code': '0013a', 'Latitude': '-45.9', 'Source': 'Location is on the Southern Hemisphere. The original data has been adapted for use in the PHPP (see PHPP manual!). Source: NIWA & Satellite data (radiation)', 'Location': 'Dunedin', 'Longitude': '170.51', 'Region': 'DN (Dunedin) NIWA Zone', 'Country': 'NZ'},
    {'Dataset': 'PH0001a-Manila/Naia', 'Comments': '2011 PHI. Vergleich mit EOSWEB, IWEC. ', 'Code': '0001a', 'Latitude': '14.517', 'Source': 'Source: Meteonorm. ', 'Location': 'Manila/Naia', 'Longitude': '121', 'Region': '', 'Country': 'PH'},
    {'Dataset': 'PL0001a-Koszalin/Kolobrzeg', 'Comments': '2005 PHI', 'Code': '0001a', 'Latitude': '54.17', 'Source': '', 'Location': 'Koszalin/Kolobrzeg', 'Longitude': '16.17', 'Region': 'Strefa I', 'Country': 'PL'},
    {'Dataset': 'PL0002a-Poznan/Pila', 'Comments': '2005 PHI', 'Code': '0002a', 'Latitude': '52.42', 'Source': '', 'Location': 'Poznan/Pila', 'Longitude': '16.88', 'Region': 'Strefa II', 'Country': 'PL'},
    {'Dataset': 'PL0003a-Warszawa', 'Comments': '2005 PHI', 'Code': '0003a', 'Latitude': '52.25', 'Source': '', 'Location': 'Warszawa', 'Longitude': '21', 'Region': 'Strefa III', 'Country': 'PL'},
    {'Dataset': 'PL0004a-Bialystok/Mikolajki', 'Comments': '2005 PHI', 'Code': '0004a', 'Latitude': '53.15', 'Source': '', 'Location': 'Bialystok/Mikolajki', 'Longitude': '23.17', 'Region': 'Strefa IV', 'Country': 'PL'},
    {'Dataset': 'PL0005a-Suwalki/Mikolajki', 'Comments': '2005 PHI', 'Code': '0005a', 'Latitude': '54.1', 'Source': '', 'Location': 'Suwalki/Mikolajki', 'Longitude': '22.95', 'Region': 'Strefa V N', 'Country': 'PL'},
    {'Dataset': 'PL0006a-Zakopane', 'Comments': '2005 PHI', 'Code': '0006a', 'Latitude': '49.3', 'Source': '', 'Location': 'Zakopane', 'Longitude': '19.95', 'Region': 'Strefa V S', 'Country': 'PL'},
    {'Dataset': 'PT0001a-Lisboa', 'Comments': '', 'Code': '0001a', 'Latitude': '38.73', 'Source': '', 'Location': 'Lisboa', 'Longitude': '-9.13', 'Region': '', 'Country': 'PT'},
    {'Dataset': 'PT0002a-Porto ', 'Comments': '2011 PHI. Vergleich mit IWEC Daten.', 'Code': '0002a', 'Latitude': '41.133', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Porto ', 'Longitude': '-8.6', 'Region': '', 'Country': 'PT'},
    {'Dataset': 'RO0001a-Satu-Mare', 'Comments': '2011 PHI', 'Code': '0001a', 'Latitude': '47.8', 'Source': 'Source: Meteonorm V6. ', 'Location': 'Satu-Mare', 'Longitude': '22.87', 'Region': '', 'Country': 'RO'},
    {'Dataset': 'RO0002a-Sibiu', 'Comments': '2015 PHI', 'Code': '0002a', 'Latitude': '45.8', 'Source': 'Source: Meteonorm. Compared with INCERC', 'Location': 'Sibiu', 'Longitude': '24.15', 'Region': '', 'Country': 'RO'},
    {'Dataset': 'RO0003a-Cluj', 'Comments': '2015 PHI', 'Code': '0003a', 'Latitude': '46.78', 'Source': 'Temp = INCERC, Other =  Meteonorm. Load data derived by PHI. ', 'Location': 'Cluj', 'Longitude': '23.57', 'Region': '', 'Country': 'RO'},
    {'Dataset': 'RS0001a-Belgrad', 'Comments': '', 'Code': '0001a', 'Latitude': '44.82', 'Source': '', 'Location': 'Belgrad', 'Longitude': '20.47', 'Region': '', 'Country': 'RS'},
    {'Dataset': u'RS0002a-Ni\xc5\xa1', 'Comments': '', 'Code': '0002a', 'Latitude': '43.33', 'Source': '', 'Location': u'Ni\xc5\xa1', 'Longitude': '21.9', 'Region': '', 'Country': 'RS'},
    {'Dataset': u'RS0003a-Pri\xc5\xa1tina', 'Comments': '', 'Code': '0003a', 'Latitude': '42.65', 'Source': '', 'Location': u'Pri\xc5\xa1tina', 'Longitude': '21.15', 'Region': '', 'Country': 'RS'},
    {'Dataset': 'RS0004a-Banja Luka', 'Comments': '', 'Code': '0004a', 'Latitude': '44.78', 'Source': '', 'Location': 'Banja Luka', 'Longitude': '17.22', 'Region': '', 'Country': 'RS'},
    {'Dataset': 'RU0001a-Moskva', 'Comments': '', 'Code': '0001a', 'Latitude': '55.77', 'Source': '', 'Location': 'Moskva', 'Longitude': '37.67', 'Region': '', 'Country': 'RU'},
    {'Dataset': 'RU0002a-Ekaterinburg', 'Comments': '', 'Code': '0002a', 'Latitude': '56.85', 'Source': '', 'Location': 'Ekaterinburg', 'Longitude': '60.6', 'Region': '', 'Country': 'RU'},
    {'Dataset': 'SD0001a-Khartoum', 'Comments': '2015 PHI', 'Code': '0001a', 'Latitude': '15.6', 'Source': 'Temp = 1981-2010; Other = Meteonorm V7. ', 'Location': 'Khartoum', 'Longitude': '32.55', 'Region': '', 'Country': 'SD'},
    {'Dataset': 'SE0001a-Stockholm', 'Comments': '', 'Code': '0001a', 'Latitude': '59.325', 'Source': '', 'Location': 'Stockholm', 'Longitude': '18.07', 'Region': '', 'Country': 'SE'},
    {'Dataset': u'SE0002a-Borl\xc3\xa4nge', 'Comments': '', 'Code': '0002a', 'Latitude': '60.433', 'Source': '', 'Location': u'Borl\xc3\xa4nge', 'Longitude': '15.5', 'Region': '', 'Country': 'SE'},
    {'Dataset': u'SE0003a-G\xc3\xb6teborg', 'Comments': '', 'Code': '0003a', 'Latitude': '57.783', 'Source': '', 'Location': u'G\xc3\xb6teborg', 'Longitude': '11.883', 'Region': '', 'Country': 'SE'},
    {'Dataset': u'SE0004a-J\xc3\xb6nk\xc3\xb6ping', 'Comments': '', 'Code': '0004a', 'Latitude': '57.75', 'Source': '', 'Location': u'J\xc3\xb6nk\xc3\xb6ping', 'Longitude': '14.17', 'Region': '', 'Country': 'SE'},
    {'Dataset': 'SE0005a-Kalmar', 'Comments': '', 'Code': '0005a', 'Latitude': '56.73', 'Source': '', 'Location': 'Kalmar', 'Longitude': '16.3', 'Region': '', 'Country': 'SE'},
    {'Dataset': 'SE0006a-Karlstad', 'Comments': '', 'Code': '0006a', 'Latitude': '59.367', 'Source': '', 'Location': 'Karlstad', 'Longitude': '13.467', 'Region': '', 'Country': 'SE'},
    {'Dataset': 'SE0007a-Kiruna', 'Comments': '', 'Code': '0007a', 'Latitude': '67.81', 'Source': '', 'Location': 'Kiruna', 'Longitude': '20.33', 'Region': '', 'Country': 'SE'},
    {'Dataset': u'SE0008a-Lule\xc3\xa5', 'Comments': '', 'Code': '0008a', 'Latitude': '65.55', 'Source': '', 'Location': u'Lule\xc3\xa5', 'Longitude': '22.133', 'Region': '', 'Country': 'SE'},
    {'Dataset': 'SE0009a-Lund', 'Comments': '', 'Code': '0009a', 'Latitude': '55.717', 'Source': '', 'Location': 'Lund', 'Longitude': '13.217', 'Region': '', 'Country': 'SE'},
    {'Dataset': u'SE0010a-\xc3\x96stersund', 'Comments': '', 'Code': '0010a', 'Latitude': '63.183', 'Source': '', 'Location': u'\xc3\x96stersund', 'Longitude': '14.5', 'Region': '', 'Country': 'SE'},
    {'Dataset': 'SE0011a-Sundsvall', 'Comments': '', 'Code': '0011a', 'Latitude': '62.533', 'Source': '', 'Location': 'Sundsvall', 'Longitude': '17.45', 'Region': '', 'Country': 'SE'},
    {'Dataset': u'SE0012a-Ume\xc3\xa5', 'Comments': '', 'Code': '0012a', 'Latitude': '63.817', 'Source': '', 'Location': u'Ume\xc3\xa5', 'Longitude': '20.25', 'Region': '', 'Country': 'SE'},
    {'Dataset': 'SI0001a-Ljubljana', 'Comments': '2015 PHI', 'Code': '0001a', 'Latitude': '46.07', 'Source': 'Temp = 1981-2010; Other derived from Meteonorm. ', 'Location': 'Ljubljana', 'Longitude': '14.52', 'Region': '', 'Country': 'SI'},
    {'Dataset': 'SK0001a-Bratislava', 'Comments': '2007 PHI. Mit Meteonorm bearbeitet. Vergleich mit Meteonorm & NOAA. Lastadaten auf Basis von NOAA Daten. 2007', 'Code': '0001a', 'Latitude': '48.17', 'Source': 'Based on: STN EN ISO 13 790', 'Location': 'Bratislava', 'Longitude': '17.17', 'Region': '', 'Country': 'SK'},
    {'Dataset': 'SK0002a-Hurbanovo', 'Comments': '2007 PHI. Mit Meteonorm bearbeitet. Vergleich mit Meteonorm & NOAA. Lastadaten auf Basis von NOAA Daten. 2007', 'Code': '0002a', 'Latitude': '47.867', 'Source': 'Based on: STN EN ISO 13 790', 'Location': 'Hurbanovo', 'Longitude': '18.2', 'Region': '', 'Country': 'SK'},
    {'Dataset': 'SK0003a-Kamenica nad Cirochou', 'Comments': '2007 PHI. Mit Meteonorm bearbeitet. Vergleich mit Meteonorm & NOAA. Lastadaten auf Basis von NOAA Daten. 2007', 'Code': '0003a', 'Latitude': '48.93', 'Source': 'Based on: STN EN ISO 13 790', 'Location': 'Kamenica nad Cirochou', 'Longitude': '22', 'Region': '', 'Country': 'SK'},
    {'Dataset': u'SK0004a-Ko\xc5\xa1ice', 'Comments': '2007 PHI. Mit Meteonorm bearbeitet. Vergleich mit Meteonorm & NOAA. Lastadaten auf Basis von NOAA Daten. 2007', 'Code': '0004a', 'Latitude': '48.73', 'Source': 'Based on: STN EN ISO 13 790', 'Location': u'Ko\xc5\xa1ice', 'Longitude': '21.25', 'Region': '', 'Country': 'SK'},
    {'Dataset': u'SK0005a-Pie\xc5\xa1\xc5\xa5any', 'Comments': '2007 PHI. Mit Meteonorm bearbeitet. Vergleich mit Meteonorm & NOAA. Lastadaten auf Basis von NOAA Daten. 2007', 'Code': '0005a', 'Latitude': '48.53', 'Source': 'Based on: STN EN ISO 13 790', 'Location': u'Pie\xc5\xa1\xc5\xa5any', 'Longitude': '17.83', 'Region': '', 'Country': 'SK'},
    {'Dataset': 'SK0006a-Poprad', 'Comments': '2007 PHI. Mit Meteonorm bearbeitet. Vergleich mit Meteonorm & NOAA. Lastadaten auf Basis von NOAA Daten. 2007', 'Code': '0006a', 'Latitude': '49.067', 'Source': 'Based on: STN EN ISO 13 790', 'Location': 'Poprad', 'Longitude': '20.25', 'Region': '', 'Country': 'SK'},
    {'Dataset': 'SK0007a-Sliac', 'Comments': '2007 PHI. Mit Meteonorm bearbeitet. Vergleich mit Meteonorm & NOAA. Lastadaten auf Basis von NOAA Daten. 2007', 'Code': '0007a', 'Latitude': '48.65', 'Source': 'Based on: STN EN ISO 13 790', 'Location': 'Sliac', 'Longitude': '19.15', 'Region': '', 'Country': 'SK'},
    {'Dataset': u'SK0008a-\xc5\xbdilina', 'Comments': '2007 PHI. Mit Meteonorm bearbeitet. Vergleich mit Meteonorm & NOAA. Lastadaten auf Basis von NOAA Daten. 2007', 'Code': '0008a', 'Latitude': '49.23', 'Source': 'Based on: STN EN ISO 13 790', 'Location': u'\xc5\xbdilina', 'Longitude': '18.61', 'Region': '', 'Country': 'SK'},
    {'Dataset': 'SY0001a-Aleppo/Neirab', 'Comments': '', 'Code': '0001a', 'Latitude': '36.183', 'Source': '', 'Location': 'Aleppo/Neirab', 'Longitude': '37.217', 'Region': '', 'Country': 'SY'},
    {'Dataset': 'TR0001b-Gaziantep', 'Comments': '2015 PHI', 'Code': '0001b', 'Latitude': '37.08', 'Source': 'Temp = 1981-2010; Other = Meteonorm. ', 'Location': 'Gaziantep', 'Longitude': '37.37', 'Region': '', 'Country': 'TR'},
    {'Dataset': 'UA0001a-Kiev', 'Comments': '', 'Code': '0001a', 'Latitude': '50.45', 'Source': '', 'Location': 'Kiev', 'Longitude': '30.5', 'Region': '', 'Country': 'UA'},
    {'Dataset': 'US0001a-Birmingham', 'Comments': '2009 PHI', 'Code': '0001a', 'Latitude': '33.5', 'Source': 'Source: Meteonorm', 'Location': 'Birmingham', 'Longitude': '-86.92', 'Region': 'Alabama', 'Country': 'US'},
    {'Dataset': 'US0002a-Anchorage', 'Comments': '2009 PHI', 'Code': '0002a', 'Latitude': '61.16', 'Source': 'Source: Meteonorm', 'Location': 'Anchorage', 'Longitude': '-150', 'Region': 'Alaska', 'Country': 'US'},
    {'Dataset': 'US0003a-Kodiak', 'Comments': '2009 PHI', 'Code': '0003a', 'Latitude': '57.75', 'Source': 'Source: Meteonorm', 'Location': 'Kodiak', 'Longitude': '-152.5', 'Region': 'Alaska', 'Country': 'US'},
    {'Dataset': 'US0004a-Little Rock', 'Comments': '2009 PHI', 'Code': '0004a', 'Latitude': '34.7', 'Source': 'Source: Meteonorm', 'Location': 'Little Rock', 'Longitude': '-92.28', 'Region': 'Arkansas', 'Country': 'US'},
    {'Dataset': 'US0005a-Phoenix', 'Comments': '2009 PHI', 'Code': '0005a', 'Latitude': '33.5', 'Source': 'Source: Meteonorm', 'Location': 'Phoenix', 'Longitude': '-112.17', 'Region': 'Arizona', 'Country': 'US'},
    {'Dataset': 'US0006a-Tucson', 'Comments': '2009 PHI', 'Code': '0006a', 'Latitude': '32.25', 'Source': 'Source: Meteonorm', 'Location': 'Tucson', 'Longitude': '-110.95', 'Region': 'Arizona', 'Country': 'US'},
    {'Dataset': 'US0007a-Bakersfield', 'Comments': '2009 PHI', 'Code': '0007a', 'Latitude': '35.37', 'Source': 'Source: Meteonorm', 'Location': 'Bakersfield', 'Longitude': '-119.02', 'Region': 'California', 'Country': 'US'},
    {'Dataset': 'US0008a-Fresno', 'Comments': '2009 PHI', 'Code': '0008a', 'Latitude': '36.68', 'Source': 'Source: Meteonorm', 'Location': 'Fresno', 'Longitude': '-119.78', 'Region': 'California', 'Country': 'US'},
    {'Dataset': 'US0009b-Los Angeles', 'Comments': '2015 PHI', 'Code': '0009b', 'Latitude': '34.05', 'Source': 'Temp = 1981-2010; Other monthly = Meteonorm; Load data derived by PHI. ', 'Location': 'Los Angeles', 'Longitude': '-118.24', 'Region': 'California', 'Country': 'US'},
    {'Dataset': 'US0010a-Sacramento', 'Comments': '2009 PHI', 'Code': '0010a', 'Latitude': '38.65', 'Source': 'Source: Meteonorm', 'Location': 'Sacramento', 'Longitude': '-121.5', 'Region': 'California', 'Country': 'US'},
    {'Dataset': 'US0011a-San Diego', 'Comments': '2009 PHI', 'Code': '0011a', 'Latitude': '32.83', 'Source': 'Source: Meteonorm', 'Location': 'San Diego', 'Longitude': '-117.17', 'Region': 'California', 'Country': 'US'},
    {'Dataset': 'US0012a-San Francisco', 'Comments': '2009 PHI', 'Code': '0012a', 'Latitude': '37.75', 'Source': 'Source: Meteonorm', 'Location': 'San Francisco', 'Longitude': '-122.45', 'Region': 'California', 'Country': 'US'},
    {'Dataset': 'US0013a-San Jose', 'Comments': '2009 PHI', 'Code': '0013a', 'Latitude': '37.33', 'Source': 'Source: Meteonorm', 'Location': 'San Jose', 'Longitude': '-122', 'Region': 'California', 'Country': 'US'},
    {'Dataset': 'US0014a-Colorado Springs', 'Comments': '2009 PHI', 'Code': '0014a', 'Latitude': '38.83', 'Source': 'Source: Meteonorm', 'Location': 'Colorado Springs', 'Longitude': '-104.83', 'Region': 'Colorado', 'Country': 'US'},
    {'Dataset': 'US0015a-Denver', 'Comments': '2009 PHI', 'Code': '0015a', 'Latitude': '39.75', 'Source': 'Source: Meteonorm', 'Location': 'Denver', 'Longitude': '-105', 'Region': 'Colorado', 'Country': 'US'},
    {'Dataset': 'US0016a-Hartford', 'Comments': '2009 PHI', 'Code': '0016a', 'Latitude': '41.77', 'Source': 'Source: Meteonorm', 'Location': 'Hartford', 'Longitude': '-72.68', 'Region': 'Connecticut', 'Country': 'US'},
    {'Dataset': 'US0017a-New Haven', 'Comments': '2009 PHI', 'Code': '0017a', 'Latitude': '41.33', 'Source': 'Source: Meteonorm', 'Location': 'New Haven', 'Longitude': '-72.9', 'Region': 'Connecticut', 'Country': 'US'},
    {'Dataset': 'US0018a-Washington', 'Comments': '2009 PHI', 'Code': '0018a', 'Latitude': '38.87', 'Source': 'Source: Meteonorm', 'Location': 'Washington', 'Longitude': '-77', 'Region': 'D.C.', 'Country': 'US'},
    {'Dataset': 'US0019a-Jacksonville', 'Comments': '2009 PHI', 'Code': '0019a', 'Latitude': '30.33', 'Source': 'Source: Meteonorm', 'Location': 'Jacksonville', 'Longitude': '-81.67', 'Region': 'Florida', 'Country': 'US'},
    {'Dataset': 'US0020a-Miami', 'Comments': '2009 PHI', 'Code': '0020a', 'Latitude': '25.87', 'Source': 'Source: Meteonorm', 'Location': 'Miami', 'Longitude': '-80.25', 'Region': 'Florida', 'Country': 'US'},
    {'Dataset': 'US0021a-Orlando', 'Comments': '2009 PHI', 'Code': '0021a', 'Latitude': '28.5', 'Source': 'Source: Meteonorm', 'Location': 'Orlando', 'Longitude': '-81.42', 'Region': 'Florida', 'Country': 'US'},
    {'Dataset': 'US0022a-Tampa', 'Comments': '2009 PHI', 'Code': '0022a', 'Latitude': '28.01', 'Source': 'Source: Meteonorm', 'Location': 'Tampa', 'Longitude': '-82.63', 'Region': 'Florida', 'Country': 'US'},
    {'Dataset': 'US0023a-Atlanta', 'Comments': '2009 PHI', 'Code': '0023a', 'Latitude': '33.83', 'Source': 'Source: Meteonorm', 'Location': 'Atlanta', 'Longitude': '-84.4', 'Region': 'Georgia', 'Country': 'US'},
    {'Dataset': 'US0024a-Augusta', 'Comments': '2009 PHI', 'Code': '0024a', 'Latitude': '33.367', 'Source': 'Source: Meteonorm', 'Location': 'Augusta', 'Longitude': '-81.967', 'Region': 'Georgia', 'Country': 'US'},
    {'Dataset': 'US0025a-Honolulu', 'Comments': '2009 PHI', 'Code': '0025a', 'Latitude': '21.32', 'Source': 'Source: Meteonorm', 'Location': 'Honolulu', 'Longitude': '-157.83', 'Region': 'Hawaii', 'Country': 'US'},
    {'Dataset': 'US0026a-Des Moines', 'Comments': '2009 PHI', 'Code': '0026a', 'Latitude': '41.58', 'Source': 'Source: Meteonorm', 'Location': 'Des Moines', 'Longitude': '-93.62', 'Region': 'Iowa', 'Country': 'US'},
    {'Dataset': 'US0027a-Boise City', 'Comments': '2009 PHI', 'Code': '0027a', 'Latitude': '36.73', 'Source': 'Source: Meteonorm', 'Location': 'Boise City', 'Longitude': '-102.52', 'Region': 'Idaho', 'Country': 'US'},
    {'Dataset': 'US0028a-Chicago', 'Comments': '2009 PHI', 'Code': '0028a', 'Latitude': '41.83', 'Source': 'Source: Meteonorm', 'Location': 'Chicago', 'Longitude': '-87.75', 'Region': 'Illinois', 'Country': 'US'},
    {'Dataset': 'US0029a-Fort Wayne', 'Comments': '2009 PHI', 'Code': '0029a', 'Latitude': '41.08', 'Source': 'Source: Meteonorm', 'Location': 'Fort Wayne', 'Longitude': '-85.13', 'Region': 'Indiana', 'Country': 'US'},
    {'Dataset': 'US0030a-Indianapolis', 'Comments': '2009 PHI', 'Code': '0030a', 'Latitude': '39.75', 'Source': 'Source: Meteonorm', 'Location': 'Indianapolis', 'Longitude': '-86.17', 'Region': 'Indiana', 'Country': 'US'},
    {'Dataset': 'US0031a-Wichita', 'Comments': '2009 PHI', 'Code': '0031a', 'Latitude': '37.72', 'Source': 'Source: Meteonorm', 'Location': 'Wichita', 'Longitude': '-97.33', 'Region': 'Kansas', 'Country': 'US'},
    {'Dataset': 'US0032a-Louisville', 'Comments': '2009 PHI', 'Code': '0032a', 'Latitude': '38.25', 'Source': 'Source: Meteonorm', 'Location': 'Louisville', 'Longitude': '-85.77', 'Region': 'Kentucky', 'Country': 'US'},
    {'Dataset': 'US0033a-Baton Rouge', 'Comments': '2009 PHI', 'Code': '0033a', 'Latitude': '30.5', 'Source': 'Source: Meteonorm', 'Location': 'Baton Rouge', 'Longitude': '-91.17', 'Region': 'Louisiana', 'Country': 'US'},
    {'Dataset': 'US0034a-New Orleans', 'Comments': '2009 PHI', 'Code': '0034a', 'Latitude': '30', 'Source': 'Source: Meteonorm', 'Location': 'New Orleans', 'Longitude': '-90.05', 'Region': 'Louisiana', 'Country': 'US'},
    {'Dataset': 'US0035a-Boston', 'Comments': '2009 PHI', 'Code': '0035a', 'Latitude': '42.33', 'Source': 'Source: Meteonorm', 'Location': 'Boston', 'Longitude': '-71.07', 'Region': 'Massachusetts', 'Country': 'US'},
    {'Dataset': 'US0036a-Baltimore', 'Comments': '2009 PHI', 'Code': '0036a', 'Latitude': '39.3', 'Source': 'Source: Meteonorm', 'Location': 'Baltimore', 'Longitude': '-76.62', 'Region': 'Maryland', 'Country': 'US'},
    {'Dataset': 'US0037a-Detroit', 'Comments': '2009 PHI', 'Code': '0037a', 'Latitude': '42.33', 'Source': 'Source: Meteonorm', 'Location': 'Detroit', 'Longitude': '-83.08', 'Region': 'Michigan', 'Country': 'US'},
    {'Dataset': 'US0038a-Grand Rapids', 'Comments': '2009 PHI', 'Code': '0038a', 'Latitude': '42.97', 'Source': 'Source: Meteonorm', 'Location': 'Grand Rapids', 'Longitude': '-85.67', 'Region': 'Michigan', 'Country': 'US'},
    {'Dataset': 'US0039a-Duluth', 'Comments': '2009 PHI', 'Code': '0039a', 'Latitude': '46.83', 'Source': 'Source: Meteonorm', 'Location': 'Duluth', 'Longitude': '-92.18', 'Region': 'Minnesota', 'Country': 'US'},
    {'Dataset': 'US0040a-Minneapolis', 'Comments': '2009 PHI', 'Code': '0040a', 'Latitude': '44.97', 'Source': 'Source: Meteonorm', 'Location': 'Minneapolis', 'Longitude': '-93.33', 'Region': 'Minnesota', 'Country': 'US'},
    {'Dataset': 'US0042a-Kansas City', 'Comments': '2009 PHI', 'Code': '0042a', 'Latitude': '39.05', 'Source': 'Source: Meteonorm', 'Location': 'Kansas City', 'Longitude': '-94.5', 'Region': 'Missouri', 'Country': 'US'},
    {'Dataset': 'US0043a-St. Louis', 'Comments': '2009 PHI', 'Code': '0043a', 'Latitude': '38.62', 'Source': 'Source: Meteonorm', 'Location': 'St. Louis', 'Longitude': '-90.2', 'Region': 'Missouri', 'Country': 'US'},
    {'Dataset': 'US0044a-Jackson', 'Comments': '2009 PHI', 'Code': '0044a', 'Latitude': '32.33', 'Source': 'Source: Meteonorm', 'Location': 'Jackson', 'Longitude': '-90.18', 'Region': 'Mississippi', 'Country': 'US'},
    {'Dataset': 'US0045a-Charlotte', 'Comments': '2009 PHI', 'Code': '0045a', 'Latitude': '35.22', 'Source': 'Source: Meteonorm', 'Location': 'Charlotte', 'Longitude': '-80.85', 'Region': 'North Carolina', 'Country': 'US'},
    {'Dataset': 'US0046a-Raleigh', 'Comments': '2009 PHI', 'Code': '0046a', 'Latitude': '35.77', 'Source': 'Source: Meteonorm', 'Location': 'Raleigh', 'Longitude': '-78.63', 'Region': 'North Carolina', 'Country': 'US'},
    {'Dataset': 'US0047a-Winston-Salem', 'Comments': '2009 PHI', 'Code': '0047a', 'Latitude': '36.08', 'Source': 'Source: Meteonorm', 'Location': 'Winston-Salem', 'Longitude': '-80.3', 'Region': 'North Carolina', 'Country': 'US'},
    {'Dataset': 'US0048a-Omaha', 'Comments': '2009 PHI', 'Code': '0048a', 'Latitude': '41.25', 'Source': 'Source: Meteonorm', 'Location': 'Omaha', 'Longitude': '-96', 'Region': 'Nebraska', 'Country': 'US'},
    {'Dataset': 'US0050a-Albuquerque', 'Comments': '2009 PHI', 'Code': '0050a', 'Latitude': '35.08', 'Source': 'Source: Meteonorm', 'Location': 'Albuquerque', 'Longitude': '-106.63', 'Region': 'New Mexico', 'Country': 'US'},
    {'Dataset': 'US0051a-Rosewell', 'Comments': '2009 PHI', 'Code': '0051a', 'Latitude': '33.3', 'Source': 'Source: Meteonorm', 'Location': 'Rosewell', 'Longitude': '-104.53', 'Region': 'New Mexico', 'Country': 'US'},
    {'Dataset': 'US0052a-Las Vegas', 'Comments': '2009 PHI', 'Code': '0052a', 'Latitude': '36.17', 'Source': 'Source: Meteonorm', 'Location': 'Las Vegas', 'Longitude': '-115.17', 'Region': 'Neavda', 'Country': 'US'},
    {'Dataset': 'US0053a-Reno', 'Comments': '2009 PHI', 'Code': '0053a', 'Latitude': '39.53', 'Source': 'Source: Meteonorm', 'Location': 'Reno', 'Longitude': '-119.82', 'Region': 'Neavda', 'Country': 'US'},
    {'Dataset': 'US0054a-Buffalo', 'Comments': '2009 PHI', 'Code': '0054a', 'Latitude': '42.88', 'Source': 'Source: Meteonorm', 'Location': 'Buffalo', 'Longitude': '-78.88', 'Region': 'New York', 'Country': 'US'},
    {'Dataset': 'US0055b-New York', 'Comments': '2016 PHI', 'Code': '0055b', 'Latitude': '40.78', 'Source': 'Temp = 1981-2010; Other derivede from Meteonrom and TMY3', 'Location': 'New York', 'Longitude': '-73.97', 'Region': 'New York', 'Country': 'US'},
    {'Dataset': 'US0056b-Rochester', 'Comments': '2015 PHI', 'Code': '0056b', 'Latitude': '43.12', 'Source': 'Temp = Normals Data 1981-2010. Radiation & Load data based on TMY3 (Class I). ', 'Location': 'Rochester', 'Longitude': '-77.68', 'Region': 'New York', 'Country': 'US'},
    {'Dataset': 'US0057a-Cincinnati', 'Comments': '2009 PHI', 'Code': '0057a', 'Latitude': '39.17', 'Source': 'Source: Meteonorm', 'Location': 'Cincinnati', 'Longitude': '-84.43', 'Region': 'Ohio', 'Country': 'US'},
    {'Dataset': 'US0058a-Cleveland', 'Comments': '2009 PHI', 'Code': '0058a', 'Latitude': '41.47', 'Source': 'Source: Meteonorm', 'Location': 'Cleveland', 'Longitude': '-81.72', 'Region': 'Ohio', 'Country': 'US'},
    {'Dataset': 'US0059a-Columbus', 'Comments': '2009 PHI', 'Code': '0059a', 'Latitude': '39.98', 'Source': 'Source: Meteonorm', 'Location': 'Columbus', 'Longitude': '-83.05', 'Region': 'Ohio', 'Country': 'US'},
    {'Dataset': 'US0060a-Toledo', 'Comments': '2009 PHI', 'Code': '0060a', 'Latitude': '41.67', 'Source': 'Source: Meteonorm', 'Location': 'Toledo', 'Longitude': '-83.58', 'Region': 'Ohio', 'Country': 'US'},
    {'Dataset': 'US0061a-Oklahoma City', 'Comments': '2009 PHI', 'Code': '0061a', 'Latitude': '35.47', 'Source': 'Source: Meteonorm', 'Location': 'Oklahoma City', 'Longitude': '-97.55', 'Region': 'Oklahoma', 'Country': 'US'},
    {'Dataset': 'US0062a-Tulsa', 'Comments': '2009 PHI', 'Code': '0062a', 'Latitude': '36.12', 'Source': 'Source: Meteonorm', 'Location': 'Tulsa', 'Longitude': '-95.97', 'Region': 'Oklahoma', 'Country': 'US'},
    {'Dataset': 'US0063a-Portland', 'Comments': '2009 PHI', 'Code': '0063a', 'Latitude': '45.53', 'Source': 'Source: Meteonorm', 'Location': 'Portland', 'Longitude': '-122.67', 'Region': 'Oregon', 'Country': 'US'},
    {'Dataset': 'US0064a-Philadelphia', 'Comments': '2009 PHI', 'Code': '0064a', 'Latitude': '40', 'Source': 'Source: Meteonorm', 'Location': 'Philadelphia', 'Longitude': '-75.17', 'Region': 'Pennsylvania', 'Country': 'US'},
    {'Dataset': 'US0065b-Pittsburgh', 'Comments': '2015 PHI', 'Code': '0065b', 'Latitude': '40.5', 'Source': 'Temp = 1981-2010; Other = Meteonorm / TMY3', 'Location': 'Pittsburgh', 'Longitude': '-80.08', 'Region': 'Pennsylvania', 'Country': 'US'},
    {'Dataset': 'US0066a-Providence', 'Comments': '2009 PHI', 'Code': '0066a', 'Latitude': '41.73', 'Source': 'Source: Meteonorm', 'Location': 'Providence', 'Longitude': '-71.25', 'Region': 'Rhode Island', 'Country': 'US'},
    {'Dataset': 'US0067a-Charleston', 'Comments': '2009 PHI', 'Code': '0067a', 'Latitude': '32.9', 'Source': 'Source: Meteonorm', 'Location': 'Charleston', 'Longitude': '-80.033', 'Region': 'South Carolina', 'Country': 'US'},
    {'Dataset': 'US0068a-Sioux Falls', 'Comments': '2009 PHI', 'Code': '0068a', 'Latitude': '43.567', 'Source': 'Source: Meteonorm', 'Location': 'Sioux Falls', 'Longitude': '-96.733', 'Region': 'South Dakota', 'Country': 'US'},
    {'Dataset': 'US0069a-Memphis', 'Comments': '2009 PHI', 'Code': '0069a', 'Latitude': '35.12', 'Source': 'Source: Meteonorm', 'Location': 'Memphis', 'Longitude': '-90', 'Region': 'Tennessee', 'Country': 'US'},
    {'Dataset': 'US0070a-Nashville', 'Comments': '2009 PHI', 'Code': '0070a', 'Latitude': '36.2', 'Source': 'Source: Meteonorm', 'Location': 'Nashville', 'Longitude': '-86.77', 'Region': 'Tennessee', 'Country': 'US'},
    {'Dataset': 'US0071a-Amarillo', 'Comments': '2009 PHI', 'Code': '0071a', 'Latitude': '35.23', 'Source': 'Source: Meteonorm', 'Location': 'Amarillo', 'Longitude': '-101.83', 'Region': 'Texas', 'Country': 'US'},
    {'Dataset': 'US0072a-Austin', 'Comments': '2009 PHI', 'Code': '0072a', 'Latitude': '30.33', 'Source': 'Source: Meteonorm', 'Location': 'Austin', 'Longitude': '-97.75', 'Region': 'Texas', 'Country': 'US'},
    {'Dataset': 'US0073a-Corpus Christi', 'Comments': '2009 PHI', 'Code': '0073a', 'Latitude': '28', 'Source': 'Source: Meteonorm', 'Location': 'Corpus Christi', 'Longitude': '-97.9', 'Region': 'Texas', 'Country': 'US'},
    {'Dataset': 'US0074a-Dallas', 'Comments': '2009 PHI', 'Code': '0074a', 'Latitude': '32.83', 'Source': 'Source: Meteonorm', 'Location': 'Dallas', 'Longitude': '-96.83', 'Region': 'Texas', 'Country': 'US'},
    {'Dataset': 'US0076a-Houston', 'Comments': '2009 PHI', 'Code': '0076a', 'Latitude': '29.83', 'Source': 'Source: Meteonorm', 'Location': 'Houston', 'Longitude': '-95.33', 'Region': 'Texas', 'Country': 'US'},
    {'Dataset': 'US0077a-San Antonio', 'Comments': '2009 PHI', 'Code': '0077a', 'Latitude': '29.5', 'Source': 'Source: Meteonorm', 'Location': 'San Antonio', 'Longitude': '-98.5', 'Region': 'Texas', 'Country': 'US'},
    {'Dataset': 'US0078a-Provo', 'Comments': '2009 PHI', 'Code': '0078a', 'Latitude': '40.22', 'Source': 'Source: Meteonorm', 'Location': 'Provo', 'Longitude': '-111.72', 'Region': 'Utah', 'Country': 'US'},
    {'Dataset': 'US0079a-Salt Lake City', 'Comments': '2009 PHI', 'Code': '0079a', 'Latitude': '40.75', 'Source': 'Source: Meteonorm', 'Location': 'Salt Lake City', 'Longitude': '-111.92', 'Region': 'Utah', 'Country': 'US'},
    {'Dataset': 'US0080a-Norfolk', 'Comments': '2009 PHI', 'Code': '0080a', 'Latitude': '36.85', 'Source': 'Source: Meteonorm', 'Location': 'Norfolk', 'Longitude': '-76.28', 'Region': 'Virginia', 'Country': 'US'},
    {'Dataset': 'US0081a-Virginia Beach', 'Comments': '2009 PHI', 'Code': '0081a', 'Latitude': '36.85', 'Source': 'Source: Meteonorm', 'Location': 'Virginia Beach', 'Longitude': '-75.98', 'Region': 'Virginia', 'Country': 'US'},
    {'Dataset': 'US0082a-Seattle', 'Comments': '2009 PHI', 'Code': '0082a', 'Latitude': '47.58', 'Source': 'Source: Meteonorm', 'Location': 'Seattle', 'Longitude': '-122.33', 'Region': 'Washington', 'Country': 'US'},
    {'Dataset': 'US0083a-Spokane', 'Comments': '2009 PHI', 'Code': '0083a', 'Latitude': '47.67', 'Source': 'Source: Meteonorm', 'Location': 'Spokane', 'Longitude': '-117.42', 'Region': 'Washington', 'Country': 'US'},
    {'Dataset': 'US0084a-Madison', 'Comments': '2009 PHI', 'Code': '0084a', 'Latitude': '43.08', 'Source': 'Source: Meteonorm', 'Location': 'Madison', 'Longitude': '-89.42', 'Region': 'Wisconsin', 'Country': 'US'},
    {'Dataset': 'US0085a-Milwaukee', 'Comments': '2009 PHI', 'Code': '0085a', 'Latitude': '43.05', 'Source': 'Source: Meteonorm', 'Location': 'Milwaukee', 'Longitude': '-87.93', 'Region': 'Wisconsin', 'Country': 'US'},
    {'Dataset': 'US0086a-Burlington ', 'Comments': '2014 PHI', 'Code': '0086a', 'Latitude': '44.47', 'Source': 'Source: TMY3. ', 'Location': 'Burlington ', 'Longitude': '-73.15', 'Region': 'Vermont', 'Country': 'US'},
    {'Dataset': 'US0087a-Beaver Island', 'Comments': '2013 PHI. Compared with satellite and TMY data of the region. Longterm locally measured data not available. Load data: EOSWEB (on the safe side)', 'Code': '0087a', 'Latitude': '45.67', 'Source': 'Source: Meteonorm V7 Interpolation (User Defined Site). Use carefully. PHI December 2013.', 'Location': 'Beaver Island', 'Longitude': '-85.53', 'Region': 'Michigan', 'Country': 'US'},
    {'Dataset': 'US0088a-Traverse City', 'Comments': '2014 PHI', 'Code': '0088a', 'Latitude': '44.73', 'Source': 'Source: TMY3 (Category I)', 'Location': 'Traverse City', 'Longitude': '-85.58', 'Region': 'Michigan', 'Country': 'US'},
    {'Dataset': 'US0089a-Marthas Vineyard', 'Comments': '2015 PHI', 'Code': '0089a', 'Latitude': '41.4', 'Source': 'Source: MeteonormV7. Load data derived from TMY3. ', 'Location': 'Marthas Vineyard', 'Longitude': '-70.61', 'Region': 'Massachusetts', 'Country': 'US'},
    {'Dataset': 'US0090a-South Fallsburg', 'Comments': '2012 PHI', 'Code': '0090a', 'Latitude': '41.7', 'Source': 'Source: EOSWEB (Dew Point = TMY3)', 'Location': 'South Fallsburg', 'Longitude': '-74.6', 'Region': 'New York', 'Country': 'US'},
    {'Dataset': 'US0091a-Medford', 'Comments': '2015 PHI', 'Code': '0091a', 'Latitude': '42.38', 'Source': 'Temp: 1981-2010 US Normals Data. Other: Meteonorm. Load data derived by PHI. ', 'Location': 'Medford', 'Longitude': '-122.87', 'Region': 'Oregon', 'Country': 'US'},
    {'Dataset': 'US0092a-Aspen', 'Comments': '2015 PHI.', 'Code': '0092a', 'Latitude': '39.21', 'Source': 'Source: Meteonorm7. Load data derived by PHI. ', 'Location': 'Aspen', 'Longitude': '-106.86', 'Region': 'Colorado', 'Country': 'US'},
    {'Dataset': 'US0094a-Gunnison', 'Comments': '2015 PHI', 'Code': '0094a', 'Latitude': '38.53', 'Source': 'Temp: 1981-2010 US Normals Data. Other: Meteonorm. Load data derived by PHI. ', 'Location': 'Gunnison', 'Longitude': '-106.97', 'Region': 'Colorado', 'Country': 'US'},
    {'Dataset': 'US0095a-Allentown', 'Comments': '2015 PHI', 'Code': '0095a', 'Latitude': '40.65', 'Source': 'Temp = 1981-2010; Load data derived by PHI with reference to TMY3 data. ', 'Location': 'Allentown', 'Longitude': '-75.45', 'Region': 'Pennsylvania', 'Country': 'US'},
    {'Dataset': 'US0096a-Lewistown', 'Comments': '2015 PHI', 'Code': '0096a', 'Latitude': '40.59', 'Source': 'Temp = 1981-2010; Other = Meteonorm & Satellite data. ', 'Location': 'Lewistown', 'Longitude': '-77.57', 'Region': 'Pennsylvania', 'Country': 'US'},
    {'Dataset': 'US0097a-Elmira', 'Comments': '2015 PHI', 'Code': '0097a', 'Latitude': '42.1', 'Source': 'Temp = 1981-2010; Load data derived by PHI. ', 'Location': 'Elmira', 'Longitude': '-76.84', 'Region': 'New York', 'Country': 'US'},
    {'Dataset': 'US0099a-Olympia', 'Comments': '2016 PHI', 'Code': '0099a', 'Latitude': '46.97', 'Source': 'Temp = 1981-2010. Other derived from Meteonorm & TRY3 data. ', 'Location': 'Olympia', 'Longitude': '-122.9', 'Region': 'Washington', 'Country': 'US'},
    {'Dataset': 'US0100a-Portland', 'Comments': '2013 PHI', 'Code': '0100a', 'Latitude': '43.65', 'Source': 'Derived from TMY3 and satellite data.', 'Location': 'Portland', 'Longitude': '-70.3', 'Region': 'Maine', 'Country': 'US'}
    ]

#-------------------------------------------------------------------------------
############    Utils    ###########
@contextmanager
def idf2ph_rhDoc():
    """ Switches the sc.doc to the Rhino Active Doc temporaily """
    
    try:
        sc.doc = Rhino.RhinoDoc.ActiveDoc
        yield
    finally:
        sc.doc = ghdoc

def preview(classObj):
    # For looking at the contents of a Class Object
    # Pass in any class obj and it'll sift through all the keys and print to the consol
    print '-------'
    for eachItem in classObj.__dict__.keys():
        print eachItem, "::", classObj.__dict__[eachItem]
        try:
            for each in classObj.__dict__[eachItem].__dict__.keys():
                print "   >", each, "::", classObj.__dict__[eachItem].__dict__[each]
        except:
            pass

#-------------------------------------------------------------------------------
############    Def    #############

def phpp_geomFromVerts(_idfObj):
    """
    Takes in an IDF Class Object and reads the Vertex information
    Builds new geometry from the vertex data provided
    
    Args:
        _idfObj: An IDF-Class Object from the IDF-Reader with some Vertex data to read
    Returns (list): 
        0: boundary (Polyline) the perimeter edges built from the vertex points
        1: srfc (Surface) the new surface built from the vertext points
        2: surfaceArea (m2)
        3: centroid (Point3d)
        4: normalVector (Vector3d) Normal for the new surface
    """
    
    vertsTXT = []
    vertsGH = []
    vertNumber = []
    for eachKey in _idfObj.__dict__.keys():
        if 'Vertex' in eachKey:
            numOnly = [int(i) for i in eachKey.split() if i.isdigit()] # Pulls num from "!- X,Y,Z Vertex 1 {m}"
            numOnly = numOnly[0]
            vertNumber.append( '{:03d}'.format(numOnly) )
            vertsTXT.append([eachKey, getattr(_idfObj, eachKey)])
    
    # Order verts using Vertex number as the key
    verst_orderd = [x for _,x in sorted(zip(vertNumber, vertsTXT))]
    
    for eachVert in verst_orderd:
        vertXYZ = eachVert[1].split(' ') # Create a list of the Vert positions only
        vertsGH.append( ghc.ConstructPoint(vertXYZ[0], vertXYZ[1], vertXYZ[2]) )
    
    boundary = ghc.PolyLine(vertsGH, True) # Create Closed PLine of the srfc boundary
    srfc = ghc.BoundarySurfaces(boundary) # Create the Surface Boundary from edge
    surfaceArea = ghc.Area(srfc)[0]  # Get the Surface Area (m2)
    centroid = ghc.Area(srfc)[1]
    normalVector = ghc.EvaluateSurface(srfc, centroid)[1]
    
    return boundary, srfc, surfaceArea, centroid, normalVector

def phpp_calcNorthAngle(_objNormVec, _refNorthVec):
    """ Takes in a Surface's Normal Vector and the project's north angle vector and computes the angle 0--360 between
    
    http://frasergreenroyd.com/obtaining-the-angle-between-two-vectors-for-360-degrees/
    Results 0=north, 90=east, 180=south, 270=west
    
    Args:
        _objNormVec: A Vector3d of the surface's normal
        _refNorthVec: A Vector3d of the project's north vector 
    Returns: 
        angle: the angle of North for the surface (Degrees)
    """
    
    # Get the input Vector's X and Y parts
    x1 = _objNormVec.X
    y1 = _objNormVec.Y
    
    x2 = _refNorthVec.X
    y2 = _refNorthVec.Y
    
    # Calc the angle between
    angle = math.atan2(y2, x2) - math.atan2(y1, x1)
    angle = angle * 360 / (2 * math.pi)
    
    if angle < 0:
        angle = angle + 360
    
    # Return Angle in Degrees
    return angle

def phpp_GetWindowSize(_geom):
    """ Takes in Brep Geometry and returns the width and height (maybe)
    
    Args:
        _geom: Some sort of Brep Geometry from Rhino or GH
    Returns (list): 
        0: width (float),
        1: height (float),
        2: windowGeom (Brep) 
    """
    try:
        windowGeom = rs.coercebrep(_geom)
    except:
        windowGeom = _geom
    
    dims = ghc.Dimensions(windowGeom)
    width = dims[1]
    height = dims[0]
    
    return width, height, windowGeom

def phpp_makeHBMaterial_Opaque(_name, _roughness='Rough', _thickness=1, _conductivity=1, _density=2500, _specHeat=460, _absTherm=0.9, _absSol=0.7, _absVis=0.7):
    """ This outputs the TXT for a normal Opaque EP Material """
    
    # Set up the inputs
    values = [_name.upper().replace(" ", "_"), _roughness, _thickness, _conductivity, _density, _specHeat, _absTherm, _absSol, _absVis]
    comments = ["Name", "Roughness", "Thickness {m}", "Conductivity {W/m-K}", "Density {kg/m3}","Specific Heat {J/kg-K}", "Thermal Absorptance","Solar Absorptance", "Visible Absorptance" ]
    
    # Create the Text output
    materialStr = "Material,\n"
    for count, (value, comment) in enumerate(zip(values, comments)):
        if count!= len(values) - 1:
            materialStr += str(value) + ",    !-" + str(comment) + "\n"
        else:
            materialStr += str(value) + ";    !-" + str(comment)
            
    return materialStr

def phpp_makeHBMaterial_NoMass(_name, _resistance):
    """ This outputs the TXT for a NoMass EP Material """
    
    # Set up the inputs
    values = [_name.upper().replace(" ", "_"), 'Rough', _resistance, 0.9, 0.7, 0.7]
    comments = ["Name", "Roughness", "Thermal Resistance {m2-K/W}", "Thermal Absorptance", "Solar Absorptance","Visible Absorptance"]
    
    # Create the Text output
    materialStr = "Material:NoMass,\n"
    for count, (value, comment) in enumerate(zip(values, comments)):
        if count!= len(values) - 1:
            materialStr += str(value) + ",    !-" + str(comment) + "\n"
        else:
            materialStr += str(value) + ";    !-" + str(comment)
            
    return materialStr

def phpp_makeHBMaterial(name, U_Value, SHGC, VT):
    """ This is the same def from the 'EPWindowMat' HB Component
    
    Args:
        name: (str) Name for the new Material
        U_Value: (float) U-Value in W/m2-k
        SHGC: (float) Solar Heat Gain Coefficient
        VT: (float) Visible Transmittance
    Returns: 
        materialStr: An IDF-Style String to pass along to the HB Library writer. Plug into an 'addToEPLibrary' HB Component
    """
    
    values = [name.upper(), U_Value, SHGC, VT]
    comments = ["Name", "U Value", "Solar Heat Gain Coeff", "Visible Transmittance"]
    
    materialStr = "WindowMaterial:SimpleGlazingSystem,\n"
    
    for count, (value, comment) in enumerate(zip(values, comments)):
        if count!= len(values) - 1:
            materialStr += str(value) + ",    !-" + str(comment) + "\n"
        else:
            materialStr += str(value) + ";    !-" + str(comment)
            
    return materialStr

def phpp_makeHBConstruction(_nm, _layer1, _layer2=None, _layer3=None):
    """ Simplified version of the HB 'EPCOnstruction' Component.
    
    Note: Since its for a window, only one layer material is used
    
    Args:
        _nm: (str) Name for the new Construction
        _layer1: (str) The Name of the Material to use in Layer 1
    Returns: 
        constructionStr: An IDF-Style String to pass along to the HB Library writer. Plug into an 'addToEPLibrary' HB Component
    """
    
    constructionStr = "Construction,\n" + _nm.upper() + ",    !- Name\n"
    constructionStr += _layer1 + ",    !- Layer 1" + "\n"
    if _layer2 != None:
        constructionStr += _layer2 + ",    !- Layer 2" + "\n"
    if _layer3 != None:
        constructionStr += _layer3 + ",    !- Layer 3" + "\n"
    
    return constructionStr

def phpp_getWindowLibraryFromRhino():
    """ Loads any window object entries from the DocumentUseText library from the Active Rhino document 
    
    Args:
        None
    Returns:
        PHPPLibrary_ (dict): A dictionary of all the window entries found with their parameters
    """
    
    PHPPLibrary_ = {}
    
    with idf2ph_rhDoc():
        lib_GlazingTypes = {}
        lib_FrameTypes = {}
        lib_PsiInstalls = {}
        
        # First, try and pull in the Rhino Document's PHPP Library
        # And make new Frame and Glass Objects. Add all of em' to new dictionaries
        if rs.IsDocumentUserText():
            print 'Getting Window Library data from the Rhino Document'
            for eachKey in rs.GetDocumentUserText():
                if 'PHPP_lib_Glazing' in eachKey:
                    tempDict = json.loads(rs.GetDocumentUserText(eachKey))
                    newGlazingObject = PHPP_Glazing(
                                    tempDict['Name'],
                                    tempDict['gValue'],
                                    tempDict['uValue']
                                    )
                    lib_GlazingTypes[tempDict['Name']] = newGlazingObject
                elif '_PsiInstall_' in eachKey:
                    tempDict = json.loads(rs.GetDocumentUserText(eachKey))
                    newPsiInstallObject = PHPP_Window_Install(
                                    [
                                    tempDict['Left'],
                                    tempDict['Right'],
                                    tempDict['Bottom'],
                                    tempDict['Top']
                                    ]
                                    )
                    lib_PsiInstalls[tempDict['Typename']] = newPsiInstallObject
                elif 'PHPP_lib_Frame' in eachKey:
                    tempDict = json.loads(rs.GetDocumentUserText(eachKey))
                    newFrameObject = PHPP_Frame(
                                    tempDict['Name'],
                                    [
                                    tempDict['uFrame_L'],
                                    tempDict['uFrame_R'],
                                    tempDict['uFrame_B'],
                                    tempDict['uFrame_T']
                                    ],
                                    [
                                    tempDict['wFrame_L'],
                                    tempDict['wFrame_R'],
                                    tempDict['wFrame_B'],
                                    tempDict['wFrame_T']
                                    ],
                                    [
                                    tempDict['psiG_L'],
                                    tempDict['psiG_R'],
                                    tempDict['psiG_B'],
                                    tempDict['psiG_T']
                                    ],
                                    [
                                    tempDict['psiInst_L'],
                                    tempDict['psiInst_R'],
                                    tempDict['psiInst_B'],
                                    tempDict['psiInst_T']
                                    ]
                                    )
                    lib_FrameTypes[tempDict['Name']] = newFrameObject
            
            PHPPLibrary_['lib_GlazingTypes'] = lib_GlazingTypes
            PHPPLibrary_['lib_FrameTypes'] = lib_FrameTypes
            PHPPLibrary_['lib_PsiInstalls'] = lib_PsiInstalls
    
    return PHPPLibrary_

def phpp_createSrfcHBMatAndConst(_constName, _conductance, _intInsul):
    """
    Take in construction Params and create EP Materials and EP Constructions
    so they can be added to the HB Hive / Master dataset.
    Note that this creates a simplified 'sandwich' assembly with a thin (1") 
    mass layer inside and out, and a single 'NoMass' thermal control layer in 
    the center. For reference on this method, see:
    https://www.energyplus.net/sites/default/files/docs/site_v8.3.0/EngineeringReference/03-SurfaceHeatBalance/index.html
    section "Conduction Transfer Function (CTF) Calculations Special Case: R-Value Only Layers"
    -----
    Inputs:
        _constName: The name of the Construction
        _conductance: The Conductance Value of the Construction (W/m2k)
        _intInsul: 'x' = IS internal insulation, ''= Is NOT internal insulation
    Returns (tuple): 
        newHBMat: The New HB Material to add to the Hive
        newHBMat_Mass: A New 'Mass' Material to add to the Hive
        constructionName: The cleaned up name for the Construction
        new_EPConstruction: The new HB Construction to add to the Hive 
    """
    
    # Create the Names, clean up
    constructionName = "PHPP_CONST_{}".format(_constName).upper().replace(" ", "_")
    matName = "PHPP_MAT_{}".format(_constName).upper().replace(" ", "_")
    
    # Add the tag to the name if Interior Insulated
    intInsulFlag = '__Int__'
    if _intInsul == 1:
        constructionName = constructionName + intInsulFlag
        matName = matName + intInsulFlag
    
    # Create a new NoMass HB Material for the center of the Constructon
    matResistance = 1 / _conductance # = Resistance (m2k/w)
    newHBMat = phpp_makeHBMaterial_NoMass(
            matName,
            matResistance # m2k/W
            )
    
    # Create the new MassLayer to use on the interior and exterior faces
    newHBMat_Mass = phpp_makeHBMaterial_Opaque(_name="MASSLAYER", _thickness=0.0254, _conductivity=2)
    
    # Create a new HB Construction with a matching name
    new_EPConstruction = phpp_makeHBConstruction(_nm=constructionName, _layer1='MASSLAYER', _layer2=matName, _layer3='MASSLAYER')
    
    return newHBMat, newHBMat_Mass, constructionName, new_EPConstruction

#-------------------------------------------------------------------------------
###### Classes for PHPP Objects ####
class PHPP_WindowObject:
    """For storing Window Object geom and parameters for the Frame, Glass and Installation."""
    
    def __init__(self, _nm, _glassType, _frameType, _installs, _geom=None, _varType='a', _instDepth=0.1):
        """
        Args:
            _nm (str): Name of the Window Object
            _glassType (obj): A 'PHPP_Glazing' object with glass parameters
            _frameType (obj): A 'PHPP_Frame' oject with frame parameters
            _installs (obj): A 'PHPP_Window_Install' object with information about how the window is installed in the surface
            _geom (geom): The Rhino geometry of the window
            _varType (str): The Variant 'Type' (a, b, c, ...). Only useful if using the 'Variants' worksheet in PHPP
            _instDepth (float): The depth (m) the window is inset into the wall measured from ext. face of wall to ext. face of glass
        """
        self.Name = _nm
        self.Type_Glass = _glassType
        self.Type_Frame = _frameType
        self.Installs = _installs
        self.Geometry = _geom
        self.Type_Variant = _varType
        self.InstallDepth = float(_instDepth)
        self.OldShadingFac_Winter = 0.75 # Defaults
        self.OldShadingFac_Summer = 0.75
        
        # Figure out all the geometry and parameter values
        self.setWindowParams()
        self.intPts = []
    
    def setWindowParams(self):
        
        self.UwInstalled = self.getUwInstalled()
        self.Geometry_Inset = self.getInsetWindowSurface(True)
        self.SurfaceNormal = self.getSurfaceNormal(self.Geometry)
        self.GlazingSrfc = self.getGlazingSurface()
        
        self.Edge_Bottom, self.Edge_Left, self.Edge_Top, self.Edge_Right = self.getEdgesInOrder(self.Geometry)
        self.WindowHeight = ghc.Length(self.Edge_Left)
        self.WindowWidth = ghc.Length(self.Edge_Bottom)
        self.GlazingHeight = self.WindowHeight - self.Type_Frame.fTop - self.Type_Frame.fBottom
        self.GlazingWidth = self.WindowWidth - self.Type_Frame.fLeft - self.Type_Frame.fRight
        
        self.AngleFromHoriz = ghc.Degrees(ghc.Angle(self.SurfaceNormal, ghc.UnitZ(1)))[0]
        self.Azimuth = phpp_calcNorthAngle(self.SurfaceNormal, ghc.UnitY(1)) # Assumes Y is North, should get this from Zone....
    
    def calcShadingFactor_Simple(self, _shadingGeom, _lat=40):
        self.Winter_ShadingFactor = 0.75
        self.Summer_ShadingFactor = 0.75
        
        # ----------------------------------------------------------------------
        # Find the relevant geometry in the scene and figure out the critical dimensions
        self.h_hori, self.d_hori, self.Checkline_hori= self.findHorizontalShadingValues(_shadingGeom, 99)
        self.d_over, self.o_over, self.Checkline_over = self.findOverhangShading(_shadingGeom, 99)
        self.o_reveal, self.d_reveal, self.Checkline_side1, self.Checkline_side2 = self.findRevealShading(_shadingGeom, 99)
        
        
        # ----------------------------------------------------------------------
        # Calc the actual Shading Factors based on the geometry found in the scene
        # This re-creates the PHPP v9.6a shading factor algortithms. I think Andrew Peel made these?
        calculator = PHPP_Window_Shading_Calculator(_lat)
        
        self.OldShadingFac_Winter_Horiz = calculator.Winter_HorizShadingFactor(self.h_hori, self.d_hori, self.AngleFromHoriz, self.Azimuth, self.GlazingHeight )
        self.OldShadingFac_Summer_Horiz = calculator.Summer_HorizShadingFactor(self.h_hori, self.d_hori, self.AngleFromHoriz, self.Azimuth, self.GlazingHeight )
        
        self.OldShadingFac_Winter_Reveal = calculator.Winter_RevealShadingFactor(self.o_reveal, self.d_reveal, self.AngleFromHoriz, self.Azimuth, self.GlazingWidth)
        self.OldShadingFac_Summer_Reveal = calculator.Summer_RevealShadingFactor(self.o_reveal, self.d_reveal, self.AngleFromHoriz, self.Azimuth, self.GlazingWidth)
        
        self.OldShadingFac_Winter_Overhead = calculator.Winter_OverhangShadingFactor(self.o_over, self.d_over, self.AngleFromHoriz, self.Azimuth, self.GlazingHeight  )
        self.OldShadingFac_Summer_Overhead = calculator.Summer_OverhangShadingFactor(self.o_over, self.d_over, self.AngleFromHoriz, self.Azimuth, self.GlazingHeight  )
        
        self.OldShadingFac_Winter = self.OldShadingFac_Winter_Horiz * self.OldShadingFac_Winter_Reveal * self.OldShadingFac_Winter_Overhead
        self.OldShadingFac_Summer = self.OldShadingFac_Summer_Horiz * self.OldShadingFac_Summer_Reveal * self.OldShadingFac_Summer_Overhead
    
    def getShadingFactors_Simple(self):
        return self.OldShadingFac_Winter, self.OldShadingFac_Summer
    
    def setShadingFactors(self, _winterFac, _summerFac):
        """Sets Winter and Summer shading Factors. 0=fully shaded, 1=fully unshaded
        Arguments:
            _winterFac (Float): A single value representing the winter shading factor (0-1.0)
            _winterFac (Float): A single value representing the summer shading factor (0-1.0)
        Returns:
            None
        """
        self.OldShadingFac_Winter = _winterFac
        self.OldShadingFac_Summer = _summerFac
    
    def getUwInstalled(self):
        """ Caculates the U-W-Installed values for a window using the PHPP method
        
        Args:
            _geom: Some sort of Brep Geometry from Rhino or GH, a planar surface of some sort
            _frame: a 'PHPP_Frame' Class Object with frame params
            _glass: a 'PHPP_Glazing' Class Object with glazing params
            _installs: an 'Installs' Class Object with Install conditions (0|1) for all four sides
        Returns: 
            uw_inst: a U-W-Installed value (float) in W/m2-k taking into account all the details input
        """
        
        _geom = self.Geometry
        _frame = self.Type_Frame
        _glass = self.Type_Glass
        _installs = self.Installs
        
        # Sort out the Geometry of the Window and Params
        winEdges = ghc.DeconstructBrep(rs.coercebrep(_geom))[1]
        winBoundary = ghc.JoinCurves(winEdges, preserve=False)
        plane, xInterval, yInterval = ghc.DeconstuctRectangle(winBoundary)
        
        len_inst_Left = xInterval[1]
        len_inst_Right = xInterval[1]
        len_inst_Bottom = yInterval[1]
        len_inst_Top = yInterval[1]
        
        len_glazEdg_Left = len_inst_Left - _frame.fBottom - _frame.fTop 
        len_glazEdg_Right = len_inst_Right - _frame.fBottom - _frame.fTop 
        len_glazEdg_Bottom = len_inst_Bottom - _frame.fLeft - _frame.fRight
        len_glazEdg_Top = len_inst_Top - _frame.fLeft - _frame.fRight
        
        a_win = ghc.Area(rs.coercebrep(_geom))[0]
        a_frame_left = len_inst_Left * _frame.fLeft
        a_frame_right = len_inst_Right * _frame.fRight
        a_frame_bottom = len_glazEdg_Bottom * _frame.fBottom
        a_frame_top = len_glazEdg_Top * _frame.fTop
        a_glass = a_win - a_frame_left - a_frame_right - a_frame_bottom - a_frame_top
        
        # Clac Heat Loss Values for all the elements
        hL_glass = _glass.uValue * a_glass
        
        hL_frame_Left = a_frame_left * _frame.uLeft
        hL_frame_Right = a_frame_right * _frame.uRight
        hL_frame_Bottom = a_frame_bottom * _frame.uBottom
        hL_frame_Top = a_frame_top * _frame.uTop
        hL_frame = hL_frame_Left + hL_frame_Right + hL_frame_Bottom + hL_frame_Top
        
        hL_psi_g_Left = len_glazEdg_Left * _frame.psigLeft
        hL_psi_g_Right = len_glazEdg_Right * _frame.psigRight
        hL_psi_g_Bottom = len_glazEdg_Bottom * _frame.psigBottom
        hL_psi_g_Top = len_glazEdg_Top * _frame.psigTop
        hL_psi_g = hL_psi_g_Left + hL_psi_g_Right + hL_psi_g_Bottom + hL_psi_g_Top
        
        # Update the Psi-Install values incase any UD overrides
        if _installs.Inst_L == 0 or _installs.Inst_L == 1:
            psiInstLeft = _frame.psiInstLeft * _installs.Inst_L
        else:
            psiInstLeft = _installs.Inst_L
        if _installs.Inst_R == 0 or _installs.Inst_R == 1:
            psiInstRight = _frame.psiInstRight * _installs.Inst_R
        else:
            psiInstRight = _installs.Inst_B
        if _installs.Inst_B == 0 or _installs.Inst_B == 1:
            psiInstBottom = _frame.psiInstLeft * _installs.Inst_B
        else:
            psiInstBottom = _installs.Inst_B
        if _installs.Inst_T == 0 or _installs.Inst_T == 1:
            psiInstTop = _frame.psiInstTop * _installs.Inst_T
        else:
            psiInstTop = _installs.Inst_T
        
        # Calc the total heat loss of the window installs
        hL_psi_inst_Left = len_inst_Left * psiInstLeft
        hL_psi_inst_Right = len_inst_Right * psiInstRight
        hL_psi_inst_Bottom = len_inst_Bottom * psiInstBottom
        hL_psi_inst_Top = len_inst_Top * psiInstTop
        hL_psi_inst = hL_psi_inst_Left + hL_psi_inst_Right + hL_psi_inst_Bottom + hL_psi_inst_Top
        
        """# For Preview
        print _installs.Inst_L, '>>', psiInstLeft, '>>',hL_psi_inst_Left
        print _installs.Inst_R, '>>', psiInstRight, '>>',hL_psi_inst_Right
        print _installs.Inst_B, '>>', psiInstBottom, '>>',hL_psi_inst_Bottom
        print _installs.Inst_T, '>>', psiInstTop, '>>',hL_psi_inst_Top
        """
        
        # Calculate the U-w-Intalled
        uw_inst = (hL_glass + hL_frame + hL_psi_g + hL_psi_inst)/ a_win
        
        return uw_inst
    
    def cleanExtrude(self, _geom, _direction, _extrudeDepth, _install):
        # Guards against 0 extrude
        if _install == 0 or _extrudeDepth == 0:
            return None
        else:
            return ghc.Extrude( _geom, ghc.Amplitude(_direction, _extrudeDepth) )
    
    def getInsetWindowSurface(self, _move):
        if _move == True:
            winSurface = ghc.Move(self.Geometry, ghc.Amplitude(self.getSurfaceNormal(self.Geometry), self.InstallDepth*-1)).geometry
        else:
            winSurface = self.Geometry
        
        return winSurface
    
    def getVerticalEdges(self):
        if 'Edge_Left' not in self.__dict__:
            self.Edge_Bottom, self.Edge_Left, self.Edge_Top, self.Edge_Right = self.getEdgesInOrder(self.Geometry)
        
        return self.Edge_Left, self.Edge_Right
    
    def getGlazingSurface(self):
        geom = self.Geometry_Inset
        
        # Create the frame and instet surfaces
        WinCenter = ghc.Area(geom).centroid
        WinNormalVector = self.getSurfaceNormal(geom)
        #WinNormalVector = ghc.EvaluateSurface(geom, WinCenter).normal
        
        #Create the Inset Curve
        WinEdges = ghc.DeconstructBrep(geom).edges
        WinPerimeter = ghc.JoinCurves(WinEdges, False)
        #FrameWidth = self.Type_Frame.frameWidths[0] # Use only the left width value?
        FrameWidth = self.Type_Frame.fLeft
        print ':::', FrameWidth, type(FrameWidth)
        InsetCurve = rs.OffsetCurve(WinPerimeter, WinCenter, FrameWidth, WinNormalVector, 0)
        
        # In case the curve goes 'out' and the offset fails
        # Or is too small and results in multiple offset Curves
        if len(InsetCurve)>1:
            warning = 'Error. The window named: "{}" is too small. The frame offset of {} m can not be done. Check the frame sizes?'.format(self.Name, FrameWidth)
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
            InsetCurve = rs.OffsetCurve(WinPerimeter, WinCenter, 0.05, WinNormalVector, 0)
            InsetCurve = rs.coercecurve( InsetCurve[0] )
        else:
            InsetCurve = rs.coercecurve( InsetCurve[0] )
        
        return ghc.BoundarySurfaces(InsetCurve)
    
    def findHorizontalShadingValues(self, _shadingGeom, _extents=99):
        """
        Arguments:
            _shadingGeom: (list) A list of possible shading objects to test against
            _extents: (float) A number (m) to limit the shading search to. Default = 99m
        Returns:
            h_hori: Distance (m) out from the glazing surface of any horizontal shading objects found
            d_hori: Distance (m) up from the base of the window to the top of any horizontal shading objects found
        """
        
        #Find Starting Point
        glazingEdges = self.getEdgesInOrder(self.GlazingSrfc)
        glazingBottomEdge = glazingEdges[0]
        ShadingOrigin = ghc.CurveMiddle( glazingBottomEdge )
        UpVector = ghc.VectorXYZ(0,0,1).vector
        
        #Find if there are any shading objects and if so put them in a list
        HorizonShading = []
        
        HorizontalLine = ghc.LineSDL(ShadingOrigin, self.SurfaceNormal, _extents)
        VerticalLine = ghc.LineSDL(ShadingOrigin, UpVector, _extents)
        for shadingObj in _shadingGeom:
            if ghc.BrepXCurve(shadingObj, HorizontalLine).points != None:
                HorizonShading.append( shadingObj )
        
        ###########
        #Find any intersection Curves
        IntersectionSurface = ghc.SumSurface(HorizontalLine, VerticalLine)
        IntersectionCurve = []
        IntersectionPoints = []
        
        for shadingObj in HorizonShading:
            if ghc.BrepXBrep(shadingObj, IntersectionSurface).curves != None:
                IntersectionCurve.append(ghc.BrepXBrep(shadingObj, IntersectionSurface))
        for pnt in IntersectionCurve:
            IntersectionPoints.append(ghc.ControlPoints(pnt).points)
        
        ###########
        #Run the Top Corner Finder if there are any intersecting objects...
        if len(IntersectionPoints) != 0:
            #Find the top/closets point for each of the objects that could possibly shade
            KeyPoints = []
            for pnt in IntersectionPoints:
                Rays = []
                Angles = []
                if pnt:
                    for k in range(len(pnt)):
                        Rays.append(ghc.Vector2Pt(ShadingOrigin,pnt[k], False).vector)
                        Angles.append(ghc.Angle(self.SurfaceNormal , Rays[k]).angle)
                    KeyPoints.append(pnt[Angles.index(max(Angles))])
        
            #Find the relevant highest / closest point
            Rays = []
            Angles = []
            for i in range(len(KeyPoints)):
                Rays.append(ghc.Vector2Pt(self.SurfaceNormal, KeyPoints[i], False).vector)
                Angles.append(ghc.Angle(self.SurfaceNormal, Rays[i]).angle)
            KeyPoint = KeyPoints[Angles.index(max(Angles))]
        
            #use the point it finds to deliver the Height and Distance for the PHPP Shading Calculator
            h_hori = KeyPoint.Z - ShadingOrigin.Z #Vertical distance
            Hypot = ghc.Length(ghc.Line(ShadingOrigin, KeyPoint))
            d_hori = math.sqrt(Hypot**2 - h_hori**2)
            CheckLine = ghc.Line(ShadingOrigin, KeyPoint)
        else:
            h_hori = 0
            d_hori = 0
            CheckLine = HorizontalLine
        
        return h_hori, d_hori, CheckLine
    
    def findOverhangShading(self, _shadingGeom, _extents=99):
        # Figure out the glass surface (inset a bit) and then
        # find the origin point for all the subsequent shading calcs (top, middle)
        glzgCenter = ghc.Area(self.GlazingSrfc).centroid
        glazingEdges = self.getEdgesInOrder(self.GlazingSrfc)
        glazingTopEdge = glazingEdges[2]
        ShadingOrigin = ghc.CurveMiddle(glazingTopEdge)
        
        # In order to also work for windows which are not vertical, find the 
        # 'direction' from the glazing origin and the top/middle ege point
        UpVector = ghc.Vector2Pt(glzgCenter, ShadingOrigin, True).vector
        
        #-----------------------------------------------------------------------
        # First, need to filter the scene to find the objects that are 'above'
        # the window. Create a 'test plane' that is 99m tall and 2m deep, test if
        # any objects intersect that plane. If so, add them to the set of things
        # test in the next step
        edge1 = ghc.LineSDL(ShadingOrigin, UpVector, 99)
        edge2 = ghc.LineSDL(ShadingOrigin, self.getSurfaceNormal(self.GlazingSrfc), 2)
        intersectionTestPlane = ghc.SumSurface(edge1, edge2)
        
        OverhangShadingObjs = [x for x in _shadingGeom 
                        if ghc.BrepXBrep(intersectionTestPlane, x).curves != None]
        
        #-----------------------------------------------------------------------
        # Using the filtered set of shading objects, find the 'edges' of shading 
        # geom and then decide where the maximums shading point is
        # Create a new 'test' plane coming off the origin (99m in both directions this time).
        # Test to find any intersection shading objs and all their crvs/points with this plane
        HorizontalLine = ghc.LineSDL(ShadingOrigin, self.SurfaceNormal, _extents)
        VerticalLine = ghc.LineSDL(ShadingOrigin, UpVector, _extents)
        
        IntersectionSurface = ghc.SumSurface(HorizontalLine, VerticalLine)
        IntersectionCurves = [ghc.BrepXBrep(obj, IntersectionSurface).curves 
                                for obj in OverhangShadingObjs
                                if ghc.BrepXBrep(obj, IntersectionSurface).curves != None]
        IntersectionPoints = []
        for crv in IntersectionCurves:
            for pt in ghc.ControlPoints(crv).points:
                IntersectionPoints.append(pt)
        
        #If there are any intersection Points found, choose the right one to use to calc shading....
        if len(IntersectionPoints) != 0:
            #Find the top/closets point for each of the objects that could possibly shade
            Rays = []
            Angles = []
            for i, pt in enumerate(IntersectionPoints):
                if pt == None:
                    continue
                Rays.append(ghc.Vector2Pt(ShadingOrigin, pt, False).vector)
                Angles.append(ghc.Angle(self.SurfaceNormal , Rays[i]).angle)
            KeyPoint = IntersectionPoints[Angles.index(min(Angles))]
            
            #use the point it finds to deliver the Height and Distance for the PHPP Shading Calculator
            d_over = KeyPoint.Z - ShadingOrigin.Z #Vertical distance
            Hypot = ghc.Length(ghc.Line(ShadingOrigin, KeyPoint))
            o_over = math.sqrt(Hypot**2 - d_over**2)
            CheckLine = ghc.Line(ShadingOrigin, KeyPoint)
        else:
            d_over = 0
            o_over = 0
            CheckLine = VerticalLine
        
        return d_over, o_over, CheckLine
    
    def CalcRevealDims(self, RevealShaderObjs_input, SideIntersectionSurface, Side_OriginPt, Side_Direction):
        #Test shading objects for their edge points
        Side_IntersectionCurve = []
        Side_IntersectionPoints = []
        for i in range(len(RevealShaderObjs_input)): #This is the list of shading objects to filter
            if ghc.BrepXBrep(RevealShaderObjs_input[i], SideIntersectionSurface).curves != None:
                Side_IntersectionCurve.append(ghc.BrepXBrep(RevealShaderObjs_input[i], SideIntersectionSurface).curves)
        for i in range(len(Side_IntersectionCurve)):
            for k in range(len(ghc.ControlPoints(Side_IntersectionCurve[i]).points)):
                Side_IntersectionPoints.append(ghc.ControlPoints(Side_IntersectionCurve[i]).points[k])
        
        #Find the top/closets point for each of the objects that could possibly shade
        Side_KeyPoints = []
        Side_Rays = []
        Side_Angles = []
        for i in range(len(Side_IntersectionPoints)):
            if Side_OriginPt != Side_IntersectionPoints[i]:
                Ray = ghc.Vector2Pt(Side_OriginPt, Side_IntersectionPoints[i], False).vector
                Angle = math.degrees(ghc.Angle(self.SurfaceNormal, Ray).angle)
                if  Angle < 89.9:
                    Side_Rays.append(Ray)
                    Side_Angles.append(float(Angle))
                    Side_KeyPoints.append(Side_IntersectionPoints[i])
        Side_KeyPoint = Side_KeyPoints[Side_Angles.index(min(Side_Angles))]
        Side_KeyRay = Side_Rays[Side_Angles.index(min(Side_Angles))]
        
        #use the Key point found to calculte the Distances for the PHPP Shading Calculator
        Side_Hypot = ghc.Length(ghc.Line(Side_OriginPt, Side_KeyPoint))
        Deg = (ghc.Angle(Side_Direction, Side_KeyRay).angle) #note this is in Radians
        Side_o_reveal =  math.sin(Deg) * Side_Hypot
        Side_d_reveal = math.sqrt(Side_Hypot**2 - Side_o_reveal**2)
        Side_CheckLine = ghc.Line(Side_OriginPt, Side_KeyPoint)
        
        return [Side_o_reveal, Side_d_reveal, Side_CheckLine]
    
    def findRevealShading(self, _shadingGeom, _extents=99):
        
        WinCenter = ghc.Area(self.GlazingSrfc).centroid
        edges = self.getEdgesInOrder(self.GlazingSrfc)
        
        #Create the Intersection Surface for each side
        Side1_OriginPt = ghc.CurveMiddle( edges[1] )
        Side1_NormalLine = ghc.LineSDL(Side1_OriginPt, self.SurfaceNormal, _extents)
        Side1_Direction = ghc.Vector2Pt(WinCenter, Side1_OriginPt, False).vector
        Side1_HorizLine = ghc.LineSDL(Side1_OriginPt, Side1_Direction, _extents)
        Side1_IntersectionSurface = ghc.SumSurface(Side1_NormalLine, Side1_HorizLine)
        
        #Side2_OriginPt = SideMidPoints[1] #ghc.CurveMiddle(self.Edge_Left)
        Side2_OriginPt = ghc.CurveMiddle( edges[3] )
        Side2_NormalLine = ghc.LineSDL(Side2_OriginPt, self.SurfaceNormal, _extents)
        Side2_Direction = ghc.Vector2Pt(WinCenter, Side2_OriginPt, False).vector
        Side2_HorizLine = ghc.LineSDL(Side2_OriginPt, Side2_Direction, _extents)
        Side2_IntersectionSurface = ghc.SumSurface(Side2_NormalLine, Side2_HorizLine)
        
        #Find any Shader Objects and put them all into a list
        Side1_RevealShaderObjs = []
        testStartPt = ghc.Move(WinCenter, ghc.Amplitude(self.SurfaceNormal, 0.1)).geometry #Offsets the test line just a bit
        Side1_TesterLine = ghc.LineSDL(testStartPt, Side1_Direction, _extents) #extend a line off to side 1
        for i in range(len(_shadingGeom)):
            if ghc.BrepXCurve(_shadingGeom[i],Side1_TesterLine).points != None:
                Side1_RevealShaderObjs.append(_shadingGeom[i])
        
        Side2_RevealShaderObjs = []
        Side2_TesterLine = ghc.LineSDL(testStartPt, Side2_Direction, _extents) #extend a line off to side 2
        for i in range(len(_shadingGeom)):
            if ghc.BrepXCurve(_shadingGeom[i],Side2_TesterLine).points != None:
                Side2_RevealShaderObjs.append(_shadingGeom[i])
        
        NumShadedSides = 0
        if len(Side1_RevealShaderObjs) != 0:
            Side1_o_reveal = self.CalcRevealDims(Side1_RevealShaderObjs, Side1_IntersectionSurface, Side1_OriginPt, Side1_Direction)[0]
            Side1_d_reveal = self.CalcRevealDims(Side1_RevealShaderObjs, Side1_IntersectionSurface, Side1_OriginPt, Side1_Direction)[1]
            Side1_CheckLine = self.CalcRevealDims(Side1_RevealShaderObjs, Side1_IntersectionSurface, Side1_OriginPt, Side1_Direction)[2]
            NumShadedSides = NumShadedSides + 1
        else:
            Side1_o_reveal =  0
            Side1_d_reveal = 0
            Side1_CheckLine = Side1_HorizLine
        
        if len(Side2_RevealShaderObjs) != 0:
            Side2_o_reveal = self.CalcRevealDims(Side2_RevealShaderObjs, Side2_IntersectionSurface, Side2_OriginPt, Side2_Direction)[0]
            Side2_d_reveal = self.CalcRevealDims(Side2_RevealShaderObjs, Side2_IntersectionSurface, Side2_OriginPt, Side2_Direction)[1]
            Side2_CheckLine = self.CalcRevealDims(Side2_RevealShaderObjs, Side2_IntersectionSurface, Side2_OriginPt, Side2_Direction)[2]
            NumShadedSides = NumShadedSides + 1
        else:
            Side2_o_reveal =  0
            Side2_d_reveal = 0
            Side2_CheckLine = Side2_HorizLine
        
        o_reveal = (Side1_o_reveal + Side2_o_reveal )/ max(1,NumShadedSides)
        d_reveal = (Side1_d_reveal + Side2_d_reveal )/ max(1,NumShadedSides)
        
        return o_reveal, d_reveal, Side1_CheckLine, Side2_CheckLine
    
    def getEdgesInOrder(self, _srfc):
        # Sort the Edges using the Degree about center as the Key
        vectorList = self.getSrfcEdgeVectors(_srfc)
        edgeAngleDegrees = self.calcEdgeAngle(vectorList)
        srfcEdges_Unordered = ghc.DeconstructBrep(_srfc).edges
        srfcEdges_Ordered = ghc.SortList( edgeAngleDegrees, srfcEdges_Unordered).values_a
        
        return srfcEdges_Ordered
    
    def getWindowRevealGeom(self):
        # Create the Window Reveal (side) Geometry
        
        if 'Reveal_Bottom' not in self.__dict__:
            self.createWindowReveals()
        
        return [self.Reveal_Bottom, self.Reveal_Left, self.Reveal_Top, self.Reveal_Right]
    
    def createWindowReveals(self, _orientation=-1):
        
        # Build the Window Reveal Geometry
        self.Reveal_Bottom = self.cleanExtrude( self.Edge_Bottom, self.SurfaceNormal*_orientation, self.InstallDepth, self.Installs.Inst_B)
        self.Reveal_Left = self.cleanExtrude( self.Edge_Left, self.SurfaceNormal*_orientation, self.InstallDepth, self.Installs.Inst_L)
        self.Reveal_Top = self.cleanExtrude( self.Edge_Top, self.SurfaceNormal*_orientation, self.InstallDepth, self.Installs.Inst_T)
        self.Reveal_Right = self.cleanExtrude( self.Edge_Right, self.SurfaceNormal*_orientation, self.InstallDepth, self.Installs.Inst_R)
    
    def getWindowAlignedPlane(self, _srfc):
        """Finds an Aligned Plane for Surface input
        
        Note, will try and correct to make sure the aligned plane's Y-Axis aligns 
        to the surface and goes 'up' (world Z) if it can.
        
        Arguments:
            None
        Returns:
            srfcPlane: A single Plane object, aligned to the surface
        """
        
        # Get the UV info for the surface
        srfcPlane = rs.SurfaceFrame(_srfc, [0.5, 0.5])
        centroid = ghc.Area(_srfc).centroid
        uVector = srfcPlane.XAxis
        vVector = srfcPlane.YAxis
        
        # Create a Plane aligned to the UV of the srfc
        lineU = ghc.LineSDL(centroid, uVector, 1)
        lineV = ghc.LineSDL(centroid, vVector, 1)
        srfcPlane = ghc.Line_Line(lineU, lineV)
        
        # Try and make sure its pointing the right directions
        if abs(round(srfcPlane.XAxis.Z, 2)) != 0:
            srfcPlane =  ghc.RotatePlane(srfcPlane, ghc.Radians(90))
        if round(srfcPlane.YAxis.Z, 2) < 0:
            srfcPlane =  ghc.RotatePlane(srfcPlane, ghc.Radians(180))
        
        return srfcPlane
    
    def getSrfcEdgeVectors(self, _srfc):
        """ Finds a Vector from center of surface to mid-point on each edge
        
        Arguments:
            None
        Returns:
            edgeVectors: (List) Vector3D for mid-point on each edge
        """
        
        # Project the Surface edges onto the World Plane
        srfcPlane = self.getWindowAlignedPlane(_srfc)
        
        worldOrigin = Rhino.Geometry.Point3d(0,0,0)
        worldXYPlane = ghc.XYPlane(worldOrigin)
        geomAtWorldZero = ghc.Orient(_srfc, srfcPlane, worldXYPlane).geometry
        edges = ghc.DeconstructBrep(geomAtWorldZero).edges
        
        # Find the mid-point for each edge and create a vector to that midpoint
        crvMidPoints = [ ghc.CurveMiddle(edge) for edge in edges ]
        edgeVectors = [ ghc.Vector2Pt(midPt, worldOrigin, False).vector for midPt in crvMidPoints ]
        
        return edgeVectors
    
    def calcEdgeAngle(self, _vectorList):
        """Takes in a list of vectors. Calculates the Vector angle about a center
        
        Assumes that the center is (0,0,0). Will calculate around 360 degrees 
        and return values in degrees not radians.
        
        Arguments:
            _vectorList: (list) Vectors to the surface's edges
        Returns:
            vectorAngles: (List) Float values of Degrees for each Vector input
        """
        
        vectorAngles = []
        
        refAngle = ghc.UnitY(1)
        x2 = refAngle.X
        y2 = refAngle.Y
        
        for vector in _vectorList:
            x1 = vector.X
            y1 = vector.Y
            
            # Calc the angle between
            angle = math.atan2(y2, x2) - math.atan2(y1, x1)
            angle = angle * 360 / (2 * math.pi)
            
            if angle < 0:
                angle = angle + 360
            
            angle = round(angle, 0)
            
            if angle >359.9 or angle < 0.001:
                angle = 0
            
            vectorAngles.append(angle)
        
        return vectorAngles
    
    def getSurfaceNormal(self, _srfc):
        centroid = ghc.Area(_srfc).centroid
        srfcNormal = rs.SurfaceNormal(_srfc, centroid)
        
        return srfcNormal
    
    def __unicode__(self):
        return u'A PHPP-Style Window Object: < {self.Name} >'.format(self=self)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
       return "{}( _nm={!r},\n_glassType={!r},\n"\
               "_frameType={!r},\n_installs={!r},\n"\
               "_geom={!r},\n_varType={!r},\n_instDepth={!r} )".format(
               self.__class__.__name__,
               self.Name,
               self.Type_Glass,
               self.Type_Frame,
               self.Installs,
               self.Geometry,
               self.Type_Variant,
               self.InstallDepth)

class PHPP_Window_Shading_Calculator():
    global __e
    __e = 2.71828182845904
    
    def __init__(self, _lat=40):
        self.Latitude = _lat
    
    def Winter_HorizShadingFactor(self, h_hori, d_hori, Tilt, Azimuth, GlazingHeight):
        #Set up the Horiz Constants
        hor_m1 = [0.011953348, 0.011953348, 0.001476536, 0.001476536, -0.001307563, -0.001307563]
        hor_b1 = [-0.261997475, -0.261997475, 0.490848602, 0.490848602, 0.883715765, 0.883715765]
        hor_Deg1 = 90
        hor_m2 = [0.000640228, 0.000640228, 0.006396405, 0.006396405, 0.000563535, 0.000563535]
        hor_b2 = [0.099820235, 0.099820235, 0.276192963, 0.276192963, 0.470439231, 0.470439231]
        hor_Deg2 = 90
        
        perp_m1 = [-0.001532016, -0.001532016, 0.001081800, 0.001081800, 0.001877231, 0.001877231]
        perp_b1 = [0.151057660, 0.151057660, 0.765554663, 0.765554663, 0.534441295, 0.534441295]
        perp_Deg1 = 90
        perp_m2 = [0.000000250, 0.000000250, -0.009571308, -0.009571308, 0.005044849, 0.005044849]
        perp_b2 = [0.662706638, 0.662706638, -0.686205976, -0.686205976, -1.651987505, -1.651987505]
        perp_Deg2 = 90
        
        # Clean inputs
        if d_hori == 0:
            hh_dh = 1
        else:
            hh_dh = (h_hori / d_hori)
        
        Azimuth_down = (int(Azimuth % 360) / 90) * 90
        V_Factor_hori = 1
        
        #Set up the Horizontal Factors
        if self.Latitude <= hor_Deg1:
            hor_S_r = hor_m1[0] * self.Latitude + hor_b1[0]
        else:
            hor_S_r = hor_m1[1] * self.Latitude + hor_b1[1]
        if self.Latitude <= hor_Deg1:
            hor_OW_r = hor_m1[2] * self.Latitude + hor_b1[2]
        else:
            hor_OW_r = hor_m1[3] * self.Latitude + hor_b1[3]
        if self.Latitude <= hor_Deg1:
            hor_N_r = hor_m1[4] * self.Latitude + hor_b1[4]
        else:
            hor_N_r = hor_m1[5] * self.Latitude + hor_b1[5]
        
        if self.Latitude <= hor_Deg2:
            hor_S_a = hor_m2[0] * self.Latitude**2 + hor_b2[0]
        else:
            hor_S_a = hor_m2[1]* self.Latitude + hor_b2[1]
        if self.Latitude <= hor_Deg2:
            hor_OW_a = hor_m2[2] * self.Latitude + hor_b2[2]
        else:
            hor_OW_a = hor_m2[3]* self.Latitude + hor_b2[3]
        if self.Latitude <= hor_Deg2:
            hor_N_a = hor_m2[4] * self.Latitude + hor_b2[4]
        else:
            hor_N_a = hor_m2[5]* self.Latitude + hor_b2[5]
        
        #Set up the Perpendicular Factors
        if self.Latitude <= perp_Deg1:
            perp_S_r = perp_m1[0] * self.Latitude + perp_b1[0]
        else:
            perp_S_r = perp_m1[1] * self.Latitude + perp_b1[1]
        if self.Latitude <= perp_Deg1:
            perp_OW_r = perp_m1[2] * self.Latitude + perp_b1[2]
        else:
            perp_OW_r = perp_m1[3] * self.Latitude + perp_b1[3]
        if self.Latitude <= perp_Deg1:
            perp_N_r = perp_m1[4] * self.Latitude + perp_b1[4]
        else:
            perp_N_r = perp_m1[5] * self.Latitude + perp_b1[5]
        
        if self.Latitude <= perp_Deg2:
            perp_S_a = perp_m2[0] * self.Latitude**4 + perp_b2[0]
        else:
            perp_S_a = perp_m2[1]* self.Latitude + perp_b2[1]
        if self.Latitude <= perp_Deg2:
            perp_OW_a = perp_m2[2] * self.Latitude + perp_b2[2]
        else:
            perp_OW_a = perp_m2[3]* self.Latitude + perp_b2[3]
        if self.Latitude <= perp_Deg2:
            perp_N_a = perp_m2[4] * self.Latitude + perp_b2[4]
        else:
            perp_N_a = perp_m2[5]* self.Latitude + perp_b2[5]
        
        #hor1 Factor Calcs
        if Azimuth_down == 0:
            hor1 = hor_N_r + (1 - hor_N_r) / (1 + hh_dh**2)**hor_N_a
        elif Azimuth_down == 180:
            hor1 = hor_S_r + (1 - hor_S_r) / (1 + hh_dh**2)**hor_S_a
        else:
            hor1 = hor_OW_r + (1 - hor_OW_r) / (1 + hh_dh**2)**hor_OW_a
        
        #hor2 Factor Calcs
        if Azimuth_down == 270:
            hor2 = hor_N_r + (1 - hor_N_r) / (1 + hh_dh**2)**hor_N_a
        elif Azimuth_down == 90:
            hor2 = hor_S_r + (1 - hor_S_r) / (1 + hh_dh**2)**hor_S_a
        else:
            hor2 = hor_OW_r + (1 - hor_OW_r) / (1 + hh_dh**2)**hor_OW_a
        
        #Calc Senk1 Factor
        if Azimuth_down == 0:
            senk1 = perp_N_r * __e**(perp_N_a * hh_dh) + 1 - perp_N_r
        elif Azimuth_down == 180:
            senk1 = perp_S_r + (1 - perp_S_r) / (1 + hh_dh**2)**perp_S_a
        else:
            senk1 = perp_OW_r * __e**(perp_OW_a * hh_dh) + 1 - perp_OW_r
        
        #Calc Senk2 Factor
        if Azimuth_down == 270:
            senk2 = perp_N_r * __e**(perp_N_a * hh_dh) + 1 - perp_N_r
        elif Azimuth_down == 90:
            senk2 = perp_S_r + (1 - perp_S_r) / (1 + hh_dh**2)**perp_S_a
        else:
            senk2 = perp_OW_r * __e**(perp_OW_a * hh_dh) + 1 - perp_OW_r
        
        #Calc main Factors
        ipol_hor = hor1 + 1/2 * (hor2 - hor1) * (1 - math.cos(2 * (Azimuth - Azimuth_down)*(math.pi/180)))
        ipol_senk = senk1 + 1/2 * (senk2 - senk1) * (1 - math.cos(2 * (Azimuth - Azimuth_down) * (math.pi/180)))
        
        if math.sin(math.radians(Tilt)) != 0:
            x = 1 - h_hori / GlazingHeight /abs(math.sin(math.radians(Tilt)))
        else:
            x = 0
        V_Factor_hori = max(ipol_hor + 1/2 *(ipol_senk - ipol_hor) * (1-math.cos(2 * Tilt * (math.pi/180))), x)
        
        return V_Factor_hori
    
    def Winter_RevealShadingFactor(self, o_reveal, d_reveal, Tilt, Azimuth, GlazingWidth):
        #Calc the first values
        Ti = o_reveal /(0.5 * GlazingWidth + d_reveal)
        Azimuth_down = (int(Azimuth % 360) / 90) * 90
        
        #Set up the Horiz Constants
        hor_m1 = [-0.000548191, -0.000548191, -0.000559231, -0.000559231, -0.000548195, -0.000548195]
        hor_b1 = [0.813773754, 0.813773754, 0.863835221, 0.863835221, 0.813773898, 0.813773898]
        hor_Deg1 = 90
        hor_m2 = [-0.0038284, -0.0038284, -0.012573537, -0.012573537, -0.003828399, -0.003828399]
        hor_b2 = [-0.262761285, -0.262761285, -0.13186139, -0.13186139, -0.262761072, -0.262761072]
        hor_Deg2 = 90
        
        perp_m1 = [0.001198624, 0.001198624, 0.000748647, 0.000748647, 0.000212172, 0.000212173]
        perp_b1 = [0.792233903, 0.792233903, 0.771014921, 0.771014921, 0.605894803, 0.605894803]
        perp_Deg1 = 90
        perp_m2 = [0.005799242, 0.005799242, -0.005084283, -0.005084283, 0.00009799056, 0.00009799056]
        perp_b2 = [-0.547245304, -0.547245304, -0.293059757, -0.293059757, -0.500894234, -0.500894234]
        perp_Deg2 = 90
        
        #Set up the Horizontal Factors
        if self.Latitude <= hor_Deg1:
            hor_S_r = hor_m1[0] * self.Latitude + hor_b1[0]
        else:
            hor_S_r = hor_m1[1] * self.Latitude + hor_b1[1]
        if self.Latitude <= hor_Deg1:
            hor_OW_r = hor_m1[2] * self.Latitude + hor_b1[2]
        else:
            hor_OW_r = hor_m1[3] * self.Latitude + hor_b1[3]
        if self.Latitude <= hor_Deg1:
            hor_N_r = hor_m1[4] * self.Latitude + hor_b1[4]
        else:
            hor_N_r = hor_m1[5] * self.Latitude + hor_b1[5]
        
        if self.Latitude <= hor_Deg2:
            hor_S_a = hor_m2[0] * self.Latitude + hor_b2[0]
        else:
            hor_S_a = hor_m2[1] * self.Latitude + hor_b2[1]
        if self.Latitude <= hor_Deg2:
            hor_OW_a = hor_m2[2] * self.Latitude + hor_b2[2]
        else:
            hor_OW_a = hor_m2[3] * self.Latitude + hor_b2[3]
        if self.Latitude <= hor_Deg2:
            hor_N_a = hor_m2[4] * self.Latitude + hor_b2[4]
        else:
            hor_N_a = hor_m2[5] * self.Latitude + hor_b2[5]
        
        #Set up the Perpendicular Factors
        if self.Latitude <= perp_Deg1:
            perp_S_r = perp_m1[0] * self.Latitude + perp_b1[0]
        else:
            perp_S_r = perp_m1[1] * self.Latitude + perp_b1[1]
        if self.Latitude <= perp_Deg1:
            perp_OW_r = perp_m1[2] * self.Latitude + perp_b1[2]
        else:
            perp_OW_r = perp_m1[3] * self.Latitude + perp_b1[3]
        if self.Latitude <= perp_Deg1:
            perp_N_r = perp_m1[4] * self.Latitude + perp_b1[4]
        else:
            perp_N_r = perp_m1[5] * self.Latitude + perp_b1[5]
        
        if self.Latitude <= perp_Deg2:
            perp_S_a = perp_m2[0] * self.Latitude + perp_b2[0]
        else:
            perp_S_a = perp_m2[1] * self.Latitude + perp_b2[1]
        if self.Latitude <= perp_Deg2:
            perp_OW_a = perp_m2[2] * self.Latitude + perp_b2[2]
        else:
            perp_OW_a = perp_m2[3] * self.Latitude + perp_b2[3]
        if self.Latitude <= perp_Deg2:
            perp_N_a = perp_m2[4] * self.Latitude + perp_b2[4]
        else:
            perp_N_a = perp_m2[5] * self.Latitude + perp_b2[5]
        
        #Calc hori1
        if Azimuth_down == 0:
            hor1 = hor_N_r * __e**(hor_N_a * Ti) + 1 - hor_N_r #North
        elif Azimuth_down == 180:
            hor1 = hor_S_r * __e**(hor_S_a * Ti) + 1 - hor_S_r #South
        else:
            hor1 = hor_OW_r * __e**(hor_OW_a * Ti) + 1 - hor_OW_r #East / West
        
        #Calc hori2
        if Azimuth_down == 270:
            hor2 = hor_N_r * __e**(hor_N_a * Ti) + 1 - hor_N_r #North
        elif Azimuth_down == 90:
            hor2 = hor_S_r * __e**(hor_S_a * Ti) + 1 - hor_S_r #South
        else:
            hor2 = hor_OW_r * __e**(hor_OW_a * Ti) + 1 - hor_OW_r #East / West
        
        #Calc senk1
        if Azimuth_down == 0:
            senk1 = perp_N_r * __e**(perp_N_a * Ti) + 1 - perp_N_r #North
        elif Azimuth_down == 180:
            senk1 = perp_S_r * __e**(perp_S_a * Ti) + 1 - perp_S_r #South
        else:
            senk1 = perp_OW_r * __e**(perp_OW_a * Ti) + 1 - perp_OW_r #East / West
        
        #Calc senk2
        if Azimuth_down == 270:
            senk2 = perp_N_r * __e**(perp_N_a * Ti) + 1 - perp_N_r #North
        elif Azimuth_down == 90:
            senk2 = perp_S_r * __e**(perp_S_a * Ti) + 1 - perp_S_r #South
        else:
            senk2 = perp_OW_r * __e**(perp_OW_a * Ti) + 1 - perp_OW_r #East / West
        
        #Calc the shading Factor
        ipol_hor = hor1 +1/2* (hor1 - hor2)*(1 - math.cos(2*(Azimuth - Azimuth_down) * (math.pi/180)))
        ipol_senk = senk1 +1/2* (senk2 - senk1)*(1 - math.cos(2*(Azimuth - Azimuth_down) * (math.pi/180)))
        V_Factor_Reveal = ipol_hor + 1/2 * (ipol_senk - ipol_hor) * (1 - math.cos( 2* Tilt * (math.pi/180)))
        
        return V_Factor_Reveal
    
    def Winter_OverhangShadingFactor(self, o_over, d_over, Tilt, Azimuth, GlazingHeight):
        #Set Up Input Values
        Azimuth_down = (int(Azimuth % 360) / 90) * 90
        Ti = o_over / (0.5 * GlazingHeight + d_over)
        
        #Set up the Horiz Constants
        hor_m1 = [0.001283722, 0.001283722, 0.000100694, 0.000100694, -0.00122305, -0.00122305]
        hor_b1 = [0.145829537, 0.145829537, 0.450066972, 0.450066972, 0.808786341, 0.808786341]
        hor_Deg1 = [90, 90, 90]
        hor_m2 = [-0.0000150610, -0.0000150610, -0.003128425, -0.003128425, -0.013504735, -0.013504735]
        hor_b2 = [-0.418020571, -0.418020571, -0.267285533, -0.267285533, -0.097018948, -0.097018948]
        hor_Deg2 = [90, 90, 90]
        
        perp_m1 = [0.002770419, 0.002770419, 0.00061612, 0.00061612, -0.010354458, 0.004146509]
        perp_b1 = [0.837446196, 0.837446196, 0.818195665, 0.818195665, 0.728350162, 0.515381429]
        perp_Deg1 = [90, 90, 15]
        perp_m2 = [0.009666083, 0.009666084, 0.002385135, 0.002385135, -0.000246449, -0.000246449]
        perp_b2 = [-0.620970425, -0.620970425, -0.449635241, -0.449635241, -0.393434489, -0.393434489]
        perp_Deg2 = [90, 90, 90]
        
        #Set up the Horizontal Factors
        if self.Latitude <= hor_Deg1[0]:
            hor_S_r = hor_m1[0] * self.Latitude + hor_b1[0]
        else:
            hor_S_r = hor_m1[1] * self.Latitude + hor_b1[1]
        if self.Latitude <= hor_Deg1[1]:
            hor_OW_r = hor_m1[2] * self.Latitude + hor_b1[2]
        else:
            hor_OW_r = hor_m1[3] * self.Latitude + hor_b1[3]
        if self.Latitude <= hor_Deg1[2]:
            hor_N_r = hor_m1[4] * self.Latitude + hor_b1[4]
        else:
            hor_N_r = hor_m1[5] * self.Latitude + hor_b1[5]
        
        if self.Latitude <= hor_Deg2[0]:
            hor_S_a = hor_m2[0] * self.Latitude + hor_b2[0]
        else:
            hor_S_a = hor_m2[1] * self.Latitude + hor_b2[1]
        if self.Latitude <= hor_Deg2[1]:
            hor_OW_a = hor_m2[2] * self.Latitude + hor_b2[2]
        else:
            hor_OW_a = hor_m2[3] * self.Latitude + hor_b2[3]
        if self.Latitude <= hor_Deg2[2]:
            hor_N_a = hor_m2[4] * self.Latitude + hor_b2[4]
        else:
            hor_N_a = hor_m2[5] * self.Latitude + hor_b2[5]
        
        #Set up the Perpendicular Factors
        if self.Latitude <= perp_Deg1[0]:
            perp_S_r = perp_m1[0] * self.Latitude + perp_b1[0]
        else:
            perp_S_r = perp_m1[1] * self.Latitude + perp_b1[1]
        if self.Latitude <= perp_Deg1[1]:
            perp_OW_r = perp_m1[2] * self.Latitude + perp_b1[2]
        else:
            perp_OW_r = perp_m1[3] * self.Latitude + perp_b1[3]
        if self.Latitude <= perp_Deg1[2]:
            perp_N_r = perp_m1[4] * self.Latitude + perp_b1[4]
        else:
            perp_N_r = perp_m1[5] * self.Latitude + perp_b1[5]
        
        if self.Latitude <= perp_Deg2[0]:
            perp_S_a = perp_m2[0] * self.Latitude + perp_b2[0]
        else:
            perp_S_a = perp_m2[1] * self.Latitude + perp_b2[1]
        if self.Latitude <= perp_Deg2[1]:
            perp_OW_a = perp_m2[2] * self.Latitude + perp_b2[2]
        else:
            perp_OW_a = perp_m2[3] * self.Latitude + perp_b2[3]
        if self.Latitude <= perp_Deg2[2]:
            perp_N_a = perp_m2[4] * self.Latitude + perp_b2[4]
        else:
            perp_N_a = perp_m2[5] * self.Latitude + perp_b2[5]
        
        #Calc hori1
        if Azimuth_down == 0:
            hor1 = hor_N_r * __e**(hor_N_a * Ti) + 1 - hor_N_r #North
        elif Azimuth_down == 180:
            hor1 = hor_S_r * __e**(hor_S_a * Ti) + 1 - hor_S_r #South
        else:
            hor1 = hor_OW_r * __e**(hor_OW_a * Ti) + 1 - hor_OW_r #East / West
        
        #Calc hori2
        if Azimuth_down == 270:
            hor2 = hor_N_r * __e**(hor_N_a * Ti) + 1 - hor_N_r #North
        elif Azimuth_down == 90:
            hor2 = hor_S_r * __e**(hor_S_a * Ti) + 1 - hor_S_r #South
        else:
            hor2 = hor_OW_r * __e**(hor_OW_a * Ti) + 1 - hor_OW_r #East / West
        
        #Calc senk1
        if Azimuth_down == 0:
            senk1 = perp_N_r * __e**(perp_N_a * Ti) + 1 - perp_N_r #North
        elif Azimuth_down == 180:
            senk1 = perp_S_r * __e**(perp_S_a * Ti) + 1 - perp_S_r #South
        else:
            senk1 = perp_OW_r * __e**(perp_OW_a * Ti) + 1 - perp_OW_r #East / West
        
        #Calc senk2
        if Azimuth_down == 270:
            senk2 = perp_N_r * __e**(perp_N_a * Ti) + 1 - perp_N_r #North
        elif Azimuth_down == 90:
            senk2 = perp_S_r * __e**(perp_S_a * Ti) + 1 - perp_S_r #South
        else:
            senk2 = perp_OW_r * __e**(perp_OW_a * Ti) + 1 - perp_OW_r #East / West
        
        #Calc the Shading Factors
        ipol_hor = hor1 + 1/2 * (hor2 - hor1) * (1-math.cos(2* (Azimuth - Azimuth_down) * (math.pi/180)))
        ipol_senk = senk1 + 1/2 * (senk2 - senk1) * (1-math.cos(2* (Azimuth - Azimuth_down) * (math.pi/180)))
        V_Factor_Overhang = ipol_hor + 1/2 * (ipol_senk - ipol_hor)*(1-math.cos(2* Tilt * (math.pi/180)))
        
        return V_Factor_Overhang
    
    def Summer_HorizShadingFactor(self, h_hori, d_hori, Tilt, Azimuth, GlazingHeight):
        #Clean Input Values
        Azimuth_down = (int(Azimuth % 360) / 90) * 90
        
        if d_hori == 0:
            hh_dh = 1
        else:
            hh_dh = (h_hori / d_hori)
       
       #Set up the Horiz Constants
        hor_m1 = [0.011087354, 0.011087354, -0.000357811, -0.000357811, -0.021768375, 0.001776364]
        hor_b1 = [0.245266666, 0.245266666, 0.439737536, 0.439737536, 0.813891468, 0.060265281]
        hor_Deg1 = [90, 90, 30, 90]
        hor_m2 = [0.019420642, -0.00671882, -0.001892317, -0.001892317, 0.007111836, -0.006513076]
        hor_b2 = [-0.363311546, 0.074358676, -0.251478828, -0.251478828, -0.183533648, 0.038607144]
        hor_Deg2 = [15, 90, 90, 15, 90]
        
        perp_m1 = [-0.005497456, 0.009245854, 0.000687231, 0.000687231, -0.012690832, 0.007982084]
        perp_b1 = [0.591562786, 0.379954102, 0.763980001, 0.763980001, 0.79326207, 0.169870363]
        perp_Deg1 = [15, 90, 90, 30, 90]
        perp_m2 = [0.034365419, -0.007993084, -0.003687613, -0.003687613, -0.041641853, 0.025080489]
        perp_b2 = [-1.497302397, -0.036651195, -0.714322989, -0.714322989, -0.309580707, -2.725745415]
        perp_Deg2 = [30, 90, 90, 38, 90]
        
        #Set up the Horizontal Factors
        if self.Latitude <= hor_Deg1[0]:
            hor_S_r = hor_m1[0] * self.Latitude + hor_b1[0]
        else:
            hor_S_r = hor_m1[1] * self.Latitude + hor_b1[1]
        if self.Latitude <= hor_Deg1[1]:
            hor_OW_r = hor_m1[2] * self.Latitude + hor_b1[2]
        else:
            hor_OW_r = hor_m1[3] * self.Latitude + hor_b1[3]
        if self.Latitude <= hor_Deg1[2]:
            hor_N_r = hor_m1[4] * self.Latitude + hor_b1[4]
        else:
            hor_N_r = hor_m1[5] * self.Latitude + hor_b1[5]
        
        if self.Latitude <= hor_Deg2[0]:
            hor_S_a = hor_m2[0] * self.Latitude + hor_b2[0]
        else:
            hor_S_a = hor_m2[1] * self.Latitude + hor_b2[1]
        if self.Latitude <= hor_Deg2[2]:
            hor_OW_a = hor_m2[2] * self.Latitude + hor_b2[2]
        else:
            hor_OW_a = hor_m2[3] * self.Latitude + hor_b2[3]
        if self.Latitude <= hor_Deg2[4]:
            hor_N_a = hor_m2[4] * self.Latitude + hor_b2[4]
        else:
            hor_N_a = hor_m2[5] * self.Latitude + hor_b2[5]
        
        #Set up the Perpendicular Factors
        if self.Latitude <= perp_Deg1[0]:
            perp_S_r = perp_m1[0] * self.Latitude + perp_b1[0]
        else:
            perp_S_r = perp_m1[1] * self.Latitude + perp_b1[1]
        if self.Latitude <= perp_Deg1[2]:
            perp_OW_r = perp_m1[2] * self.Latitude + perp_b1[2]
        else:
            perp_OW_r = perp_m1[3] * self.Latitude + perp_b1[3]
        if self.Latitude <= perp_Deg1[3]:
            perp_N_r = perp_m1[4] * abs(self.Latitude) + perp_b1[4]
        else:
            perp_N_r = perp_m1[5] * self.Latitude + perp_b1[5]
        
        if self.Latitude <= perp_Deg2[0]:
            perp_S_a = perp_m2[0] * self.Latitude + perp_b2[0]
        else:
            perp_S_a = perp_m2[1] * self.Latitude + perp_b2[1]
        if self.Latitude <= perp_Deg2[2]:
            perp_OW_a = perp_m2[2] * self.Latitude + perp_b2[2]
        else:
            perp_OW_a = perp_m2[3] * self.Latitude + perp_b2[3]
        if self.Latitude <= perp_Deg2[3]:
            perp_N_a = perp_m2[4] * abs(self.Latitude) + perp_b2[4]
        else:
            perp_N_a = perp_m2[5] * self.Latitude + perp_b2[5]
        
        #Calc hor1
        if Azimuth_down == 0:
            hor1 = hor_N_r * __e**(hor_N_a * hh_dh) + 1 - hor_N_r
        elif Azimuth_down == 180:
            hor1 = hor_S_r * __e**(hor_S_a * hh_dh) + 1 - hor_S_r
        else:
            hor1 = hor_OW_r * __e**(hor_OW_a * hh_dh) + 1 - hor_OW_r
        
        #Calc hor2
        if Azimuth_down == 270:
            hor2 = hor_N_r * __e**(hor_N_a * hh_dh) + 1 - hor_N_r
        elif Azimuth_down == 90:
            hor2 = hor_S_r * __e**(hor_S_a * hh_dh) + 1 - hor_S_r
        else:
            hor2 = hor_OW_r * __e**(hor_OW_a * hh_dh) + 1 - hor_OW_r
        
        #Calc senk1
        if Azimuth_down == 0:
            senk1 = perp_N_r * __e**(perp_N_a * hh_dh) + 1 - perp_N_r
        elif Azimuth_down == 180:
            senk1 = perp_S_r * __e**(perp_S_a * hh_dh) + 1 - perp_S_r
        else:
            senk1 = perp_OW_r * __e**(perp_OW_a * hh_dh) + 1 - perp_OW_r
        
        #Calc senk2
        if Azimuth_down == 270:
            senk2 = perp_N_r * __e**(perp_N_a * hh_dh) + 1 - perp_N_r
        elif Azimuth_down == 90:
            senk2 = perp_S_r * __e**(perp_S_a * hh_dh) + 1 - perp_S_r
        else:
            senk2 = perp_OW_r * __e**(perp_OW_a * hh_dh) + 1 - perp_OW_r
        
        #Calc main Factors
        ipol_hor = hor1 + 1/2 * (hor2 - hor1) * (1 - math.cos(2 * (Azimuth - Azimuth_down)*(math.pi/180)))
        ipol_senk = senk1 + 1/2 * (senk2 - senk1) * (1 - math.cos(2 * (Azimuth - Azimuth_down) * (math.pi/180)))
        
        if math.sin(math.radians(Tilt)) != 0:
            x = 1 - h_hori / GlazingHeight /abs(math.sin(math.radians(Tilt)))
        else:
            x = 0
        V_Factor_hori = max(ipol_hor + 1/2 *(ipol_senk - ipol_hor) * (1-math.cos(2 * Tilt * (math.pi/180))), x)
        
        return V_Factor_hori
    
    def Summer_RevealShadingFactor(self, o_reveal, d_reveal, Tilt, Azimuth, GlazingWidth):
        #Calc the first values
        Ti = o_reveal /(0.5 * GlazingWidth + d_reveal)
        Azimuth_down = (int(Azimuth % 360) / 90) * 90
        
        #Set up the Horiz Constants
        hor_m1 = [0.000104144, 0.000104144, 0.00015311, 0.00015311, 0.000104139, 0.000104139]
        hor_b1 = [0.787949282, 0.787949282, 0.877404098, 0.877404098, 0.787949573,0.787949573]
        hor_Deg1 = 90
        hor_m2 = [-0.001831421, -0.001831421, 0.007418363, -0.0053981, -0.001831426, -0.001831426]
        hor_b2 = [-0.213224793, -0.213224793, -0.163302446, 0.055931854, -0.213224648, -0.213224648]
        hor_Deg2 = [90, 15, 90, 90]
        
        perp_m1 = [0.005554356, 0.005554356, 0.0000715566, 0.0000715566, -0.003146017, -0.003146017]
        perp_b1 = [0.469559126, 0.469559126, 0.701581453, 0.701581453, 0.655929265, 0.655929265]
        perp_Deg1 = 90
        perp_m2 = [-0.007758874, 0.010281299, 0.004793887, -0.003376084, 0.002061171, 0.002061171]
        perp_b2 = [-0.374274227, -0.907643643, -0.214135787, -0.056284218, -0.682640388, -0.682640388]
        perp_Deg2 = [30, 90, 23, 90, 90]
        
        #Set up the Horizontal Factors
        if self.Latitude <= hor_Deg1:
            hor_S_r = hor_m1[0] * self.Latitude + hor_b1[0]
        else:
            hor_S_r = hor_m1[1] * self.Latitude + hor_b1[1]
        if self.Latitude <= hor_Deg1:
            hor_OW_r = hor_m1[2] * self.Latitude + hor_b1[2]
        else:
            hor_OW_r = hor_m1[3] * self.Latitude + hor_b1[3]
        if self.Latitude <= hor_Deg1:
            hor_N_r = hor_m1[4] * self.Latitude + hor_b1[4]
        else:
            hor_N_r = hor_m1[5] * self.Latitude + hor_b1[5]
        
        if self.Latitude <= hor_Deg2[0]:
            hor_S_a = hor_m2[0] * self.Latitude + hor_b2[0]
        else:
            hor_S_a = hor_m2[1] * self.Latitude + hor_b2[1]
        if self.Latitude <= hor_Deg2[1]:
            hor_OW_a = hor_m2[2] * self.Latitude + hor_b2[2]
        else:
            hor_OW_a = hor_m2[3] * self.Latitude + hor_b2[3]
        if self.Latitude <= hor_Deg2[3]:
            hor_N_a = hor_m2[4] * self.Latitude + hor_b2[4]
        else:
            hor_N_a = hor_m2[5] * self.Latitude + hor_b2[5]
        
        #Set up the Perpendicular Factors
        if self.Latitude <= perp_Deg1:
            perp_S_r = perp_m1[0] * self.Latitude + perp_b1[0]
        else:
            perp_S_r = perp_m1[1] * self.Latitude + perp_b1[1]
        if self.Latitude <= perp_Deg1:
            perp_OW_r = perp_m1[2] * self.Latitude + perp_b1[2]
        else:
            perp_OW_r = perp_m1[3] * self.Latitude + perp_b1[3]
        if self.Latitude <= perp_Deg1:
            perp_N_r = perp_m1[4] * self.Latitude + perp_b1[4]
        else:
            perp_N_r = perp_m1[5] * self.Latitude + perp_b1[5]
        
        if self.Latitude <= perp_Deg2[0]:
            perp_S_a = perp_m2[0] * self.Latitude + perp_b2[0]
        else:
            perp_S_a = perp_m2[1] * self.Latitude + perp_b2[1]
        if self.Latitude <= perp_Deg2[2]:
            perp_OW_a = perp_m2[2] * self.Latitude + perp_b2[2]
        else:
            perp_OW_a = perp_m2[3] * self.Latitude + perp_b2[3]
        if self.Latitude <= perp_Deg2[4]:
            perp_N_a = perp_m2[4] * self.Latitude + perp_b2[4]
        else:
            perp_N_a = perp_m2[5] * self.Latitude + perp_b2[5]
        
        #Calc hori1
        if Azimuth_down == 0:
            hor1 = hor_N_r * __e**(hor_N_a * Ti) + 1 - hor_N_r #North
        elif Azimuth_down == 180:
            hor1 = hor_S_r * __e**(hor_S_a * Ti) + 1 - hor_S_r #South
        else:
            hor1 = hor_OW_r * __e**(hor_OW_a * Ti) + 1 - hor_OW_r #East / West
        
        #Calc hori2
        if Azimuth_down == 270:
            hor2 = hor_N_r * __e**(hor_N_a * Ti) + 1 - hor_N_r #North
        elif Azimuth_down == 90:
            hor2 = hor_S_r * __e**(hor_S_a * Ti) + 1 - hor_S_r #South
        else:
            hor2 = hor_OW_r * __e**(hor_OW_a * Ti) + 1 - hor_OW_r #East / West
        
        #Calc senk1
        if Azimuth_down == 0:
            senk1 = perp_N_r * __e**(perp_N_a * Ti) + 1 - perp_N_r #North
        elif Azimuth_down == 180:
            senk1 = perp_S_r * __e**(perp_S_a * Ti) + 1 - perp_S_r #South
        else:
            senk1 = perp_OW_r * __e**(perp_OW_a * Ti) + 1 - perp_OW_r #East / West
        
        #Calc senk2
        if Azimuth_down == 270:
            senk2 = perp_N_r * __e**(perp_N_a * Ti) + 1 - perp_N_r #North
        elif Azimuth_down == 90:
            senk2 = perp_S_r * __e**(perp_S_a * Ti) + 1 - perp_S_r #South
        else:
            senk2 = perp_OW_r * __e**(perp_OW_a * Ti) + 1 - perp_OW_r #East / West
        
        #Calc the shading Factor
        ipol_hor = hor1 +1/2* (hor1 - hor2)*(1 - math.cos(2*(Azimuth - Azimuth_down) * (math.pi/180)))
        ipol_senk = senk1 +1/2* (senk2 - senk1)*(1 - math.cos(2*(Azimuth - Azimuth_down) * (math.pi/180)))
        V_Factor_Reveal = ipol_hor + 1/2 * (ipol_senk - ipol_hor) * (1 - math.cos( 2* Tilt * (math.pi/180)))
        
        return V_Factor_Reveal
    
    def Summer_OverhangShadingFactor(self, o_over, d_over, Tilt, Azimuth, GlazingHeight):
        #Set Up Input Values
        Azimuth_down = (int(Azimuth % 360) / 90) * 90
        Ti = o_over / (0.5 * GlazingHeight + d_over)
        Tu = o_over / (0.5 * GlazingHeight + d_over)
        
        #Set up the Horiz Constants
        hor_m1 = [-0.021486491, 0.002099882, 0.0000210709, 0.0000210709,0.023887087, -0.005637176]
        hor_b1 = [0.819347035, 0.081454448, 0.459193968, 0.459193968, 0.098922183, 1.092256566]
        hor_Deg1 = [30,90,90,30,90]
        hor_m2 = [0.006066841,-0.006285982,-0.001668327,-0.001668327,0.014495268,-0.006842841]
        hor_b2 = [-0.164581665,0.007487369,-0.234932205,-0.234932205,-0.372379119,0.10607216]
        hor_Deg2 = [15,90,90,23,90]
        
        perp_m1 = [0.0082589,-0.007607158,-0.001510116,-0.001510116,0.008437282,-0.007845213]
        perp_b1 = [0.254622217,0.52672416,0.155285893,0.155285893,0.193146659,0.683845013]
        perp_Deg1 = [15,90,90,30,90]
        perp_m2 = [0.076331095,-0.078055857,-0.00244834,-0.00244834,-0.016330529,-0.016330529]
        perp_b2 = [0.215819574,5.058130809,0.578662132,0.578662132,1.303608016,1.303608016]
        perp_Deg2 = [30,90,90, 90]
        
        #Set up the Horizontal Factors
        if self.Latitude <= hor_Deg1[0]:
            hor_S_r = hor_m1[0] * self.Latitude + hor_b1[0]
        else:
            hor_S_r = hor_m1[1] * self.Latitude + hor_b1[1]
        if self.Latitude <= hor_Deg1[2]:
            hor_OW_r = hor_m1[2] * self.Latitude + hor_b1[2]
        else:
            hor_OW_r = hor_m1[3] * self.Latitude + hor_b1[3]
        if self.Latitude <= hor_Deg1[3]:
            hor_N_r = hor_m1[4] * self.Latitude + hor_b1[4]
        else:
            hor_N_r = hor_m1[5] * self.Latitude + hor_b1[5]
        
        if self.Latitude <= hor_Deg2[0]:
            hor_S_a = hor_m2[0] * self.Latitude + hor_b2[0]
        else:
            hor_S_a = hor_m2[1] * self.Latitude + hor_b2[1]
        if self.Latitude <= hor_Deg2[2]:
            hor_OW_a = hor_m2[2] * self.Latitude + hor_b2[2]
        else:
            hor_OW_a = hor_m2[3] * self.Latitude + hor_b2[3]
        if self.Latitude <= hor_Deg2[3]:
            hor_N_a = hor_m2[4] * self.Latitude + hor_b2[4]
        else:
            hor_N_a = hor_m2[5] * self.Latitude + hor_b2[5]
        
        #Set up the Perpendicular Factors
        if self.Latitude <= perp_Deg1[0]:
            perp_S_r = perp_m1[0] * self.Latitude + perp_b1[0]
        else:
            perp_S_r = perp_m1[1] * self.Latitude + perp_b1[1]
        if self.Latitude <= perp_Deg1[2]:
            perp_OW_r = perp_m1[2] * self.Latitude + perp_b1[2]
        else:
            perp_OW_r = perp_m1[3] * self.Latitude + perp_b1[3]
        if self.Latitude <= perp_Deg1[3]:
            perp_N_r = perp_m1[4] * self.Latitude + perp_b1[4]
        else:
            perp_N_r = perp_m1[5] * self.Latitude + perp_b1[5]
        
        if self.Latitude <= perp_Deg2[0]:
            perp_S_a = perp_m2[0] * self.Latitude + perp_b2[0]
        else:
            perp_S_a = perp_m2[1] * self.Latitude + perp_b2[1]
        if self.Latitude <= perp_Deg2[2]:
            perp_OW_a = perp_m2[2] * self.Latitude + perp_b2[2]
        else:
            perp_OW_a = perp_m2[3] * self.Latitude + perp_b2[3]
        if self.Latitude <= perp_Deg2[3]:
            perp_N_a = perp_m2[4] * self.Latitude + perp_b2[4]
        else:
            perp_N_a = perp_m2[5] * self.Latitude + perp_b2[5]
        
        #Calc hori1
        if Azimuth_down == 0:
            hor1 = hor_N_r * __e**(hor_N_a * Ti) + 1 - hor_N_r #North
        elif Azimuth_down == 180:
            hor1 = hor_S_r * __e**(hor_S_a * Ti) + 1 - hor_S_r #South
        else:
            hor1 = hor_OW_r * __e**(hor_OW_a * Ti) + 1 - hor_OW_r #East / West
        
        #Calc hori2
        if Azimuth_down == 270:
            hor2 = hor_N_r * __e**(hor_N_a * Ti) + 1 - hor_N_r #North
        elif Azimuth_down == 90:
            hor2 = hor_S_r * __e**(hor_S_a * Ti) + 1 - hor_S_r #South
        else:
            hor2 = hor_OW_r * __e**(hor_OW_a * Ti) + 1 - hor_OW_r #East / West
        
        #Calc senk1
        if Azimuth_down == 0:
            senk1 = min(1, perp_N_r + ( 1 - perp_N_r) / (1 + Tu**2)**perp_N_a) #North
        elif Azimuth_down == 180:
            senk1 = min(1, perp_S_r + ( 1 - perp_S_r) / (1 + Tu**2)**perp_S_a) #South
        else:
            senk1 = min(1, perp_OW_r + ( 1 - perp_OW_r) / (1 + Tu**2)**perp_OW_a) #East / West
        
        #Calc senk2
        if Azimuth_down == 270:
            senk2 = perp_N_r + ( 1 - perp_N_r) / (1 + Tu**2)**perp_N_a #North
        elif Azimuth_down == 90:
            senk2 = perp_S_r + ( 1 - perp_S_r) / (1 + Tu**2)**perp_S_a #South
        else:
            senk2 = perp_OW_r + ( 1 - perp_OW_r) / (1 + Tu**2)**perp_OW_a #East / West
        
        #Calc the Shading Factors
        ipol_hor = hor1 + 1/2 * (hor2 - hor1) * (1-math.cos(2* (Azimuth - Azimuth_down) * (math.pi/180)))
        ipol_senk = senk1 + 1/2 * (senk2 - senk1) * (1-math.cos(2* (Azimuth - Azimuth_down) * (math.pi/180)))
        V_Factor_Overhang = ipol_hor + 1/2 * (ipol_senk - ipol_hor)*(1-math.cos(2* Tilt * (math.pi/180)))
        
        return V_Factor_Overhang

class PHPP_Window_Install:
    """ For storing the install conditions (0|1) of each edge in a window component """
    
    def __init__(self, _installs):
        """
        Args:
            _installs (list): A list of four 'install' types (Left, Right, Bottom, Top)
        """
        self.Installs = _installs
        self.setInstalls()
        
    def setInstalls(self):
        # In case the number of installs passed != 4, use the first one for all of them
        if len(self.Installs) != 4:
            self.Inst_L = float(self.Installs[0]) if self.Installs[0] != None else 'Auto'
            self.Inst_R = float(self.Installs[0]) if self.Installs[0] != None else 'Auto'
            self.Inst_B = float(self.Installs[0]) if self.Installs[0] != None else 'Auto'
            self.Inst_T = float(self.Installs[0]) if self.Installs[0] != None else 'Auto'
        else:
            self.Inst_L = float(self.Installs[0]) if self.Installs[0] != None else 'Auto'
            self.Inst_R = float(self.Installs[1]) if self.Installs[1] != None else 'Auto'
            self.Inst_B = float(self.Installs[2]) if self.Installs[2] != None else 'Auto'
            self.Inst_T = float(self.Installs[3]) if self.Installs[3] != None else 'Auto'
    
    def getAllasList(self):
        # So can easily output all the items in a single list
        return [int(self.Inst_L), int(self.Inst_R), int(self.Inst_B), int(self.Inst_T)]
    
    def __unicode__(self):
        return u'A PHPP Style Window Install Object: < L={self.Inst_L} | R={self.Inst_R} | T={self.Inst_T} | B={self.Inst_B} >'.format(self=self)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
       return "{}( _installs={!r} )".format(
               self.__class__.__name__, self.Installs)

class PHPP_Glazing:
    """ For storing PHPP Style Glazing Parameters """
    
    def __init__(self, _nm, _gValue, _uValue):
        """
        Args:
            _nm (str): The name of the glass type
            _gValue (float): The g-Value (SHGC) value of the glass only as per EN 410 (%)
            _uValue (float): The Thermal Trasmittance value of the center of glass (W/m2k) as per EN 673
        """
        self.Name = _nm
        try:
            self.gValue = float(_gValue)
        except:
            self.gValue = _gValue
        
        try:
            self.uValue = float(_uValue)
        except:
            self.uValue = _uValue
        
        
    def __unicode__(self):
        return u'A PHPP Style Glazing Object: < {self.Name} >'.format(self=self)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
       return "{}( _nm={!r}, _gValue={!r}, _uValue={!r} )".format(
               self.__class__.__name__,
               self.Name,
               self.gValue,
               self.uValue)

class PHPP_Frame:
    """ For Storing PHPP Style Frame Parameters """
    
    def __init__(self, _nm, _uValues, _frameWidths, _psiGlazings, _psiInstalls, _chiGlassCarrier=None):
        """
        Args:
            _nm (str): The name of the Frame Type
            _uValues (list): A list of the 4 U-Values (W/m2k) for the frame sides (Left, Right, Bottom, Top)
            _frameWidths (list): A list of the 4 U-Values (W/m2k) for the frame sides (Left, Right, Bottom, Top)
            _psiGlazings (list): A list of the 4 Psi-Values (W/mk) for the glazing spacers (Left, Right, Bottom, Top)
            _psiInstalls (list): A list of the 4 Psi-Values (W/mk) for the frame Installations (Left, Right, Bottom, Top)
            _chiGlassCarrier (list): A value for the Chi-Value (W/k) of the glass carrier for curtain walls
        """
        self.Name = _nm
        
        self.uValues = _uValues
        self.frameWidths = _frameWidths
        self.PsiGVals = _psiGlazings
        self.PsiInstalls = _psiInstalls
        self.chiGlassCarrier = _chiGlassCarrier
        
        self.cleanAttrSet(['uLeft', 'uRight', 'uBottom', 'uTop'], self.uValues)
        self.cleanAttrSet(['fLeft', 'fRight', 'fBottom', 'fTop'], self.frameWidths)
        self.cleanAttrSet(['psigLeft', 'psigRight', 'psigBottom', 'psigTop'], self.PsiGVals)
        self.cleanAttrSet(['psiInstLeft', 'psiInstRight', 'psiInstBottom', 'psiInstTop'], self.PsiInstalls)
    
    def cleanAttrSet(self, _inList, _attrList):
        # In case the input len != 4 and convert to float values
        if len(_attrList) != 4:
            try:
                val = float(_attrList[0])
            except:
                val = _attrList[0]
            
            for each in _inList:
                setattr(self, each, val)
        else:
            for i, each in enumerate(_inList):
                try:
                    val = float(_attrList[i])
                except:
                    val = _attrList[i]
                
                setattr(self, _inList[i], val)
    
    def __unicode__(self):
        return u'A PHPP Style Frame Object: < {self.Name} >'.format(self=self)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
       return "{}( _nm={!r}, _uValues={!r}, _frameWidths={!r}, _psiGlazings={!r}, "\
              "_psiInstalls={!r}, _chiGlassCarrier={!r} )".format(
               self.__class__.__name__,
               self.Name,
               self.uValues,
               self.frameWidths,
               self.PsiGVals,
               self.PsiInstalls,
               self.chiGlassCarrier )

class PHPP_Sys_Duct:
    def __init__(self, _lenM=5, _wMM=104, _iThckMM=52, _iLambda=0.04):
        self.DuctLength = _lenM
        self.DuctWidth = _wMM
        self.InsulationThickness = _iThckMM
        self.InsulationLambda = _iLambda
        
    def __unicode__(self):
        return u'HRV Duct Object Params'
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class PHPP_Sys_Ventilation:
    
    def __init__(self,
                _systemType='1-Balanced PH ventilation with HR',
                _systemName = 'Vent-1',
                _unitName='Default_Unit',
                _unitHR=0.75,
                _unitMR=0.0,
                _unitEE=0.45,
                _d01=PHPP_Sys_Duct(),
                _d02=PHPP_Sys_Duct(),
                _frsotT=-5,
                _ext=False,
                _exhaustObjs=[]):
        
        self.SystemType = _systemType
        self.SystemName = _systemName
        self.Unit_Name = _unitName
        self.Unit_HR = _unitHR
        self.Unit_MR = _unitMR
        self.Unit_ElecEff = _unitEE
        self.Duct01 = _d01
        self.Duct02 = _d02
        self.FrostTemp = _frsotT
        self.Exterior = 'x' if _ext==True else ''
        self.ExhaustObjs = _exhaustObjs
        
        self.setVentSystemType()
    
    def setVentSystemType(self):
        if '1' in self.SystemType:
            self.SystemType = '1-Balanced PH ventilation with HR'
        elif '2' in self.SystemType:
            self.SystemType = '2-Extract air unit'
        elif '3' in self.SystemType:
            self.SystemType = '3-Only window ventilation'
        else:
            warning = 'Error setting Ventilation System Type? Input only Type 1, 2, or 3. Setting to Type 1 (HRV) as default'
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
            self.SystemType = '1-Balanced PH ventilation with HR'
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    def __unicode__(self):
        return u'Ventilation System Params for: <{self.SystemName}>'.format(self=self)
    def __repr__(self):
        return "{}( _systemType={!r}, _systemName={!r}, _unitName={!r},"\
                "_unitHR={!r}, _unitMR={!r}, _unitEE={!r}, _d01={!r},"\
                "_d02={!r}, _frsotT={!r}, _ext={!r}, _exhaustObjs={!r})".format(
                self.__class__.__name__,
                self.SystemType,        
                self.SystemName,
                self.Unit_Name,
                self.Unit_HR,
                self.Unit_MR,
                self.Unit_ElecEff,
                self.Duct01,
                self.Duct02,
                self.FrostTemp,
                self.Exterior,
                self.ExhaustObjs)

class PHPP_Sys_VentUnit:
    def __init__(self, _nm='Default_Unit', _hr=0.75, _mr=0, _elec=0.45):
        self.Name = _nm
        self.HR_Eff = float(_hr)
        self.MR_Eff = float(_mr)
        self.Elec_Eff = float(_elec)
    
    def __unicode__(self):
        return u'Ventilation Unit Object Params for: <{self.Name}>'.format(self=self)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class PHPP_Sys_ExhaustVent:
    def __init__(self, nm, airFlowRate_On, airFlowRate_Off, hrsPerDay_On, daysPerWeek_On, default_duct):
        
        if nm == None or nm == '':
            self.Name = 'Exhaust Vent'
        else:
            self.Name = nm
        
        self.VentFloorArea = 10
        self.VentHeight = 2.5
        
        self.FlowRate_On = self._evaluateInputUnits(airFlowRate_On)
        self.FlowRate_Off = self._evaluateInputUnits(airFlowRate_Off)
        self.HrsPerDay_On = hrsPerDay_On
        self.DaysPerWeek_On = daysPerWeek_On
        self.Holidays = 0
        
        self.Duct01 = default_duct
        self.Duct02 = default_duct
        
        self._applyDefaults()
    
    def _evaluateInputUnits(self, _in):
        """If values are passed including a 'cfm' string, will
        set the return value to the m3/h equivalent"""
        
        if not _in:
            return None
        
        inputVal = _in.replace(' ', '')
        
        try:
            outputVal = float(inputVal)
        except:
            # Pull out just the decimal characters
            for each in re.split(r'[^\d\.]', inputVal):
                if len(each)>0:
                    outputVal = each
            
            # Convert to m3/h if necessary
            if 'cfm' in inputVal:
                outputVal = float(outputVal) * 1.699010796 #cfm--->m3/h
        
        return float(outputVal)
    
    def _applyDefaults(self):
        if not self.FlowRate_On: self.FlowRate_On = 450
        if not self.FlowRate_Off: self.FlowRate_Off = 25
        
        if not self.HrsPerDay_On: self.HrsPerDay_On = 0.5
        if not self.DaysPerWeek_On: self.DaysPerWeek_On = 7
    
    def __str__(self):
        return u'An Exhaust-Air Object: < {self.Name} >'.format(self=self)
    def __unicode__(self):
        return unicode(self).encode('utf-8')
    def __repr__(self):
       return "{}( nm={!r}, airFlowRate_On={!r}, airFlowRate_Off={!r},"\
               "hrsPerDay_On={!r},daysPerWeek_On={!r},"\
               "default_duct={!r} )".format(
               self.__class__.__name__,
               self.Name,
               self.FlowRate_On,
               self.FlowRate_Off,
               self.HrsPerDay_On,
               self.DaysPerWeek_On,
               self.Duct01)

class PHPP_TFA_Surface:
    def __init__(self, _tfaSrfc,  _zoneBreps, _roomVentFlowRates, _inset=0, _offsetZ=0):
        self.ID = random.randint(1000,9999)
        self.Neighbors = None
        self.InsetLen = _inset # meters
        self.OffsetZ = _offsetZ # meters
        self.TFASurface = _tfaSrfc
        self.ZoneBreps = _zoneBreps
        self.RoomVentFlowRates = _roomVentFlowRates
        
        if rs.coerceguid(self.TFASurface):
            # If the input is a GUID, try and bring 
            # in params and geometry from Rhino Scene
            [self.RoomName,
            self.RoomNumber,
            self.TFAfactor,
            self.V_sup,
            self.V_eta,
            self.V_trans,
            self.NonResUse,
            self.LightingControl,
            self.MotionControl ] = self.getSrfcAttrsFromRhino(self.TFASurface, self.RoomVentFlowRates)
            
            [self.Surface,
            self.Area_Gross,
            self.Centroid] = self.getGeomDataFromRhino(self.TFASurface)
        else:
            # Try and bring in params and geometry from Grasshopper
            [self.RoomName,
            self.RoomNumber,
            self.TFAfactor,
            self.Surface,
            self.Area_Gross,
            self.Centroid,
            self.V_sup,
            self.V_eta,
            self.V_trans,
            self.NonResUse,
            self.LightingControl,
            self.MotionControl ] = self.getParamsFromGH(self.TFASurface, self.RoomVentFlowRates)
            
        self.HostZoneName, self.HostZoneBrep, self.HostError = self.findHostZone(self.ZoneBreps)
        self.setOccupancyType()
        
        # Clean up the name
        if self.RoomName == None:
            self.RoomName = "{}_Room".format(self.HostZoneName)
        
        # Find the Surface Box Dims
        self.getRoomBoxDims()
    
    def setOccupancyType(self):
        if self.NonResUse != None and self.LightingControl != None and self.MotionControl != None:
            self.OccupancyType = 'Non-Residential'
        else:
            self.OccupancyType = 'Residential'
    
    def addNeighbor(self, _neighborName):
        self.Neighbors = _neighborName
        #self.Neighbors.append(_neighborName)
    
    def insetSurface(self, _srfc, _dist):
        # Take in a surface and and inset distance, 
        # Return an inset surface
        
        if _dist != 0:
            srfcPerim = ghc.JoinCurves( ghc.BrepEdges(_srfc)[0], preserve=False )
            
            # Get the inset Curve
            srfcCentroid = Rhino.Geometry.AreaMassProperties.Compute(_srfc).Centroid
            plane = ghc.XYPlane(srfcCentroid)
            srfcPerim_Inset_Pos = ghc.OffsetCurve(srfcPerim, _dist, plane, 1)
            srfcPerim_Inset_Neg = ghc.OffsetCurve(srfcPerim, _dist*-1, srfcCentroid, 1)
            #srfcPerim_Inset = ghc.OffsetonSrf(srfcPerim, _dist, _srfc) # Old. Remove
            
            # Choose the right Offset Curve. The one with the smaller area
            srfcInset_Pos = ghc.BoundarySurfaces( srfcPerim_Inset_Pos ) # Create a Surface from the boundary Curves
            srfcInset_Neg = ghc.BoundarySurfaces( srfcPerim_Inset_Neg ) # Create a Surface from the boundary Curves
            area_Pos = ghc.Area(srfcInset_Pos).area
            area_neg = ghc.Area(srfcInset_Neg).area
            if area_Pos < area_neg:
                return srfcInset_Pos
            else:
                return srfcInset_Neg
        
        else:
            return _srfc
    
    def findHostZone(self, _zoneBreps):
        srfcInset = self.insetSurface(self.Surface, 0.01) # Inset the TFA surface slightly and move 'up' (Z-axis) slightly
        srfcInset = ghc.Move(srfcInset, ghc.UnitZ(0.01) )[0]   # Move it 'up' 10mm just a tiny bit off floor
        
        # Find which Honeybee Zone the TFA Surface is 'inside' of 
        foundHost = []
        for zone in _zoneBreps:
            inside = ghc.ShapeInBrep(zone.geometry, srfcInset) 
            if inside == 0: # 0=Inside, 1=Intersecting, 2=Outside
                foundHost.append(True)
                hostName = zone.name  # Add the new Host Attribute to the TFA Surface Obj.
                hostBrep = zone.geometry
                break
            else:
                foundHost.append(False)
                hostName = None
                hostBrep = None
        
        # For if it can't find the Zone the surface is 'in'
        if True not in foundHost:
            hostError = True
        else:
            hostError = False
            
        return hostName, hostBrep, hostError
    
    def getParamsFromGH(self, _ghGeom, _ventRates):
        # Get the params for the TFA Obj from the Grasshopper Scene
        roomName = None
        roomNum = None
        roomTFAfactor = 1.0
        surface = self.insetSurface( ghc.BoundarySurfaces( _ghGeom ), self.InsetLen)
        area, centroid = ghc.Area(surface)
        
        nonResUse = '-'
        lightingControl = '-'
        motionControl = '-'
        
        # Bring in the Ventilation Flow Rates (if an) from the GH Panels
        if len(_ventRates)==3:
            Vent_Sup = _ventRates[0]
            Vent_Eta = _ventRates[1]
            Vent_Trans = _ventRates[2]
        elif _ventRates[0] == 'Automatic':
            Vent_Sup = 'Automatic'
            Vent_Eta = 'Automatic'
            Vent_Trans = 'Automatic'
            
        return (roomName, roomNum, roomTFAfactor, surface, area,
        centroid, Vent_Sup, Vent_Eta, Vent_Trans, nonResUse, lightingControl, motionControl )
    
    def cleanGet(self, _GUID, _attr, _type=None):
        sc.doc = Rhino.RhinoDoc.ActiveDoc
        if rs.IsUserText(_GUID):
            if _attr in rs.GetUserText(_GUID):
                if _type=='float':
                    output = float(rs.GetUserText(_GUID, _attr))
                else:
                    output = str(rs.GetUserText(_GUID, _attr))
            else:
                output = None
        sc.doc = ghdoc
        
        return output 
    
    def getSrfcAttrsFromRhino(self, _srfcGUID, _ventRates):
        
        # Pull the Basic User-Text Attributes from the Rhino Scene
        sc.doc = Rhino.RhinoDoc.ActiveDoc
        roomName = rs.ObjectName(_srfcGUID)
        roomNum = self.cleanGet(_srfcGUID, 'Room_Number')
        roomTFAfactor = self.cleanGet(_srfcGUID, 'TFA_Factor')
        
        # Get the Rhino Scene's Ventilation Flow Rates
        roomVentSup = self.cleanGet(_srfcGUID, 'V_sup', _type='float')
        roomVentExtr = self.cleanGet(_srfcGUID, 'V_eta', _type='float')
        roomVentTrans = self.cleanGet(_srfcGUID, 'V_trans', _type='float')
        
        # Get Any non-Res attrs
        roomNonResUse = self.cleanGet(_srfcGUID, 'useType')
        roomNonResLightingControl = self.cleanGet(_srfcGUID, 'lighting')
        roomNonResMotionControl = self.cleanGet(_srfcGUID, 'motion')
        
        # Get any GH Scene params as well. Overrider the Rhino Scene values
        if len(_ventRates)==3:
            roomVentSup = _ventRates[0]
            roomVentExtr = _ventRates[1]
            roomVentTrans = _ventRates[2]
        elif _ventRates[0] == 'Automatic':
            roomVentSup = 'Automatic' if roomVentSup==None else roomVentSup
            roomVentExtr = 'Automatic' if roomVentExtr==None else roomVentExtr
            roomVentTrans = 'Automatic' if roomVentTrans==None else roomVentTrans
        
        sc.doc = ghdoc
        
        return (roomName, roomNum, float(roomTFAfactor), roomVentSup, roomVentExtr,
        roomVentTrans, roomNonResUse, roomNonResLightingControl, roomNonResMotionControl)
    
    def getGeomDataFromRhino(self, _srfcGUID):
        # Pull the Geometry info from the Rhino Scene
        sc.doc = Rhino.RhinoDoc.ActiveDoc
        geom = rs.coercegeometry(_srfcGUID)
        #surface = ghc.BoundarySurfaces( geom ) #self.insetSurface( ghc.BoundarySurfaces( geom ), self.InsetLen )
        surface = ghc.DeconstructBrep(geom).faces
        area = rs.Area(_srfcGUID)
        centroid = Rhino.Geometry.AreaMassProperties.Compute(surface).Centroid
        sc.doc = ghdoc
        
        return surface, float(area), centroid
    
    def getArea_TFA(self):
        return self.Area_Gross * self.TFAfactor
    
    def getRoomBoxDims(self):
        worldXplane = ghc.XYPlane(Rhino.Geometry.Point3d(0,0,0))
        
        # Find the 'short' edge and the 'long' egde of the srfc
        srfcEdges = ghc.DeconstructBrep(self.TFASurface).edges
        segLengths = ghc.SegmentLengths(srfcEdges).longest_length
        srfcEdges_sorted = ghc.SortList(segLengths, srfcEdges).values_a
        endPoints = ghc.EndPoints(srfcEdges_sorted[-1])
        longEdgeVector = ghc.Vector2Pt(endPoints.start, endPoints.end, False).vector
        shortEdgeVector = ghc.Rotate(longEdgeVector, ghc.Radians(90), worldXplane).geometry
        
        # Use the edges to find the orientation and dimensions of the room
        srfcAligedPlane = ghc.ConstructPlane(ghc.Area(self.TFASurface).centroid, longEdgeVector, shortEdgeVector)
        srfcAlignedWorld = ghc.Orient(self.TFASurface, srfcAligedPlane, worldXplane).geometry
        dims = ghc.BoxProperties( srfcAlignedWorld ).diagonal
        dims = [dims.X, dims.Y]
        self.Depth = max(dims)
        self.Width = min(dims)
    
    def __unicode__(self):
        return u'The PHPP TFA Surface Object for {}-{}. TFA: {:.2f} m2'.format(self.RoomNumber, self.RoomName, self.getArea_TFA())
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _tfaSrfc={!r},  _zoneBreps={!r}, _roomVentFlowRates={!r},"\
                "_inset={!r}, _offsetZ={!r}".format(
               self.__class__.__name__,
               self.TFASurface,
               self.ZoneBreps,
               self.RoomVentFlowRates,
               self.InsetLen,
               self.OffsetZ)

class PHPP_RoomVolume:
    # An individual 'Volume' of a Room. Rooms can have one or more 'Volumes' / Zones / Areas
    
    def __init__(self, _tfaSurfaceObject, _roomGeom, _roomHeightUD=None, _vvType='Residential'):
        self.RoomHeightUD = _roomHeightUD
        self.RoomGeometry_GH = _roomGeom
        self.TFASrfcObj = _tfaSurfaceObject
        self.RoomNumber = _tfaSurfaceObject.RoomNumber
        self.RoomName = _tfaSurfaceObject.RoomName
        self.TFAsurface = _tfaSurfaceObject.Surface
        self.FloorArea_Gross = _tfaSurfaceObject.Area_Gross
        self.FloorArea_TFA = _tfaSurfaceObject.getArea_TFA()
        self.RoomVolumeTFAFactor = _tfaSurfaceObject.TFAfactor
        self.ZoneHeight, self.FloorZ = self.getHostZoneHeight(_tfaSurfaceObject)
        self.HostZoneName = _tfaSurfaceObject.HostZoneName
        self.tfaOffsetZ = _tfaSurfaceObject.OffsetZ
        self.tfaCentroid = _tfaSurfaceObject.Centroid
        self.VvType = getattr(self.TFASrfcObj, 'OccupancyType', _vvType)
        self.RoomWidth = _tfaSurfaceObject.Width
        self.RoomDepth = _tfaSurfaceObject.Depth
        
        # Room volume's ventilation flow rates
        self.V_sup = _tfaSurfaceObject.V_sup
        self.V_eta = _tfaSurfaceObject.V_eta
        self.V_trans = _tfaSurfaceObject.V_trans
        
        # Room's Non-Res Params
        self.RoomUseNonRes = _tfaSurfaceObject.NonResUse
        self.RoomLightingControl = _tfaSurfaceObject.LightingControl
        self.RoomMotionControl = _tfaSurfaceObject.MotionControl
        
        self.buildRoomVolume() # Create the Room-Volume Geometry
        
    def buildRoomVolume(self):
        # Build the RoomVolume and Figure out the Net Clear Volume (Vn50).
        # If can't make room from user-input geometry, auto-generate from the zone.
        try:
            self.RoomBrep = self.buildRoomBrepFromGeom()
            self.RoomNetClearVolume = abs(self.RoomBrep.GetVolume())
        except:
            self.RoomBrep = self.buildRoomBrepFromZone()
            self.RoomNetClearVolume = abs(self.RoomBrep.GetVolume())
        
        # Look to see if Vv should be 'Res' (TAx2.5m) or 'Non-Res' (actual clg height)
        if 'Non' in self.VvType:
            self.RoomClearHeight = self.RoomNetClearVolume / self.FloorArea_Gross
            self.RoomVentedVolume = self.FloorArea_TFA * self.RoomClearHeight
        else:
            self.RoomClearHeight = 2.5
            self.RoomVentedVolume = self.FloorArea_TFA * self.RoomClearHeight
    
    def buildRoomBrepFromGeom(self):
        # Creates a closed Brep from room geometry item passed in
        # The 'PHPP Rooms from Rhino' component will do the srfc joining and filtering
        # this def assumes a single input Brep 'open' on the bottom so can be joined to the
        # TFA surface. Note: as of May 2020, this is changed from the previous implementation where inputs
        # were a tree of curves / srfcs....
        
        closedBrep = None
        roomBrepSurfaces = []
        # See what type of geometry is being passed in. 
        
        if self.RoomGeometry_GH:
            if isinstance(self.RoomGeometry_GH, Rhino.Geometry.Brep):
                roomBrepSurfaces.append( self.RoomGeometry_GH )
            else:
                geomWarning = "Something is wrong with the input geometry for room: '{}'." \
                "Please only input a Brep (closed Polysurfce) for the room geometry." \
                "Be sure to leave OFF the bottom of the room Geometry Polysurface" \
                "so it can be joined to the TFA srfc".format(self.RoomName)
                ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, geomWarning)
            
        roomBrepSurfaces.append( self.TFAsurface ) # Add the TFA surface to each RoomBrep set
        roomBrep = ghc.BrepJoin(roomBrepSurfaces)
        
        if roomBrep.closed == True:
            closedBrep = roomBrep.breps
        
        if closedBrep:
                return closedBrep
        else:
            print ">>No Close-able room bounding geometry found for room {}-{}." \
            "Generating default room from Zone.".format(self.RoomNumber, self.RoomName)
    
    def getHostZoneHeight(self, _tfaSurfaceObject):
        hostBrep = _tfaSurfaceObject.HostZoneBrep
        verts = ghc.DeconstructBrep(hostBrep).vertices
        vertsZ = []
        
        for eachVert in verts:
            vertsZ.append(eachVert.Z)
        zDistance = abs( min(vertsZ) - max(vertsZ) )
        
        return zDistance, min(vertsZ)
    
    def buildRoomBrepFromZone(self):
        floorSurfaceEdges=[]
        flooroffsetFromUser = self.tfaOffsetZ # Up (Positive Z)
        flooroffserDraw = self.tfaCentroid.Z - self.FloorZ
        roomHeight = self.RoomHeightUD
        
        # Floor Surface
        floorSurface = rs.coercebrep(self.TFAsurface)
        floorSurface = ghc.Move(floorSurface, ghc.UnitZ(flooroffsetFromUser) )[0]  # 0 is the new translated geometry
        
        # Extrude the walls
        surfaceCentroid = Rhino.Geometry.AreaMassProperties.Compute(floorSurface).Centroid
        endPoint = ghc.ConstructPoint(surfaceCentroid.X, surfaceCentroid.Y, surfaceCentroid.Z+roomHeight)
        extrudeCurve = rs.AddLine(surfaceCentroid, endPoint)
        
        # Build the new Brep from all the surfaces
        newBrep = rs.ExtrudeSurface(surface=floorSurface, curve=extrudeCurve, cap=True)
        newBrep = rs.coercebrep(newBrep)
        
        return newBrep
    
    def __unicode__(self):
        return u'A PHPP Room-Volume Object: {}-{} TFA: {:.2f} m2'.format(self.RoomNumber, self.RoomName, self.FloorArea_TFA)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _tfaSurfaceObject={!r}, _roomGeom={!r},"\
                " _roomHeightUD={!r}, _vvType={!r}".format(
               self.__class__.__name__,
               self.TFASrfcObj,
               self.RoomGeometry_GH,
               self.RoomHeightUD,
               self.VvType)

class PHPP_Room:
    # Holds the basic room params. Made up of one or more 'RoomVolumes'.
    # Each Volume has a TFA Surface
    
    def __init__(self, _roomVolumes):
        # Just pull this info from the first Room Volume passed in
        self.RoomVolumes = _roomVolumes
        self.RoomNumber = self.RoomVolumes[0].RoomNumber
        self.RoomName = self.RoomVolumes[0].RoomName
        self.ZoneHeight = self.RoomVolumes[0].ZoneHeight
        self.HostZoneName = self.RoomVolumes[0].HostZoneName
        
        self.getTotalRoomValues(self.RoomVolumes)
        self.getRoomBreps(self.RoomVolumes)
        self.getNonRes()
    
    def getRoomBreps(self, roomVolObjs):
        self.RoomBreps = []
        for eachRoomVol in roomVolObjs:
            self.RoomBreps.append( eachRoomVol.RoomBrep )
    
    def getTotalRoomValues(self, roomVolObjs):
        # Gets the toal TFA, FA, and a weighted Avg Room TFA Factor
        
        tfaSurfaces = []
        tfaFactors = []
        floorAreas_Gross = []
        floorAreas_TFA = []
        netClearVolume = []
        ventedVolume = []
        avgClearHeight = []
        roomWidths, roomDepths = [], []
        
        for eachRoomVolume in roomVolObjs:
            roomWidths.append(eachRoomVolume.RoomWidth)
            roomDepths.append(eachRoomVolume.RoomDepth)
            tfaSurfaces.append(eachRoomVolume.TFAsurface)
            tfaFactors.append(eachRoomVolume.RoomVolumeTFAFactor)
            floorAreas_Gross.append( float(eachRoomVolume.FloorArea_Gross) )
            floorAreas_TFA.append(  float(eachRoomVolume.FloorArea_TFA) )
            netClearVolume.append( float(eachRoomVolume.RoomNetClearVolume) )
            ventedVolume.append(float(eachRoomVolume.RoomVentedVolume) )
            avgClearHeight.append( float(eachRoomVolume.RoomClearHeight ) )
        
        self.RoomWidth = max(roomWidths)
        self.RoomDepth = max(roomDepths)
        
        self.TFAsurface = tfaSurfaces #This is now a list of items????
        self.TFAfactors = tfaFactors
        self.FloorArea_Gross = sum(floorAreas_Gross)
        self.FloorArea_TFA = sum(floorAreas_TFA)
        self.RoomTFAfactor = self.FloorArea_TFA / self.FloorArea_Gross
        self.RoomNetClearVolume = sum(netClearVolume) 
        self.RoomVentedVolume = sum(ventedVolume)
        self.RoomClearHeight = sum(avgClearHeight)/len(avgClearHeight)
        
        # Get the room's Ventilation Air Flows
        v_sups = []
        v_etas = []
        v_transfs = []
        
        for eachRoomVolume in roomVolObjs:
            v_sups.append(eachRoomVolume.V_sup)
            v_etas.append(eachRoomVolume.V_eta)
            v_transfs.append(eachRoomVolume.V_trans)
        
        self.V_sup = max(v_sups)
        self.V_eta = max(v_etas)
        self.V_trans =max(v_transfs)
    
    def getNonRes(self):
        temp_Uses = []
        temp_lighting = []
        temp_motion = []
        
        for room in self.RoomVolumes:
            temp_Uses.append( getattr(room, 'RoomUseNonRes', None) )
            temp_lighting.append( getattr(room, 'RoomLightingControl', None) )
            temp_motion.append( getattr(room, 'RoomMotionControl', None) )
        
        # Get the Lighting Control Type
        # Use the value which occurs most often
        count = []
        for i, each in enumerate(temp_Uses):
            count.append(temp_Uses.count(each))
        
        self.NonRes_RoomUse = temp_Uses[count.index(max(count))]
        
        # Get the Motion Control Type
        count = []
        for i, each in enumerate(temp_lighting):
            count.append(temp_lighting.count(each))
        
        self.NonRes_RoomLightingControl = temp_lighting[count.index(max(count))]
        
        # Get Is there Motion Control
        if 'Yes' in temp_motion: self.NonRes_RoomMotionControl = 'Yes'
        elif 'No' in temp_motion: self.NonRes_RoomMotionControl = 'No'
        else: self.NonRes_RoomMotionControl = None
    
    def getVsup(self):
        try:
            return float(self.V_sup)
        except:
            if 'Auto' in self.V_sup:
                return 0
            else:
                return self.V_sup
    
    def getVeta(self):
        try:
            return float(self.V_eta)
        except:
            if 'Auto' in self.V_eta:
                return 0
            else:
                return self.V_eta
    
    def getVtrans(self):
        try:
            return float(self.V_trans)
        except:
            if 'Auto' in self.V_trans:
                return 0
            else:
                return self.V_trans
    
    def __unicode__(self):
        return u'A PHPP Room Object: {}-{} TFA: {:.2f}m2 <TFA Factor: {:.2f}>'.format(self.RoomNumber, self.RoomName, self.FloorArea_TFA, self.RoomTFAfactor)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _roomVolumes={!r}".format(
               self.__class__.__name__,
               self._roomVolumes)

class PHPP_DHW_System:
    def __init__(self, _name='DHW', _usage=None, _fwdT=60, _pCirc=[], _pBran=[], _t1=None, _t2=None, _tBf=None):
        self.SystemName = _name
        self.usage = _usage
        self.forwardTemp = _fwdT if _fwdT != None else 60
        self.circulation_piping = _pCirc
        self.branch_piping = _pBran
        self.tank1 = _t1
        self.tank2 = _t2
        self.tank_buffer = _tBf
        self.ZonesAssigned = []
    
    def getZonesAssignedList(self):
        self.ZonesAssigned = list(set(self.ZonesAssigned))
        return self.ZonesAssigned
    
    def __unicode__(self):
        return u'A PHPP Style DHW System: < {self.SystemName} >'.format(self=self)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
       return "{}( _name={!r}, _usage={!r}, _fwdT={!r}, _pCirc={!r}, "\
              "_pBran={!r}, _t1={!r}, _t2={!r}, _tBf={!r} )".format(
               self.__class__.__name__,
               self.SystemName,
               self.usage,
               self.forwardTemp,
               self.circulation_piping,
               self.branch_piping,
               self.tank1,
               self.tank2,
               self.tank_buffer)

class PHPP_DHW_usage:
    
    def __init__(self, args):
        self.UsageType = 'Res'
        self.demand_showers = args.get('showers_demand_', 16)
        self.demand_others = args.get('other_demand_', 9)
    
    def __unicode__(self):
        return u'A Residential DHW usage profile Object'.format()
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( showers_demand_={}, other_demand_={!r} )".format(
               self.__class__.__name__,
               self.demand_showers,
               self.demand_others)

class PHPP_DHW_usage_NonRes:
    
    def __init__(self, args):
        self.UsageType = 'NonRes'
        self.use_daysPerYear = args.get('useDaysPerYear_', 365)
        self.useShowers = args.get('showers_', 'x')
        self.useHandWashing = args.get('handWashBasin_', 'x')
        self.useWashStand = args.get('washStand_', 'x')###########
        self.useBidets = args.get('bidet_', 'x')
        self.useBathing = args.get('bathing_', 'x')
        self.useToothBrushing = args.get('toothBrushing_', 'x')
        self.useCooking = args.get('cookingAndDrinking_', 'x')
        self.useDishwashing = args.get('dishwashing_', 'x')
        self.useCleanKitchen = args.get('cleaningKitchen_', 'x')
        self.useCleanRooms = args.get('cleaningRooms_', 'x')
    
    def __unicode__(self):
        return u'A Non-Residential DHW usage profile Object'.format()
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( useDaysPerYear_={!r}, showers_={!r}, handWashBasin_={!r}, bidet_={!r}, bathing_={!r},"\
        "toothBrushing_={!r}, cookingAndDrinking_={!r}, dishwashing_={!r}, cleaningKitchen_={!r}, cleaningRooms_={!r})".format(
               self.__class__.__name__,
               self.use_daysPerYear,
               self.useShowers,
                self.useHandWashing,
                self.useBidets,
                self.useBathing,
                self.useToothBrushing,
                self.useCooking,
                self.useDishwashing,
                self.useCleanKitchen,
                self.useCleanRooms)

class PHPP_DHW_RecircPipe:
    def __init__(self, _len=None, _d=None, _t=None, _lam=None, _ref=None, _q=None, _p=None):
        self.length = _len
        self.diam = _d
        self.insulThck = _t
        self.insulCond = _lam
        self.insulRefl = _ref
        self.quality = _q
        self.period = _p
    
    def __unicode__(self):
        return u'A DHW Recirculation Pipe Object'
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _len={!r}, _d={!r}, _t={!r}, "\
              "_lam={!r}, _ref={!r} _q={!r}, _p={!r})".format(
               self.__class__.__name__,
               self.length,
               self.diam,
               self.insulThck,
               self.insulCond,
               self.insulRefl,
               self.quality,
               self.period)

class PHPP_DHW_branch_piping:
    def __init__(self, _d=None, _len=None, _numTaps=None, _opens=None, _utiliz=None):
        self.diameter = _d
        self.totalLength = _len
        self.totalTapPoints = _numTaps
        self.tapOpenings = _opens
        self.utilisation = _utiliz
        
    def __unicode__(self):
        return u'A DHW Branch Piping Object'
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _d={!r}, _len={!r}, _numTaps={!r}, "\
              "_opens={!r}, _utiliz={!r} )".format(
               self.__class__.__name__,
               self.diameter,
               self.totalLength,
               self.totalTapPoints,
               self.tapOpenings,
               self.utilisation)

class PHPP_DHW_tank:
    def __init__(self, _type, _solar, _hl_rate, _vol, _stndby_frac, _loc, _loc_T):
        self.type = _type
        self.solar = _solar
        self.hl_rate = _hl_rate
        self.vol = _vol
        self.stndbyFrac = _stndby_frac
        self.loction = _loc if _loc != None else "1-Inside"
        self.locaton_t = _loc_T if _loc != None else ""
        
    def __unicode__(self):
        return u'A DHW Tank Object'
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _type={!r}, _solar={!r}, _hl_rate={!r}, "\
              "_vol={!r}, _stndby_frac={!r} _loc={!r}, _loc_T={!r})".format(
               self.__class__.__name__,
               self.type,
               self.solar,
               self.hl_rate,
               self.vol,
               self.stndbyFrac,
               self.loction,
               self.locaton_t)

class PHPP_grnd_FloorElement():
    """Master class with all defs for PHPP Ground / Floor Elements"""
    
    __warnings = None
    
    def __init__(self):
        pass
    
    def getWarnings(self):
        if self.__warnings == None:
            self.__warnings = []
        return self.__warnings
    
    def getSurfaceData(self, _flrSrfcs):
        """Pulls Rhino Scene parameters for a list of Surface Objects
        
        Takes in a list of surface GUIDs and goes to Rhino. Looks in their
        UserText to get any user applied parameter data. Will also find the
        surface area of each surface in the list. 
        
        Calculates the total surface area of all surfaces in the list, as well as
        the area-weighted U-Value of the total.
        
        Will return tuple(0, 0) on any trouble getting parameter values or any fails
        
        Parameters:
        _flrSrfcs (list): A list of Surface GUIDs to look at in the Rhino Scene
        
        Returns:
        totalFloorArea, floorUValue (tuple): The total floor area (m2) and the area weighted U-Value
        """
        
        if _flrSrfcs == None:
            return 0,0
        
        if len(_flrSrfcs)>0:
            floorAreas = []
            weightedUvales = []
            
            sc.doc = Rhino.RhinoDoc.ActiveDoc
            for srfcGUID in _flrSrfcs:
                # Get the Surface Area Params
                srfcGeom = rs.coercebrep(srfcGUID)
                if srfcGeom:
                    srfcArea = ghc.Area(srfcGeom).area
                    floorAreas.append( srfcArea )
                    
                    # Get the Surface U-Values Params
                    srfcUvalue = self.getSrfcUvalue(srfcGUID)
                    weightedUvales.append(srfcUvalue * srfcArea )
                else:
                    floorAreas.append( 1 )
                    weightedUvales.append( 1 )
                    
                    warning = 'Error: Input into _floorSurfaces is not a Surface?\n'\
                    'Please ensure inputs are Surface Breps only.'
                    self.getWarnings().append(warning)
                
            sc.doc = ghdoc
            
            totalFloorArea = sum(floorAreas)
            floorUValue = sum(weightedUvales) / totalFloorArea
        
        else:
            totalFloorArea = 0
            floorUValue = 0
        
        return totalFloorArea, floorUValue
    
    def getExposedPerimData(self, _perimCrvs, _UDperimPsi):
        """Pulls Rhino Scene parameters for a list of Curve Objects
        
        Takes in a list of curve GUIDs and goes to Rhino. Looks in their
        UserText to get any user applied parameter data.
        
        Calculates the sum of all curve lengths (m) in the list, as well as
        the total Psi-Value * Length (W/m) of all curves in the list.
        
        Will return tuple(0, 0) on any trouble getting parameter values or any fails
        
        Parameters:
        _perimCrvs (list): A list of Curve GUIDs to look at in the Rhino Scene
        
        Returns:
        totalLen, psiXlen (tuple): The total length of all curves in the list (m) and the total Psi*Len value (W/m)
        """
        
        def getLengthIfNumber(_in, _psi):
            try:
                length = float(_in)
            except:
                length = None
            
            try:
                psi = float(_psi)
            except:
                psi = 0.05
        
            return length, psi
        
        if _perimCrvs == None:
            return 0, 0, None
        
        psiXlen = 0
        totalLen = 0
        warning = None
        
        if len(_perimCrvs)>0:
            sc.doc = Rhino.RhinoDoc.ActiveDoc
            for crvGUID in _perimCrvs:
                
                # See if its just Numbers passed in. If so, use them and break out
                length, crvPsiValue = getLengthIfNumber(crvGUID, _UDperimPsi)
                if length and crvPsiValue:
                    totalLen += length
                    psiXlen += (length * crvPsiValue)
                    continue
                
                isCrvGeom = rs.coercecurve(crvGUID)
                isBrepGeom = rs.coercebrep(crvGUID)
                
                if isCrvGeom:
                    crvLen = ghc.Length(isCrvGeom)
                    try:
                        crvPsiValue = float(_UDperimPsi)
                    except:
                        crvPsiValue, warning = self.getCurvePsiValue(crvGUID)
                    
                    totalLen += crvLen
                    psiXlen += (crvLen * crvPsiValue)
                elif isBrepGeom:
                    srfcEdges = ghc.DeconstructBrep(isBrepGeom).edges
                    srfcPerim = ghc.JoinCurves(srfcEdges, False)
                    crvLen = ghc.Length(srfcPerim)
                    
                    try:
                        crvPsiValue = float(_UDperimPsi)
                    except:
                        crvPsiValue = 0.05
                    
                    totalLen += crvLen
                    psiXlen += (crvLen * crvPsiValue) # Default 0.05 W/mk
                    warning = 'Note: You passed in a surface without any Psi-Values applied.\n'\
                    'I will apply a default 0.5 W/mk Psi-Value to ALL the edges of the\n'\
                    'surface passed in.'
                else:
                    warning = 'Error in GROUND: Input into _exposedPerimCrvs is not a Curve or Surface?\n'\
                    'Please ensure inputs are Curve / Polyline or Surface only.'
            sc.doc = ghdoc
            
            return totalLen, psiXlen, warning
        else:
            return 0, 0, None
        
    def getCurvePsiValue(self, _perimCrvGUID):
        """Takes in a single Curve GUID and returns its length and Psi*Len
        
        Will look at the UserText of the curve to get the Psi Value Type
        Name and then the Document UserText library to get the Psi-Value of the
        Type. 
        
        Returns 0.5 W/mk as default on any errors.
        
        Parameters:
        _perimCrvGUID (GUID): A single GUID 
        
        Returns:
        crvPsiValue (float): The Curve's UserText Param for 'Psi-Value' if found.
        """
        
        warning = None
        crvPsiValueName = rs.GetUserText(_perimCrvGUID, 'Typename')
        if crvPsiValueName:
            for k in rs.GetDocumentUserText():
                if 'PHPP_lib_TB' not in k:
                    continue
                
                try:
                    d = json.loads(rs.GetDocumentUserText(k))
                    if d.get('Name', None) == crvPsiValueName:
                        psiValParams = rs.GetDocumentUserText(k)
                        break
                except:
                    psiValParams = None
            else:
                psiValParams = None
            
            if psiValParams:
                psiValParams = json.loads(psiValParams)
                crvPsiValue = psiValParams.get('psiValue', 0.5)
                if crvPsiValue < 0:
                    warning = 'Warning: Negative Psi-Value found for type: "{}"\nApplying 0.0 W/mk for that edge.'.format(crvPsiValueName)
                    self.getWarnings().append(warning)
                    crvPsiValue = 0
            else:   
                warning = ('Warning: Could not find a Psi-Value type in the',
                'Rhino Document UserText with the name "{}?"'.format(crvPsiValueName.upper()),
                'Check your Document UserText library to make sure that you have',
                'your most recent Thermal Bridge library file loaded?',
                'For now applying a Psi-Value of 0.5 w/mk for this edge.')
                self.getWarnings().append(warning)
                crvPsiValue = 0.5
        else:
            warning = 'Warning: could not find a Psi-Value type in the\n'\
            'UserText document library for one or more edges?\n'\
            'Check your Document UserText library to make sure that you have\n'\
            'your most recent Thermal Bridge library file loaded?\n'\
            'For now applying a Psi-Value of 0.5 w/mk for this edge.'
            self.getWarnings().append(warning)
            crvPsiValue = 0.5
        
        return crvPsiValue, warning
    
    def getSrfcUvalue(self, _srfcGUID):
        """Takes in a single Surface GUID and returns its U-Value Param
        
        Will look at the UserText of the surface to get the EP Construction
        Name and then the Document UserText library to get the U-Value of tha
        Construction Type. 
        
        Returns 1.0 W/m2k as default on any errors.
        
        Parameters:
        _srfcGUID (GUID): A single GUID value
        
        Returns:
        srfcUvalue (float): The Surface's UserText Param for 'U-Value' if found
        """
        
        srfcConstructionName = rs.GetUserText(_srfcGUID, 'EPConstruction')
        if srfcConstructionName:
            constParams = rs.GetDocumentUserText('PHPP_lib_Assmbly_' + srfcConstructionName)
            
            for k in rs.GetDocumentUserText():
                if 'PHPP_lib_Assmbly' not in k:
                    continue
                try:
                    d = json.loads(rs.GetDocumentUserText(k))
                    if d.get('Name', None) == srfcConstructionName:
                        constParams = rs.GetDocumentUserText(k)
                        break
                except:
                    constParams = None
            else:
                constParams = None
            
            if constParams:
                constParams = json.loads(constParams)
                srfcUvalue = constParams.get('uValue', 1)
            else:
                warning = ('Warning: Could not find a construction type in the',
                'Rhino Document UserText with the name "{}?"'.format(srfcConstructionName.upper()),
                'Check your Document UserText library to make sure that you have',
                'your most recent assembly library file loaded?',
                'For now applying a U-Value of 1.0 w/m2k for this surface.')
                self.getWarnings().append(warning)
                #ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, "\n".join(warning))
                srfcUvalue = 1.0
        else:
            warning = 'Warning: could not find a construction type in the\n'\
            'UserText for one or more surfaces? Are you sure you assigned a\n'\
            '"EPConstruction" parameter to the Floor Surface being input?\n'\
            'For now applying a U-Value of 1.0 w/m2k for this surface.'
            self.getWarnings().append(warning)
            #ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, warning)
            srfcUvalue = 1.0
        
        return srfcUvalue
    
    def getParamsFromRH(self, _HBZoneObj, _type):
        """ Finds the Floor Type Element surfaces in a Honeybee Zone and gets Params
        
        Resets self.FloorArea and self.FloorUValue based on the values found in
        the Honeybee Zone. If more than one SlabOnGrade is found, creates a 
        weighted U-Value average of all the surfaces.
        Arguments:
            _HBZoneObj: A Single Honeybee Zone object
            _type: (str) the type name ('SlabOnGrade', 'UngergoundSlab', 'ExposedFloor') of the floor element
        Returns:
            None
        """
        
        # Try and get the floor surface data
        slabAreas = []
        slabWeightedUvalue = []
        for srfc in _HBZoneObj.surfaces:
            if srfc.srfType[srfc.type] == _type:
                cnstrName = srfc.EPConstruction
                result = hb_EPMaterialAUX.decomposeEPCnstr(str(cnstrName).upper())
                if result != -1:
                    materials, comments, UValue_SI, UValue_IP = result
                else:
                    UValue_SI = 1
                    warning = 'Warning: Can not find material "{}" in the HB Library for surface?'.format(str(cnstrName))
                    self.getWarnings().append(warning)
                    #ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
                
                slabArea = srfc.getArea()
                slabAreas.append( slabArea )
                slabWeightedUvalue.append( UValue_SI * slabArea )
    
        if len(slabAreas)>0 and len(slabWeightedUvalue)>0:
            self.FloorArea = sum(slabAreas)
            self.FloorUvalue = sum(slabWeightedUvalue) / self.FloorArea
        
        #Try and get any below grade wall U-Values (area weighted)
        wallBG_Areas = []
        wallBG_UxA = []
        for srfc in _HBZoneObj.surfaces:
            if srfc.srfType[srfc.type] == 'UndergroundWall':
                result = hb_EPMaterialAUX.decomposeEPCnstr(str(cnstrName).upper())
                if result != -1:
                    materials, comments, UValue_SI, UValue_IP = result
                else:
                    UValue_SI = 1
                    warning = 'Warning: Can not find material "{}" in the HB Library for surface?'.format(str(cnstrName))
                    self.getWarnings().append(warning)
                    #ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
                
                wallArea = srfc.getArea()
                wallBG_Areas.append( wallArea )
                wallBG_UxA.append( UValue_SI * wallArea )
        
        if len(wallBG_Areas)>0 and len(wallBG_UxA)>0:
            self.WallU_BG = sum(wallBG_UxA) / sum(wallBG_Areas)

class PHPP_grnd_SlabOnGrade(PHPP_grnd_FloorElement):
    """Ground subclass for Slab On Grade Floor Elements
    
    Attributes:
        _flrSrfcs (list): A list of Floor Surfaces from Rhino Scene
        _perimCrvs (list): A list of Curves from Rhino Scene
        _depth (float): The Width / Depth (m) of the perimeter insulation
        _thick (float): The Thickness (m) of the perimeter insualtion
        _cond (float): The Thermal Conductivity (W/mk) of the perimeter insulation
        _orient (str): The Orientation of the perimeter insualtion. Input either 'Vertical' or 'Horizontal'
    """
    
    def __init__(self, _zoneName=None, _flrSrfcs=None, _perimCrvs=None, _perimPsi=0.5 ,_depth=1.0, _thick=0.101, _cond=0.04, _orient='Vertical'):
        self.Warning = None
        self.Type = '01_SlabOnGrade'
        self.soilThermalConductivity = 2.0 # MJ/m3-K
        self.soilHeatCapacity = 2.0 # W/mk
        self.groundWaterDepth = 3.0 # m
        self.groundWaterFlowrate = 0.05 # m/d 
        
        self.Zone = _zoneName
        self.FloorSurface = _flrSrfcs
        self.FloorArea, self.FloorUvalue = self.getSurfaceData(self.FloorSurface)
        self.PerimCurves = _perimCrvs
        self.PerimPsiVal = _perimPsi
        self.PerimLen, self.PerimPsixLen, self.Warning = self.getExposedPerimData(self.PerimCurves, self.PerimPsiVal)
        self.perimInsulDepth = _depth
        self.perimInsulThick = _thick
        self.perimInsulConductivity = _cond
        self.perimInsulOrientation = _orient
    
    def __unicode__(self):
        return u'A Ground Floor Object of Type: "{!r}":'.format(self.Type)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _zoneName={}, _flrSrfcs={!r}, _perimCrvs={!r}, _perimPsi={!r},"\
                "_depth={!r}, _thick={!r}, _cond={!r}, _orient={!r})".format(
               self.__class__.__name__,
               self.Zone,
               self.FloorSurface,
               self.PerimCurves,
               self.PerimPsiVal,
               self.perimInsulDepth, 
               self.perimInsulThick, 
               self.perimInsulConductivity, 
               self.perimInsulOrientation)

class PHPP_grnd_HeatedBasement(PHPP_grnd_FloorElement):
    """Ground subclass for Heated Basement Slab
    """
    def __init__(self, _zoneName=None, _flrSrfcs=None, _perimCrvs=None, _perimPsi=0.5, _wallHeight_BG=1.0, _wallU_BG=1.0):
        self.Warning = None
        self.Type = '02_HeatedBasement'
        self.soilThermalConductivity = 2.0 # MJ/m3-K
        self.soilHeatCapacity = 2.0 # W/mk
        self.groundWaterDepth = 3.0 # m
        self.groundWaterFlowrate = 0.05 # m/d 
        
        self.Zone = _zoneName
        self.FloorSurfaces = _flrSrfcs
        self.FloorArea, self.FloorUvalue = self.getSurfaceData(self.FloorSurfaces)
        self.ExposedPerimCrvs = _perimCrvs
        self.PerimPsiVal = _perimPsi
        self.PerimLen, self.PerimPsixLen, self.Warning = self.getExposedPerimData(self.ExposedPerimCrvs, self.PerimPsiVal)
        self.WallHeight_BG = _wallHeight_BG
        self.WallU_BG = _wallU_BG
    
    def __unicode__(self):
        return u'A Ground Floor Object of Type: "{!r}":'.format(self.Type)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _zoneName={}, _flrSrfcs={!r}, _perimCrvs={!r}, _perimPsi={!r}, _wallHeight_BG={!r}, _wallU_BG={!r})".format(
               self.__class__.__name__,
               self.Zone,
               self.FloorSurfaces,
               self.ExposedPerimCrvs, 
               self.PerimPsiVal,
               self.WallHeight_BG, 
               self.WallU_BG)

class PHPP_grnd_UnheatedBasement(PHPP_grnd_FloorElement):
    """Ground subclass for a Suspended Floor over an Un-Heated Basement
    """
    def __init__(self, _zoneName=None, _flrSrfcs=None, _perimCrvs=None, _perimPsi=0.5, _wallHeight_AG=1.0, _wallU_AG=1.0, _wallHeight_BG=1.0, _wallU_BG=1.0, _flrU=1.0, _ach=1.0, _vol=1.0):
        self.Warning = None
        self.Type = '03_UnheatedBasement'
        self.soilThermalConductivity = 2.0 # MJ/m3-K
        self.soilHeatCapacity = 2.0 # W/mk
        self.groundWaterDepth = 3.0 # m
        self.groundWaterFlowrate = 0.05 # m/d 
        
        self.Zone = _zoneName
        self.FloorSurfaces = _flrSrfcs
        self.FloorArea, self.FloorUvalue = self.getSurfaceData(self.FloorSurfaces)
        self.ExposedPerimCrvs = _perimCrvs
        self.PerimPsiVal = _perimPsi
        self.PerimLen, self.PerimPsixLen, self.Warning = self.getExposedPerimData(self.ExposedPerimCrvs, self.PerimPsiVal)
        self.WallHeight_BG = _wallHeight_BG
        self.WallU_BG = _wallU_BG
        self.WallHeight_AG = _wallHeight_AG
        self.WallU_AG = _wallU_AG
        self.FloorU = _flrU
        self.ACH = _ach
        self.Volume = _vol
    
    def __unicode__(self):
        return u'A Ground Floor Object of Type: "{!r}":'.format(self.Type)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _zoneName={}, _flrSrfcs={!r}, _perimCrvs={!r}, _perimPsi={!r}, _wallHeight_AG={!r}, _wallU_AG={!r},"\
              "_wallHeight_BG={!r}, _wallU_BG={!r} _flrU={!r}, _ach={!r}, _vol={!r})".format(
               self.__class__.__name__,
               self.Zone,
               self.FloorSurfaces,
               self.ExposedPerimCrvs,
               self.PerimPsiVal,
               self.WallHeight_AG,
               self.WallU_AG, 
               self.WallHeight_BG,
               self.WallU_BG, 
               self.FloorU, 
               self.ACH, 
               self.Volume)

class PHPP_grnd_SuspendedFloor(PHPP_grnd_FloorElement):
    """Ground subclass for a Suspended Floor over an unheated Crawlspace
    """
    def __init__(self, _zoneName=None, _flrSrfcs=None, _perimCrvs=None, _perimPsi=0.5, _wallHeight=1.0, _wallU=1.0, _crawlU=1.0, _ventOpen=1.0, _windVel=4.0, _windFac=0.05):
        self.Warning = None
        self.Type = '04_SuspenedFlrOverCrawl'
        self.soilThermalConductivity = 2.0 # MJ/m3-K
        self.soilHeatCapacity = 2.0 # W/mk
        self.groundWaterDepth = 3.0 # m
        self.groundWaterFlowrate = 0.05 # m/d 
        
        self.Zone = _zoneName
        self.FloorSurfaces = _flrSrfcs
        self.FloorArea, self.FloorUvalue = self.getSurfaceData(self.FloorSurfaces)
        self.ExposedPerimCrvs = _perimCrvs
        self.PerimPsiVal = _perimPsi
        self.PerimLen, self.PerimPsixLen, self.Warning = self.getExposedPerimData(self.ExposedPerimCrvs, self.PerimPsiVal)
        self.WallHeight = _wallHeight
        self.WallU = _wallU
        self.CrawlU = _crawlU
        self.VentOpeningArea = _ventOpen
        self.windVelocity = _windVel
        self.windFactor = _windFac
    
    def __unicode__(self):
        return u'A Ground Floor Object of Type: "{!r}":'.format(self.Type)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
        return "{}( _zoneName={}, _flrSrfcs={!r}, _perimCrvs={!r}, _perimPsi={!r}, _wallHeight={!r}, "\
              "_wallU={!r}, _crawlU={!r} _ventOpen={!r}, _windVel={!r}, _windFac={!r})".format(
               self.__class__.__name__,
               self.Zone,
               self.FloorSurfaces,
               self.ExposedPerimCrvs,
               self.PerimPsiVal,
               self.WallHeight,
               self.WallU,
               self.CrawlU,
               self.VentOpeningArea,
               self.windVelocity,
               self.windFactor)

class PHPP_ClimateDataSet:
    
    def __init__(self, _dataSet='US0055b-New York', _alt='=J23', _cntry='US-United States of America', _reg='New York'):
        self.DataSet = _dataSet
        self.Altitude = _alt
        self.Country = _cntry
        self.Region = _reg
    
    def __unicode__(self):
        return u'A Location Object for: "{!r}":'.format(self.DataSet)
    def __str__(self):
        return unicode(self).encode('utf-8')
    def __repr__(self):
        return "{}( _dataSet={!r}, _alt={!r}, _cntry={!r}, _reg{!r} )".format(
               self.__class__.__name__,
               self.DataSet,
               self.Altitude,
               self.Country,
               self.Region)

#-------------------------------------------------------------------------------
#### For reading the IDF File  #####
class IDF_Zone:
    def __init__(self, _idfObj):
        self.ZoneName = getattr(_idfObj, 'Name')
        self.InfiltrationACH50 = None
        self.Volume_Gross = None
        self.FloorArea_Gross = False
        self.Volume_Vn50 = None
        self.TFA = None
        
    def __unicode__(self):
        return u'An IDF Zone Object: {}'.format(self.ZoneName)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_ZoneList:
    def __init__(self, _idfObj):
        self.Name = getattr(_idfObj, 'Name')
        
        for eachKey in _idfObj.__dict__.keys():
            if 'Zone ' in eachKey:
                setattr(self, eachKey, getattr(_idfObj, eachKey) )
                
    def __unicode__(self):
        return u'An IDF ZoneList Object: {}'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_ZoneInfilFlowRate:
    def __init__(self, _idfObj):
        self.Name = getattr(_idfObj, 'Name')
        self.ZoneName = getattr(_idfObj, 'Zone or ZoneList Name')
        self.ScheduleName = getattr(_idfObj, 'Schedule Name')
        self.DesignFlowRate = getattr(_idfObj, 'Design Flow Rate {m3/s}')
        self.FlowRatePerFloorArea = getattr(_idfObj, 'Flow per Zone Floor Area {m3/s-m2}')
        self.FlowRatePerSurfaceArea =  getattr(_idfObj, 'Flow per Exterior Surface Area {m3/s-m2}')
        self.ACH =  getattr(_idfObj, 'Air Changes per Hour {1/hr}')
        
    def __unicode__(self):
        return u'An IDF Zone Infiltration Flow Rate Object: {}'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_Obj_building:
    # Basic Building Data and North Orientation information
    
    def __init__(self, _idfObj):
        self.BldgName = getattr(_idfObj, 'Name')
        self.NorthAngle = getattr(_idfObj, 'North Axis {deg}') # in Degrees off North (North=0 ,East=90, etc..)
        self.NorthVector = self.calcNorthAnglefromVec(self.NorthAngle) # Calc degrees from the Vector
    
    def calcNorthAnglefromVec(self, _northAngle):
        _northAngle = float(_northAngle) * -1 # *-1 to go clockwise?
        northVec = rs.VectorRotate((0,1,0), float(_northAngle), (0,0,1)) # Vector to Rotate(Y), Angle(Deg), RotationAxis(Z)
        return northVec
    
    def __unicode__(self):
        return u'The Building Data Object: [{}]'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_Obj_MaterialLayer:
    # For holding onto Params for 
    # 'Material' and 'Material:AirGap' objects
    
    def __init__(self, _idfObj, noMass=False):
        self.Name = getattr(_idfObj, 'Name').replace('__Int__', '')
        self.MatType = getattr(_idfObj, 'objName')
        
        # Set object defaults
        self.LayerThickness = None
        self.LayerConductivity = None
        self.LayerConductance = None
        self.Type = None
        
        if noMass == False:
            # Get the actual data from the IDF Object to set up the rest
            self.getLayerData(_idfObj)
            self.setLayerData()
        elif noMass == True:
            self.getNoMassData(_idfObj)
    
    def getNoMassData(self, _idfObj):
        # Get all the relevant data from the IDF Object
        for eachKey in _idfObj.__dict__.keys():
            if 'Thermal Resistance {m2-K/W}' in eachKey:
                self.LayerConductance = 1 / float(getattr(_idfObj, 'Thermal Resistance {m2-K/W}'))
                self.LayerThickness = 1
                self.LayerConductivity = self.LayerConductance
    
    def getLayerData(self, _idfObj):
        # Get all the relevant data from the IDF Object
        for eachKey in _idfObj.__dict__.keys():
            if 'Thickness {m}' in eachKey:
                self.LayerThickness = getattr(_idfObj, 'Thickness {m}')
            elif 'Conductivity {W/m-K}' in eachKey:
                self.LayerConductivity = getattr(_idfObj, 'Conductivity {W/m-K}')
            elif 'Thermal Resistance {m2-K/W}' in eachKey:
                self.LayerConductance = 1 / float(getattr(_idfObj, 'Thermal Resistance {m2-K/W}'))
    
    def setLayerData(self):
        # Sort out the layer conductances/Resistances (m2-k/W)
        if 'AirGap' in self.MatType:
            self.LayerThickness = 1
        
        self.LayerThickness = float(self.LayerThickness) if self.LayerThickness else 0.1 # Apply default thickness if none
        
        if self.LayerConductance == None:
            self.LayerConductance = float(self.LayerConductivity) * float(self.LayerThickness)
        elif self.LayerConductance != None and self.LayerConductivity == None:
            print self.Name,  self.LayerThickness
            self.LayerConductivity = float(self.LayerConductance) / float(self.LayerThickness)
        
    def __unicode__(self):
        return u'EnergyPlus Material Params: [{}]'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_Obj_MaterialWindowSimple:
    # For holding onto Params for
    # WindowMaterial:SimpleGlazingSystem Objects
    
    def __init__(self, _idfObj):
        self.Name = getattr(_idfObj, 'Name' )
        self.uValue = getattr(_idfObj, 'U-Factor {W/m2-K}' )
        self.gValue = getattr(_idfObj, 'Solar Heat Gain Coefficient' )
        self.VT = getattr(_idfObj, 'Visible Transmittance' )
    
    def __unicode__(self):
        return u'EnergyPlus WindowMaterial:Simple Params: [{}]'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_Obj_MaterialWindowGlazing:
    # For holding onto Params for
    # WindowMaterial:Glazing Objects
    
    def __init__(self, _idfObj):
        self.Name = getattr(_idfObj, 'Name' )
        self.Thickness = float( getattr(_idfObj, 'Thickness {m}' ) )
        self.Conductivity = float( getattr(_idfObj, 'Conductivity {W/m-K}' ) )
        self.uValue = 1 / (self.Thickness/self.Conductivity)
        self.gValue = getattr(_idfObj, 'Solar Transmittance at Normal Incidence' )
        self.VT = getattr(_idfObj, 'Visible Transmittance at Normal Incidence' )
    
    def __unicode__(self):
        return u'EnergyPlus WindowMaterial:Glazing Params: [{}]'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_Obj_MaterialWindowGas:
    # For holding onto Params for
    # WindowMaterial:Gas Objects
    #https://www.engineersedge.com/heat_transfer/thermal-conductivity-gases.htm
    
    gasConductivities = {'Air': 0.0262,
                        'Argon': 0.0179,
                        'Krypton': 0.0095,
                        'Xenon': 0.0055
    }
    
    def __init__(self, _idfObj):
        self.Name = getattr(_idfObj, 'Name' )
        self.GasType = getattr(_idfObj, 'Gas Type' )
        self.Conductivity = self.gasConductivities[self.GasType]
        self.Thickness = float( (getattr(_idfObj, 'Thickness {m}' ) ) )
        self.uValue = 1 / (self.Thickness / self.Conductivity)
    
    def __unicode__(self):
        return u'EnergyPlus WindowMaterial:Gas Params: [{}]'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_Obj_Construction:
    """
    For holding onto Params for EnergPlus 'Construction' objects (Assemblies)
    Note that when created, the new Object will clean up its name to remove any 
    '__Int_' flags and then set the 'IntInsul' attribute as appropriate
    """
    def __init__(self, _idfObj):
        self.Name = getattr(_idfObj, 'Name')
        self.getLayerNames(_idfObj)
        self.checkInteriorInsul()
    
    def checkInteriorInsul(self):
        if '__Int__' in self.Name:
            self.IntInsul = 'x'
            self.Name = self.Name.replace('__Int__', '')
        else:
            self.IntInsul = None
            
    def getLayerNames(self, _idfObj):
        # Set the self.Layers list of names
        self.Layers = []
        self.LayerNames = []
        
        for eachKey in _idfObj.__dict__.keys():
            if 'Layer' in eachKey:
                layerNum = eachKey
                layerName = getattr(_idfObj, eachKey)
                layerName = layerName.replace('__Int__', '')
                self.Layers.append( [layerNum, layerName]  )
                self.LayerNames.append(layerName)
    
    def __unicode__(self):
        return u'EnergyPlus Construction Params: [{}]'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_Obj_surfaceWindow:
    # For holding onto Params for
    # FenestrationSurface:Detailed Objects
    
    def __init__(self, _idfObj, _winSimpleMat, _wShadFac, _sShadFac):
        self.Quantity = 1
        self.Name = getattr(_idfObj, 'Name')
        self.Dims = phpp_GetWindowSize(  phpp_geomFromVerts(_idfObj)[1]  )
        self.Width = self.Dims[0]
        self.Height = self.Dims[1]
        self.HostSrfc = getattr(_idfObj, 'Building Surface Name')
        self.winterShadingFac = _wShadFac
        self.summerShadingFac = _sShadFac
        self.EPConstuctionName = getattr(_idfObj, "Construction Name")
        
        # This will take in an EP 'WindowMaterial:SimpleGlazingSystem' and build PHPP style frame / glass
        self.setPHPPConstruction(self.EPConstuctionName, _winSimpleMat)
    
    def setPHPPConstruction(self, _constructionName, _winSimpleMat, _installs=[1,1,1,1]):
        # Sets the PHPP Style Frame, Glass and Installs 
        self.Type_Glass = PHPP_Glazing(
                _constructionName, # Name
                _winSimpleMat.gValue, # g-Value
                _winSimpleMat.uValue # U-Value
                )
        
        self.Type_Frame = PHPP_Frame(
                _constructionName, # Name
                [_winSimpleMat.uValue, _winSimpleMat.uValue, _winSimpleMat.uValue, _winSimpleMat.uValue], # Frame U-Values
                [0.12, 0.12, 0.12, 0.12], # Frame Widths (Default to 0.1m)
                [0.00, 0.00, 0.00, 0.00], # Psi-Glazing Edge Defaults
                [0.00, 0.00, 0.00, 0.00] # Psi-Installs Defaults)
                ) 
        
        self.Installs = PHPP_Window_Install(_installs)
    
    def __unicode__(self):
        return u'FenestrationSurface:Detailed: [{}]'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_Obj_surfaceOpaque:
    # For holding onto Params for
    # BuildingSurface:Detailed Objects
    
    def __init__(self, _idfObj, _northAngle):
        self.Name = getattr(_idfObj, 'Name')
        self.AssemblyName = getattr(_idfObj, 'Construction Name')
        self.srfcType = getattr(_idfObj, 'Surface Type')
        self.exposure = getattr(_idfObj, 'Outside Boundary Condition')
        self.HostZoneName = getattr(_idfObj, 'Zone Name')
        self.findGroupNumber(self.srfcType, self.exposure)
        self.getGeometryData(_idfObj, _northAngle)
    
    def getGeometryData(self, idfObj, _northAngle):
        # Build the Geometry from the Vertex points
        self.Boundary, self.Srfc, self.SurfaceArea, self.Centroid, self.NormalVector = phpp_geomFromVerts(idfObj)
        
        # Find the Rotation off North Vector
        self.AngleFromNorth = phpp_calcNorthAngle(self.NormalVector, _northAngle)
        
        # Find the Rotation off Horizontal
        self.AngleFromHoriz = ghc.Degrees(ghc.Angle(self.NormalVector, ghc.UnitZ(1)))[0]
        
        # Use Defaults at this time.
        # Someday calc the shading factors and have inputs for the rest?
        self.Factor_Shading = 0.5 # Default
        self.Factor_Absorptivity = 0.6  # Default
        self.Factor_Emissivity = 0.9   # Default
    
    def findGroupNumber(self, _srfcType, _exposureType):
        # Figure out the 'Group Number' for PHPP based on the EP Exposure type
        if _exposureType == 'Surface':
            pass
        elif _srfcType == 'Wall' and _exposureType == 'Outdoors':
            self.GroupNum = 8
        elif _srfcType == 'Wall' and _exposureType == 'Ground':
            self.GroupNum = 9
        elif _srfcType == 'Roof' and _exposureType == 'Outdoors':
            self.GroupNum = 10
        elif _srfcType == 'Floor' and _exposureType == 'Ground':
            self.GroupNum = 11
        elif _srfcType == 'Floor' and _exposureType == 'Outdoors':
            self.GroupNum = 12
        elif _exposureType == 'Adiabatic':
            self.GroupNum = 18
        else:
            self.GroupNum = 13
            print _srfcType, _exposureType
            groupWarning = "Couldn't figure out the Group Number for surface '{}'?".format(self.Name)
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, groupWarning)
    
    def __unicode__(self):
        return u'EnergyPlus BuildingSurface:Detailed Params: [{}]'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

class IDF_Obj_location:
    def __init__(self):
        self.Name = getattr(_idfObj, 'Name')
        self.Latitude = getattr(_idfObj, 'Latitude {deg}')
        self.Longitude = getattr(_idfObj, 'Longitude {deg}')
        self.TimeZone = getattr(_idfObj, 'Time Zone {hr}')
        self.Elevation = getattr(_idfObj, 'Elevation {m}')
    
    def __unicode__(self):
        return u'EnergyPlus Location Params: [{}]'.format(self.Name)
    
    def __str__(self):
        return unicode(self).encode('utf-8')

#-------------------------------------------------------------------------------
# For the main Excel Object Writer #
class PHPP_XL_Obj:
    """ A holder for an Excel writable datapoint with a worksheet, range and value """
    
    def __init__(self, _shtNm, _rangeAddress, _val):
        """
        Args:
            _shtNm (str): The Name of the Worksheet to write to
            _rangeAddress (str): The Cell Range (A1, B12, etc...) to write to on the Worksheet
            _val (str): The Value to write to the Cell Range (Value2)
        """
        self.Worksheet = _shtNm
        self.Range = _rangeAddress
        self.Value = _val
    
    def __unicode__(self):
        return u"PHPP Obj | Worksheet: {self.Worksheet}  |  Cell: {self.Range}  |  Value: {self.Value}".format(self=self)
    
    def __str__(self):
        return unicode(self).encode('utf-8')
    
    def __repr__(self):
       return "{}( _nm={!r}, _shtNm={!r}, _rangeAddress={!r}, _val={!r}".format(
               self.__class__.__name__,
               self.Worksheet,
               self.Range,
               self.Value )

####################################
# Add the Classes to the Scriptcontext
# Misc Utility Defs
sc.sticky['Preview'] = preview
sc.sticky['idf2ph_rhDoc'] = idf2ph_rhDoc

# Data
sc.sticky['phpp_ClimateData'] = getClimateData()

# PHPP Conversion Defs
sc.sticky['phpp_calcNorthAngle'] = phpp_calcNorthAngle
sc.sticky['phpp_GetWindowSize'] = phpp_GetWindowSize
sc.sticky['phpp_makeHBMaterial'] = phpp_makeHBMaterial
sc.sticky['phpp_makeHBMaterial_NoMass'] =  phpp_makeHBMaterial_NoMass
sc.sticky['phpp_makeHBMaterial_Opaque'] =  phpp_makeHBMaterial_Opaque
sc.sticky['phpp_makeHBConstruction'] = phpp_makeHBConstruction
sc.sticky['phpp_getWindowLibraryFromRhino'] = phpp_getWindowLibraryFromRhino
sc.sticky['phpp_createSrfcHBMatAndConst'] = phpp_createSrfcHBMatAndConst

# PHPP Object Classes
sc.sticky['PHPP_XL_Obj'] = PHPP_XL_Obj
sc.sticky['PHPP_WindowObject'] = PHPP_WindowObject
sc.sticky['PHPP_Glazing'] = PHPP_Glazing
sc.sticky['PHPP_Frame'] = PHPP_Frame
sc.sticky['PHPP_Window_Install'] = PHPP_Window_Install
sc.sticky['PHPP_Sys_Duct'] = PHPP_Sys_Duct
sc.sticky['PHPP_Sys_Ventilation'] = PHPP_Sys_Ventilation
sc.sticky['PHPP_Sys_VentUnit'] = PHPP_Sys_VentUnit
sc.sticky['PHPP_Sys_ExhaustVent'] = PHPP_Sys_ExhaustVent
sc.sticky['PHPP_TFA_Surface'] = PHPP_TFA_Surface
sc.sticky['PHPP_Room'] = PHPP_Room
sc.sticky['PHPP_RoomVolume'] = PHPP_RoomVolume
sc.sticky['PHPP_DHW_System'] = PHPP_DHW_System
sc.sticky['PHPP_DHW_usage'] = PHPP_DHW_usage
sc.sticky['PHPP_DHW_usage_NonRes'] = PHPP_DHW_usage_NonRes
sc.sticky['PHPP_DHW_RecircPipe'] = PHPP_DHW_RecircPipe
sc.sticky['PHPP_DHW_branch_piping'] = PHPP_DHW_branch_piping
sc.sticky['PHPP_DHW_tank'] = PHPP_DHW_tank
sc.sticky['PHPP_grnd_FloorElement'] = PHPP_grnd_FloorElement
sc.sticky['PHPP_grnd_SlabOnGrade'] = PHPP_grnd_SlabOnGrade
sc.sticky['PHPP_grnd_HeatedBasement'] = PHPP_grnd_HeatedBasement
sc.sticky['PHPP_grnd_UnheatedBasement'] = PHPP_grnd_UnheatedBasement
sc.sticky['PHPP_grnd_SuspendedFloor'] = PHPP_grnd_SuspendedFloor
sc.sticky['PHPP_ClimateDataSet'] = PHPP_ClimateDataSet

# IDF Object Classes
sc.sticky['IDF_Zone'] = IDF_Zone
sc.sticky['IDF_ZoneInfilFlowRate'] = IDF_ZoneInfilFlowRate
sc.sticky['IDF_ZoneList'] = IDF_ZoneList
sc.sticky['IDF_Obj_building'] = IDF_Obj_building
sc.sticky['IDF_Obj_MaterialLayer'] = IDF_Obj_MaterialLayer
sc.sticky['IDF_Obj_MaterialWindowSimple'] = IDF_Obj_MaterialWindowSimple
sc.sticky['IDF_Obj_MaterialWindowGlazing'] = IDF_Obj_MaterialWindowGlazing
sc.sticky['IDF_Obj_MaterialWindowGas'] = IDF_Obj_MaterialWindowGas
sc.sticky['IDF_Obj_Construction'] = IDF_Obj_Construction
sc.sticky['IDF_Obj_surfaceWindow'] = IDF_Obj_surfaceWindow
sc.sticky['IDF_Obj_surfaceOpaque'] = IDF_Obj_surfaceOpaque
sc.sticky['IDF_Obj_location'] = IDF_Obj_location

# Assignment Dictionaries
sc.sticky['IDF2PHPP_UDdict_glazing'] = {}
sc.sticky['IDF2PHPP_UDdict_frames'] = {}