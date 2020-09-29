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
Set the parameters for a Air- or Water-Source Heat Pump (HP). Sets the values on the 'HP' worksheet.
-
EM September 29, 2020
    Args:
        _unit_name: The name for the heat pump unit
        _source: The Heat Pump exterior 'source'. Input either:
> "1-Outdoor air"
> "2-Ground water"
> "3-Ground probes"
> "4-Horizontal ground collector"
        _fromPHPP: <Optional> If you prefer, you can just grab all the values from a PHPP or the 'Heat Pump Tool' from PHI and input here as a single element. Paste the values into a multiline component and connect to this input. The values will be separated by an invisible tab stop ("\t") which will be used to split the lines into the various columns. This is just an optional input to help make it easier. You can alwasys just use the direct entry inputs below if you prefer.
        _temps_source: (Deg C) List of test point values. List length should match the other inputs.
        _temps_sink: (Deg C) List of test point values. List length should match the other inputs.
        _heating_capacities: (kW) List of test point values. List length should match the other inputs.
        _COPs: (W/W) A list of the COP values at the various test points (source/sink). List length should match the other inputs.
        dt_sink_: (Deg C, Default = 5) A single temperature for the difference in sink. If using the 'Heat Pump Tool' from PHI, set this value to 0.
    Returns:
        heat_pump_: A Heat Pump object. Connect this output to the 'hp_heating_' input on a 'Heating/Cooling System' component.
"""

ghenv.Component.Name = "BT_Heating_ASHP"
ghenv.Component.NickName = "Heating | HP"
ghenv.Component.Message = 'SEP_29_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import scriptcontext as sc
import Grasshopper.Kernel as ghK

# Classes and Defs
preview = sc.sticky['Preview']

#-------------------------------------------------------------------------------
class Heating_HP:
    def __init__(self, _nm=None, _src='1-Outdoor air', _tsrcs=[], _tsnks=[], _hcs=[], _cops=[], _dtSnk=0):
        self.Name = _nm if _nm else 'Unnamed Heat Pump'
        self.Source = _src
        self.T_Source = _tsrcs
        self.T_Sink = _tsnks
        self.HC = _hcs
        self.COP = _cops
        self.dT_Sink = _dtSnk
        
        self.check_length()
    
    def check_length(self):
        if len(self.T_Source) > 15:
            msg = 'PHPP limits the number of test points to 15. For now I will\n'\
            'only use the first 15 entries in "_temps_source"'
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning ,msg)
            self.T_Source = self.T_Source[0:15]
        
        if len(self.T_Sink) > 15:
            msg = 'PHPP limits the number of test points to 15. For now I will\n'\
            'only use the first 15 entries in "_temps_sink"'
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning ,msg)
            self.T_Sink = self.T_Sink[0:15]
        
        if len(self.HC) > 15:
            msg = 'PHPP limits the number of test points to 15. For now I will\n'\
            'only use the first 15 entries in "_heating_capacities"'
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning ,msg)
            self.HC = self.HC[0:15]
        
        if len(self.COP) > 15:
            msg = 'PHPP limits the number of test points to 15. For now I will\n'\
            'only use the first 15 entries in "_COPs"'
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning ,msg)
            self.COP = self.COP[0:15]
    
    def __unicode__(self):
        return u'A Heating: Heat Pump Params Object'
    def __str__(self):
        return unicode(self).encode('utf-8')
    def __repr__(self):
        return "{}( _nm={!r}, _src={!r}, _tsrcs={!r}, _tsnks={!r}, _hcs={!r}, _cops={!r}, _dtSnk={!r})".format(
                self.__class__.__name__,
                self.Name,
                self.Source,
                self.T_Source,
                self.T_Sink,
                self.HC,
                self.COP,
                self.dT_Sink)

def clean_source_input(_in):
    input_str = str(_in)
    
    if '1' in input_str:
        return '1-Outdoor air'
    elif '2' in input_str:
        return '2-Ground water'
    elif '3' in input_str:
        return '3-Ground probes'
    elif '4' in input_str:
        return '4-Horizontal ground collector'
    else:
        return '1-Outdoor air'

def cleanInputs(_in, _nm):
    out = []
    
    if isinstance(_in, list):
        for i, item in enumerate(_in):
            try:
                out.append(float(item))
            except Exception as e:
                msg = 'Error in list: "{}", item: {}\n'\
                'Cannot convert "{}" to a number?'.format(_nm, i, item)
                ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
                break
    
    if len(out) != 0:
        return out
    else:
        return _in

def check_input_lengths(*args):
    check_length = len(args[0])
    msg = 'Error: It looks like your input lists are not all the same length?\n'\
    'Check the input list lengths for _temp_source, _temps_sinks, _heating_capacities and _COPs inputs.'
    
    for list in args:
        if len(list) != check_length:
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg)
            return 0
    
    return 1

def parse_from_PHPP(_inputList):
    if not isinstance(_inputList, list):
        return None, None, None, None
    
    src, snk, hc, cop = [], [], [], []
    for line in _inputList:
        row = line.split('\t')
        src.append(row[0])
        snk.append(row[1])
        hc.append(row[2])
        cop.append(row[3])
    
    return src, snk, hc, cop

#-------------------------------------------------------------------------------
# First, try and use any copy/paste from PHPP
# Then, get any user inputs lists

source = clean_source_input(_source)

tsrcs, tsnks, hcs, cops = parse_from_PHPP(_fromPHPP)

tsrcs = _temps_source if len(_temps_source)!=0 else tsrcs
tsnks = _temps_sink if len(_temps_sink)!=0 else tsnks
hcs = _heating_capacities if len(_heating_capacities)!=0 else hcs
cops = _COPs if len(_COPs)!=0 else cops

tsrcs = cleanInputs(tsrcs, '_temps_source')
tsnks = cleanInputs(tsnks, '_temps_sink')
hcs = cleanInputs(hcs, '_heating_capacities')
cops = cleanInputs(cops, '_COPs')
dtSnk = 0 if dt_sink_ is None else float(dt_sink_)

list_lengths = check_input_lengths(_temps_source, _temps_sink, _heating_capacities, _COPs)

#-------------------------------------------------------------------------------
# Create the Heat Pump Object

heat_pump_ = Heating_HP(_unit_name, source, tsrcs, tsnks, hcs, cops, dtSnk)
preview(heat_pump_)