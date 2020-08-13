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
Saves the open workbook in an excel instance
-
Component by Jack Hymowitz, July 31, 2020

    Args:
        save: Set to true to save the workbook
        excel: The excel instance
    Returns:
        excel: The excel instance after the component runs
"""

ghenv.Component.Name = "BT_SaveXLWorkbook"
ghenv.Component.NickName = "Save XL Workbook"
ghenv.Component.Message = 'JUL_31_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "02 | IDF2PHPP"

from ghpythonlib.componentbase import executingcomponent as component
import Grasshopper, GhPython
import System

import Rhino
import rhinoscriptsyntax as rs
import Grasshopper.Kernel as ghK

class MyComponent(component):
    
    def RunScript(self, save, excel):
        if (save is None  or save) and excel:
            if not excel.save():
                msg1 = "Unable to save"
                ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg1)
        else:
            msg1 = "No Excel Instance or save set to false"
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg1)
        return excel

