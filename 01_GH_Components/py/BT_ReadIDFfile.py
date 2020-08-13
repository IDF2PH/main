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
Use this to Read and parse an IDF file (EnergyPlus). This will go thorugh the IDF and pull out all the 'Objects'. It finds the empty newlines in the .IDF to identify each 'new' object and create a new object for each using the standard '!-' marker to establish keys. Will create key/value for EACH key found. Can use the getattr() method for keys with spaces in the name
-
EM Mar. 26, 2020

    Args:
        _idfFileAddress: Input the path/file location of the .IDF file used for the EnergyPlus simulation. Connect to the 'idfFileAddress' output from the Honeybee 'exportToOpenStudio' Component.
    Returns:
        IDF_Objs_List: A list of IDF-Objects found in the source file containing all their relevant parameters. Connect this to the '_IDF_Objs_List' input on the 'IDF-->PHPP' component in order to create PHPP writable objects from these.
        surfaces_: A text preview of all the Opaque surface objects found in the IDF along with all their parameters
        fenestration_: A text preview of all the Fenestration objects found in the IDF along with all their parameters
        constuctions_: A text preview of all the EP-Construction objects found in the IDF along with all their parameters
        materials_: A text preview of all the EP-Material objects found in the IDF along with all their parameters
            
            
"""

ghenv.Component.Name = "BT_ReadIDFfile"
ghenv.Component.NickName = "Read IDF File"
ghenv.Component.Message = 'MAR_26_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "02 | IDF2PHPP"

from System import Object
from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path
from shutil import copyfile
import os

class IDF_Class:
    # A simple class to hold onto the IDF object data
    
    def __init__(self, _varName, *_initial_data):   
        """Setting up the IDF Class Object and bringing in all its attributes
        """
        self.objName = _varName
        for dictionary in _initial_data:
            for key in dictionary:
                if key != None and dictionary[key] != None:
                    setattr(self, key, dictionary[key])
    def __repr__(self):
        return "An IDF File object with all its Params"

def idfObjPreview(_obj):
    outputList = []
    
    outputList.append(_obj.__dict__.get('objName') + '::')
    for k, v in _obj.__dict__.items():
        outputList.append(' > {}: {}'.format(k, v) )
    outputList.append('-------')
    
    return outputList

# Clear out the temporary variables
file = ''
lines_clean = []
idfFilePath = None

if _idfFileAddress:
    idfFilePath = _idfFileAddress
elif resultFileAddress_:
    outputDir = os.path.split(resultFileAddress_)[0]
    idfFileName = 'in.idf'
    idfFilePath = os.path.join(outputDir, "in.idf")

##### Bring in the data from the IDF file
if idfFilePath: 
    # Make a copy of the IDF file as a text file
    newFile = idfFilePath.replace('.idf', '.txt')
    copyfile(idfFilePath, newFile)
    
    print('>>>Opening the file to read....')
    
    # Open file for reading
    file = open(newFile, 'r')
    
    # Read lines into variable called lines
    lines = file.readlines()
    
    print('>>>Read the file successfully. Closing file.')
    # Close the file
    file.close()

if file:
    # Clean up the file a bit (get rid of the line breaks)
    for line in lines:
        if line != '\n': # get rid of the empty lines
            newLine = line.strip() # Get ride of any newlines at end of str
            newLineAsList = newLine.split('!- ') # Break into pieces
            
            # Get rid of all the extra white spaces, remove commas
            newLineAsList_Clean = []
            for each in newLineAsList:
                each = each.replace(",", "")
                each = each.replace(";", "")
                each = each.rstrip()
                newLineAsList_Clean.append( each )
            
            lines_clean.append(newLineAsList_Clean)
    
    # Organize the date into Class blocks (break at each 'header')
    tempTree = DataTree[Object]()
    p = 0
    # Go through the imported Data and create a new branch for each item
    for i in range(len(lines_clean)):
        if len(lines_clean[i]) == 1:
            p += 1
        tempTree.Add(lines_clean[i], GH_Path(p))
    
    # Create all the IDF Class object blocks
    IDF_Objs_List = []
    for each in tempTree.Branches:
        tempDict = {}
        Name = ''
        for eachListItem in each:
            if len(eachListItem) == 1:
                objName = eachListItem[0]
            else:
                tempDict[eachListItem[1]] = eachListItem[0]
        
        IDF_Objs_List.append(IDF_Class(objName, tempDict))

# Output the preview items
surfaces_ = []
fenestration_ = []
constuctions_ = []
materials_ = []

if IDF_Objs_List != None:
    for each in IDF_Objs_List:
        if 'BuildingSurface' in each.__dict__.get('objName', None):
            surfaces_ =  surfaces_ + idfObjPreview(each)
        elif 'Fenestration' in each.__dict__.get('objName', None):
            fenestration_ =  fenestration_ + idfObjPreview(each)
        elif 'Construction' in each.__dict__.get('objName', None):
            constuctions_ =  constuctions_ + idfObjPreview(each)
        elif 'Material' in each.__dict__.get('objName', None):
            materials_ =  materials_ + idfObjPreview(each)

