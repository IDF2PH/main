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
Takes in the E+ objects from the IDF-->PHPP Component and creates simplied
Excel-ready objects for writing to the PHPP
Each 'excel-ready' object has a Value, a Cell Range ('A4', 'BB56', etc...) and a Sheet Name
-
EM August 11, 2020

    Args:
        _PHPPObjs: A DataTree of the PHPP Objects to write out to Excel. Connect to the 'PHPPObjs_' in the 'IDF->PHPP Objs' Component.
        zonesInclude_: <Optional> Input the Zone Name or a list of Zone Names to output to a PHPP document. If no input, all zones found in the IDF will be output to a single PHPP Excel document. 
        zoneExclude_: <Optional> Pass in a list of string values to filter out certain zones by name. If the zone name includes the string anywhere in its name, it will be removed from the set to output.
        tfa_: <Optional> Input either a list of values (numbers) or a list of geometry representing the TFA (Treated Floor Area) of the building (m2). Set input as 'From Zone Geometry' and it will try and read any TFA from Zone's Room Objects as well. Direct input will take precedence over any Zone Rooms though. Leave blank for no TFA output to the PHPP
        thermalBridges_: <Optional> Input of Thermal Bridge Objects to write to the PHPP
        udRowStarts_: <Optional> Input a list of string values for any non-standard starting positions (rows) in your PHPP. This might be neccessary if you have modified your PHPP from the normal one you got originally. For instance, if you added new rows to the PHPP in  order to add more rooms (Additional Ventilation) or surfaces (Areas) or that sort of thing. To set the correct values here, input strings in the format " Worksheet Name, Start Key: New Start Row " - so use commas to separate the levels of the dict, then a semicolon before the value you want to input. Will accept multiline strings for multiple value resets.
        Enter any of the following valid Start Rows:
            -  Additional Ventilation, Rooms: ## (Default=56)
            -  Additional Ventilation, Vent Unit Selection: ## (Default=97)
            -  Additional Ventilation, Vent Ducts: ## (Default=127)
            -  Components, Ventilator: ## (Default=15)
            -  Areas, TB: ## (Default=145)
            -  Areas, Surfaces: ## (Default=41)
            -  Electricity non-res, Lighting: ## (Default= 19)
            -  Electricity non-res, Office Equip: ## (Default=62)
            -  Electricity non-res, Kitchen: ## (Default=77)
    Returns:
        toPHPP_Geom_: A DataTree of the final clean, Excel-Ready output objects. Each output object has a Worksheet-Name, a Cell Range, and a Value. Connect to the 'Geom_' input on the 'Write 2PHPP' Component to write to Excel.
"""

ghenv.Component.Name = "BT_CreateXLObj_Geom"
ghenv.Component.NickName = "Create Excel Obj - Geom"
ghenv.Component.Message = 'AUG_11_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "02 | IDF2PHPP"

import Grasshopper.Kernel as ghK
from string import ascii_uppercase
import itertools
from System import Object
from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path
import rhinoscriptsyntax as rs
import ghpythonlib.components as gh
import scriptcontext as sc
from collections import defaultdict
import statistics

# Classes and Defs
PHPP_XL_Obj = sc.sticky['PHPP_XL_Obj'] 
preview = sc.sticky['Preview']
PHPP_DHW_System = sc.sticky['PHPP_DHW_System']
PHPP_DHW_usage = sc.sticky['PHPP_DHW_usage'] 
PHPP_DHW_branch_piping = sc.sticky['PHPP_DHW_branch_piping']
PHPP_DHW_tank = sc.sticky['PHPP_DHW_tank']
PHPP_DHW_RecircPipe = sc.sticky['PHPP_DHW_RecircPipe']

def getUvalues(_inputBranch):
    uID_Count = 1
    uValueUID_Names = []
    uValuesConstructorStartRow = 10
    uValuesList = []
    print 'Creating the U-Values Objects...'
    for eachConst in _inputBranch:
        # for each Construction Assembly in the model....
        
        # Get the Construction's Name and the Materal Layers in the EP Model
        construcionNameEP = getattr(eachConst, 'Name')
        layers = sorted(getattr(eachConst, 'Layers'))
        intInsuFlag = eachConst.IntInsul if eachConst.IntInsul != None else ''
        
        # Filter out any of the Window Constructions
        isWindow = False
        opaqueMaterialNames = []
        for eachMat in _PHPPObjs.Branch(0):
            opaqueMaterialNames.append(eachMat.Name) # Get all the Opaque Construction Material Names
        
        # Check if the material matches any of the Opaque ones
        for eachLayer in layers:
            if eachLayer[1] in opaqueMaterialNames:
                eachLayer[1]
                break
            else:
                # If not... it must be a window (maybe?)
                isWindow = True
        
        if isWindow == True:
            pass
        else:
            # Fix the name to remove 'PHPP_CONST_'
            if 'PHPP_CONST_' in construcionNameEP:
                constName_clean = construcionNameEP.split('PHPP_CONST_')[1].replace('_', ' ')
            else:
                constName_clean = construcionNameEP.replace('_', ' ')
            
            # Create the list of User-ID Constructions to match PHPP
            uValueUID_Names.append('{:02d}ud-{}'.format(uID_Count, constName_clean) )
            
            # Create the Objects for the Header Piece (Name, Rsi, Rse)
            nameAddress = '{}{}'.format('M', uValuesConstructorStartRow + 1) # Construction Name
            rSi = '{}{}'.format('M', uValuesConstructorStartRow + 3) # R-surface-int
            rSe = '{}{}'.format('M', uValuesConstructorStartRow + 4) # R-surface-ext
            intIns = '{}{}'.format('S', uValuesConstructorStartRow + 1) # Interior Insulation Flag
            
            uValuesList.append( PHPP_XL_Obj('U-Values', nameAddress, constName_clean))
            uValuesList.append( PHPP_XL_Obj('U-Values', rSi, 0)) # For now, zero out
            uValuesList.append( PHPP_XL_Obj('U-Values', rSe, 0)) # For now, zero out
            if eachConst.IntInsul != None:
                uValuesList.append( PHPP_XL_Obj('U-Values', intIns, 'x'))
            
            # Create the actual Material Layers for PHPP U-Value
            layerCount = 0
            for layer in layers:
                # For each layer in the Construction Assembly...
                for eachMatLayer in  _PHPPObjs.Branch(0):
                    # See if the Construction's Layer material name matches one in the Materials list....
                    # If so, use those parameters from the Material Layer
                    if layer[1] == getattr(eachMatLayer, 'Name'):
                        # Filter out any MASSLAYERs
                        if layer[1] != 'MASSLAYER':
                            
                            # Clean the name
                            if 'PHPP_MAT_' in layer[1]:
                                layerMatName = layer[1].split('PHPP_MAT_')[1].replace('_', ' ')
                            else:
                                layerMatName = layer[1].replace('_', ' ')
                            
                            layerNum = layer[0]
                            layerMatCond = getattr(eachMatLayer, 'LayerConductivity')
                            layerThickness = getattr(eachMatLayer, 'LayerThickness')*1000 # Cus PHPP uses mm for thickness
                            
                            # Set up the Range tagets
                            layer1Address_L = '{}{}'.format('L', uValuesConstructorStartRow + 7 + layerCount) # Material Name
                            layer1Address_M = '{}{}'.format('M', uValuesConstructorStartRow + 7 + layerCount) # Conductivity
                            layer1Address_S = '{}{}'.format('S', uValuesConstructorStartRow + 7 + layerCount) # Thickness
                            
                            # Create the Layer Objects
                            uValuesList.append( PHPP_XL_Obj('U-Values', layer1Address_L, layerMatName))# Material Name
                            uValuesList.append( PHPP_XL_Obj('U-Values', layer1Address_M, layerMatCond)) # Conductivity
                            uValuesList.append( PHPP_XL_Obj('U-Values', layer1Address_S, layerThickness)) # Thickness
                            
                            layerCount+=1
            
            uID_Count += 1
            uValuesConstructorStartRow += 21
    
    return uValuesList, uValueUID_Names

def getComponents(_inputBranch):
    winComponentStartRow = 15
    frame_Count = 0
    glass_Count = 0
    winComponentsList = []
    glassNameDict = {}
    frameNameDict = {}
    
    print('Creating the Components:Window Objects...')
    
    for eachWin in _inputBranch:
        # For each PHPP Style Window Object in the model....
        
        ########## Glass ##########
        # Pull out the Glass info from the window
        gNm = getattr(eachWin.Type_Glass, 'Name')
        gV = getattr(eachWin.Type_Glass, 'gValue')
        uG = getattr(eachWin.Type_Glass, 'uValue')
        
        if gNm not in glassNameDict.keys():
            # Add the new glass type to the dict of UD Names:
            # ie: {'Ikon: SDH': '01ud-Ikon: SDH', ....}
            glassNameDict[gNm] = '{:02d}ud-{}'.format(glass_Count+1, gNm)
            
            # Set the glass range addresses
            Address_Gname = '{}{}'.format('IE', winComponentStartRow + glass_Count) # Name
            Address_Gvalue = '{}{}'.format('IF', winComponentStartRow + glass_Count) # g-Value
            Address_Uvalue = '{}{}'.format('IG', winComponentStartRow + glass_Count) # U-Value
            
            # Create the PHPP write Objects
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Gname, gNm))# Glass Type Name
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Gvalue, gV))# g-Value
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Uvalue, uG))# U-Value
            
            glass_Count +=1
            
        # Add the new PHPP UD Glass Name to the Window:Simple Object
        setattr(eachWin, 'UD_glass_Name', glassNameDict[gNm] )
        
        ########## Frames ##########
        # Get the Frame info
        fNm = getattr(eachWin.Type_Frame, 'Name')
        uF_L, uF_R, uF_B, uF_T  = eachWin.Type_Frame.uLeft, eachWin.Type_Frame.uRight, eachWin.Type_Frame.uBottom, eachWin.Type_Frame.uTop
        wF_L, wF_R, wF_B, wF_T  = eachWin.Type_Frame.fLeft, eachWin.Type_Frame.fRight, eachWin.Type_Frame.fBottom, eachWin.Type_Frame.fTop
        psiG_L, psiG_R, psiG_B, psiG_T  = eachWin.Type_Frame.psigLeft, eachWin.Type_Frame.psigRight, eachWin.Type_Frame.psigBottom, eachWin.Type_Frame.psigTop
        psiI_L, psiI_R, psiI_B, psiI_T  = eachWin.Type_Frame.psiInstLeft, eachWin.Type_Frame.psiInstRight, eachWin.Type_Frame.psiInstBottom, eachWin.Type_Frame.psiInstTop
        
        if fNm not in frameNameDict.keys():
            # Add the new frame type to the dict of UD Names:
            # ie: {'Ikon: SDH': '01ud-Ikon: SDH', ....}
            frameNameDict[fNm] = '{:02d}ud-{}'.format(frame_Count+1, fNm) # was glass_count????
            
            # Set the frame range address
            Address_Fname = '{}{}'.format('IL', winComponentStartRow + frame_Count)
            Address_Uf_Left = '{}{}'.format('IM', winComponentStartRow + frame_Count)
            Address_Uf_Right = '{}{}'.format('IN', winComponentStartRow + frame_Count)
            Address_Uf_Bottom = '{}{}'.format('IO', winComponentStartRow + frame_Count)
            Address_Uf_Top = '{}{}'.format('IP', winComponentStartRow + frame_Count)
            Address_W_Left = '{}{}'.format('IQ', winComponentStartRow + frame_Count)
            Address_W_Right = '{}{}'.format('IR', winComponentStartRow + frame_Count)
            Address_W_Bottom = '{}{}'.format('IS', winComponentStartRow + frame_Count)
            Address_W_Top = '{}{}'.format('IT', winComponentStartRow + frame_Count)
            Address_Psi_g_Left = '{}{}'.format('IU', winComponentStartRow + frame_Count)
            Address_Psi_g_Right = '{}{}'.format('IV', winComponentStartRow + frame_Count)
            Address_Psi_g_Bottom = '{}{}'.format('IW', winComponentStartRow + frame_Count)
            Address_Psi_g_Top = '{}{}'.format('IX', winComponentStartRow + frame_Count)
            Address_Psi_I_Left = '{}{}'.format('IY', winComponentStartRow + frame_Count)
            Address_Psi_I_Right = '{}{}'.format('IZ', winComponentStartRow + frame_Count)
            Address_Psi_I_Bottom = '{}{}'.format('JA', winComponentStartRow + frame_Count)
            Address_Psi_I_Top = '{}{}'.format('JB', winComponentStartRow + frame_Count)
            
            # Create the PHPP Objects for the Frames
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Fname, fNm))# Frame Type Name
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Uf_Left, uF_L)) # Frame Type U-Values
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Uf_Right, uF_R))
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Uf_Bottom, uF_B))
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Uf_Top, uF_T))
            
            winComponentsList.append( PHPP_XL_Obj('Components', Address_W_Left, wF_L)) # Frame Type Widths
            winComponentsList.append( PHPP_XL_Obj('Components', Address_W_Right, wF_R))
            winComponentsList.append( PHPP_XL_Obj('Components', Address_W_Bottom, wF_B))
            winComponentsList.append( PHPP_XL_Obj('Components', Address_W_Top, wF_T))
            
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Psi_g_Left, psiG_L)) # Frame Type Psi-Glazing
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Psi_g_Right, psiG_R))
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Psi_g_Bottom, psiG_B))
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Psi_g_Top, psiG_T))
            
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Psi_I_Left, psiI_L)) # Frame Type Psi-Installs
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Psi_I_Right, psiI_R))
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Psi_I_Bottom, psiI_B))
            winComponentsList.append( PHPP_XL_Obj('Components', Address_Psi_I_Top, psiI_T))
            
            frame_Count +=1
            
        # Add the PHPP UD Frame Name to the Window:Simple Object
        setattr(eachWin, 'UD_frame_Name', frameNameDict[fNm] )
    
    return winComponentsList

def getAreas(_inputBranch, _zones):
    areasRowStart = 41
    areaCount = 0
    uID_Count = 1
    areasList = []
    surfacesIncluded = []
    print "Creating the 'Areas' Objects..."
    for surface in _inputBranch:
        # for each Opaque Surface in the model....
        
        # First, see if the Surface should be included in the output
        for eachZoneName in _zones:
            if surface.HostZoneName == eachZoneName:
                includeSurface = True
                break
            else:
                includeSurface = False
        
        if includeSurface:
            # Get the Surface Parameters
            nm = getattr(surface, 'Name')
            groupNum = getattr(surface, 'GroupNum')
            quantity = 1
            surfaceArea = getattr(surface, 'SurfaceArea')
            assemblyName = getattr(surface, 'AssemblyName').replace('_', ' ') 
            angleFromNorth = getattr(surface, 'AngleFromNorth')
            angleFromHoriz = getattr(surface, 'AngleFromHoriz')
            shading = getattr(surface, 'Factor_Shading')
            abs = getattr(surface, 'Factor_Absorptivity')
            emmis = getattr(surface, 'Factor_Emissivity')
            
            # Find the right UID name (with the numeric prefix)
            for uIDName in uValueUID_Names:
                if assemblyName in uIDName[5:] or uIDName[5:] in assemblyName: # compare to slice without prefix
                    assemblyName = uIDName
            
            # Setup the Excel Address Locations
            Address_Name = '{}{}'.format('L', areasRowStart + areaCount)
            Address_GroupNum = '{}{}'.format('M', areasRowStart + areaCount)
            Address_Quantity = '{}{}'.format('P', areasRowStart + areaCount)
            Address_Area = '{}{}'.format('V', areasRowStart + areaCount)
            Address_Assembly = '{}{}'.format('AC', areasRowStart + areaCount)
            Address_AngleNorth = '{}{}'.format('AG', areasRowStart + areaCount)
            Address_AngleHoriz = '{}{}'.format('AH', areasRowStart + areaCount)
            Address_ShadingFac = '{}{}'.format('AJ', areasRowStart + areaCount)
            Address_Abs = '{}{}'.format('AK', areasRowStart + areaCount)
            Address_Emmis = '{}{}'.format('AL', areasRowStart + areaCount)
            
            areasList.append( PHPP_XL_Obj('Areas', Address_Name, nm))# Surface Name
            areasList.append( PHPP_XL_Obj('Areas', Address_GroupNum, groupNum))# Surface Group Number
            areasList.append( PHPP_XL_Obj('Areas', Address_Quantity, quantity))# Surface Quantity
            areasList.append( PHPP_XL_Obj('Areas', Address_Area, surfaceArea))# Surface Area (m2)
            areasList.append( PHPP_XL_Obj('Areas', Address_Assembly, assemblyName))# Assembly Type Name
            areasList.append( PHPP_XL_Obj('Areas', Address_AngleNorth, angleFromNorth))# Orientation Off North
            areasList.append( PHPP_XL_Obj('Areas', Address_AngleHoriz, angleFromHoriz))# Orientation Off Horizontal
            areasList.append( PHPP_XL_Obj('Areas', Address_ShadingFac, shading))# Shading Factor
            areasList.append( PHPP_XL_Obj('Areas', Address_Abs, abs))# Absorptivity
            areasList.append( PHPP_XL_Obj('Areas', Address_Emmis, emmis))# Emmissivity
            
            # Add the PHPP UD Surface Name to the Surface Object
            setattr(surface, 'UD_Srfc_Name', '{:d}-{}'.format(uID_Count, nm) )
            
            # Keep track of which Surfaces are included in the output
            surfacesIncluded.append(nm)
            
            uID_Count += 1
            areaCount += 1
    
    areasList.append( PHPP_XL_Obj('Areas', 'L19', 'Suspended Floor') )
    return areasList, surfacesIncluded

def getThermalBridges(_inputBranch, _startRows):
    tb_RowStart = _startRows.get('Areas').get('TB')
    tb_List = []
    print "Creating the 'Thermal Bridging' Objects..."
    for i, tb in enumerate(_inputBranch):
        # for each Thermal Bridge in the model....
        if tb.Name == 'Estimated':
            i = 0
        else:
            i = i+1
        
         # Setup the Excel Address Locations
        Address_Name = '{}{}'.format('L', tb_RowStart + i)
        Address_GroupNo = '{}{}'.format('M', tb_RowStart + i)
        Address_Quantity = '{}{}'.format('P', tb_RowStart + i)
        Address_Length = '{}{}'.format('R', tb_RowStart + i)
        Address_PsiValue = '{}{}'.format('X', tb_RowStart + i)
        
        tb_List.append( PHPP_XL_Obj('Areas', Address_Name, tb.Name))
        tb_List.append( PHPP_XL_Obj('Areas', Address_GroupNo, tb.GroupNo))
        tb_List.append( PHPP_XL_Obj('Areas', Address_Quantity, 1))
        tb_List.append( PHPP_XL_Obj('Areas', Address_Length, tb.Length))
        tb_List.append( PHPP_XL_Obj('Areas', Address_PsiValue, tb.PsiValue))
    
    return tb_List

def getWindows(_inputBranch, _surfacesIncluded, _srfcBranch):
    windowsRowStart = 24
    windowsCount = 0
    winSurfacesList = []
    
    print "Creating the 'Windows' Objects..."
    for window in _inputBranch:
        # for each Window Surface Object in the model....
        # Get the window's basic params
        quant = getattr(window, 'Quantity')
        nm = getattr(window, 'Name')
        w = getattr(window, 'Width')
        h = getattr(window, 'Height')
        host = getattr(window, 'HostSrfc')
        glassType = getattr(window.Type_Glass, 'Name')
        frameType = getattr(window.Type_Frame, 'Name')
        glassTypeUD = getattr(window, 'UD_glass_Name')
        frameTypeUD = getattr(window, 'UD_frame_Name')
        variantType = getattr(window, 'Type_Variant', 'a')
        
        # See if the Window should be included in the output
        includeWindow = False
        for eachSurfaceName in _surfacesIncluded:
            if eachSurfaceName == host:
                includeWindow = True
                break
            else:
                includeWindow = False
        
        if includeWindow:
            # Find the Window's Host Surface UD
            for srfc in _srfcBranch:
                if host == getattr(srfc, 'Name'):
                    hostUD = getattr(srfc, 'UD_Srfc_Name')
           
           # Get the Window Range Addresses
            Address_varType = '{}{}'.format('F', windowsRowStart + windowsCount)
            Address_winQuantity = '{}{}'.format('L', windowsRowStart + windowsCount)
            Address_winName = '{}{}'.format('M', windowsRowStart + windowsCount)
            Address_w = '{}{}'.format('Q', windowsRowStart + windowsCount)
            Address_h = '{}{}'.format('R', windowsRowStart + windowsCount)
            Address_hostName = '{}{}'.format('S', windowsRowStart + windowsCount)
            Address_glassType = '{}{}'.format('T', windowsRowStart + windowsCount)
            Address_frameType = '{}{}'.format('U', windowsRowStart + windowsCount)
            Address_install_Left = '{}{}'.format('AA', windowsRowStart + windowsCount)
            Address_install_Right = '{}{}'.format('AB', windowsRowStart + windowsCount)
            Address_install_Bottom = '{}{}'.format('AC', windowsRowStart + windowsCount)
            Address_install_Top = '{}{}'.format('AD', windowsRowStart + windowsCount)
            
            # Create the PHPP Window Object
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_varType, variantType)) # Quantity
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_winQuantity, quant)) # Quantity
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_winName, nm)) # Name
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_w, w)) # Width
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_h, h)) # Height
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_hostName, hostUD)) # Host Name
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_glassType, glassTypeUD)) # Glass UD Name
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_frameType, frameTypeUD)) # Frame UD Name
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_install_Left, window.Installs.Inst_L)) # Install Condition Left
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_install_Right, window.Installs.Inst_R)) # Install Condition Right
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_install_Bottom, window.Installs.Inst_B)) # Install Condition Bottom
            winSurfacesList.append( PHPP_XL_Obj('Windows', Address_install_Top, window.Installs.Inst_T)) # Install Condition Top
            
            windowsCount += 1
            
    return winSurfacesList

def getShading(_inputBranch, _surfacesIncluded):
    shadingRowStart = 17
    shadingCount = 0
    shadingList = []
    print "Creating the 'Shading' Objects..."
    for window in _inputBranch:
        # for each Window IDF Object in the model....
        
        # First, see if the Window should be included in the output
        host = getattr(window, 'HostSrfc')
        includeWindow = False
        for eachSurfaceName in _surfacesIncluded:
            if eachSurfaceName == host:
                includeWindow = True
                break
            else:
                includeWindow = False
        
        if includeWindow:
            # Get the Window's shading Params
            try:
                winterShadingFactor = getattr(window, 'winterShadingFac')
                summerShadingFactor = getattr(window, 'summerShadingFac')
            except:
                winterShadingFactor = 0.75
                summerShadingFactor = 0.75
            
            # Create the PHPP Objects
            Address1_start = '{}{}'.format('AF', shadingRowStart + shadingCount)
            Address1_end = '{}{}'.format('AG', shadingRowStart + shadingCount)
            shadingList.append( PHPP_XL_Obj( 'Shading', Address1_start , winterShadingFactor))
            shadingList.append( PHPP_XL_Obj( 'Shading', Address1_end , summerShadingFactor))
            
            shadingCount += 1
        
    return shadingList

def getTFA(tfaFromUser, tfaBranch, _zones):
    ##########################################
    ##############     TFA     ###############
    tfa = []
    
    if len(tfaFromUser)>0:
        if tfaFromUser[0] == 'From Zone Geometry':
            # Pulling data in from the HB Zone Objects
            print("Trying to find any Honeybee Zone Room TFA info...")
            try:
                tfaSurfaceAreas = [0]
                for each in tfaBranch:
                    # First, see if the Surface should be included in the output
                    for zoneName in _zones:
                        if each.HostZoneName == zoneName:
                            includeRoom = True
                            break
                        else:
                            includeRoom = False
                    
                    if includeRoom:
                        # Get the room's TFA info
                        roomTFA = each.FloorArea_TFA
                        tfaSurfaceAreas.append( roomTFA )
                # Total up the TFA Areas for output
                tfaTotal = sum(tfaSurfaceAreas)
                tfa.append( PHPP_XL_Obj('Areas', 'V34', tfaTotal )) # TFA (m2)
            except:
                pass
        else:
            print("Determining the TFA from user input...")
            tfaSurfaceAreas = [0]
            for each in tfaFromUser:
                try:
                    # if its a number or list of numbers
                    tfaSurfaceAreas.append(float(each))
                except:
                    # if it isn't a number, try bringing it in as geometry
                    area = gh.Area(each)[0] # Position 0 is the area in m2
                    tfaSurfaceAreas.append(float(area))
            
            if sum(tfaSurfaceAreas) != 0:
                tfaTotal = sum(tfaSurfaceAreas)
                tfa.append( PHPP_XL_Obj('Areas', 'V34', tfaTotal )) # TFA (m2)
    
    return tfa

def getAddnlVentRooms(_inputBranch, _ventSystems, _zones, _startRows):
    print "Creating 'Additional Ventilation' Rooms... "
    addnlVentRooms = []
    ventUnitsUsed = []
    roomRowStart = _startRows.get('Additional Ventilation').get('Rooms', 57)
    ventUnitRowStart = _startRows.get('Additional Ventilation').get('Vent Unit Selection', 97)
    ventSystemsInlcuded = set()
    i = 0
    
    for i, roomObj in enumerate(_inputBranch):
        
        # First, see if the Room should be included in the output
        for zoneName in _zones:
            if roomObj.HostZoneName == zoneName:
                includeRoom = True
                break
            else:
                includeRoom = False
        
        if includeRoom:
            ventSystemsInlcuded.add(roomObj.VentSystemName)
            
            # Try and sort out the Room's Ventilation airflow and schedule if there is any
            try:
                roomAirFlow_sup = roomObj.getVsup()
                roomAirFlow_eta = roomObj.getVeta()
                roomAirFlow_trans = roomObj.getVtrans()
            except:
                roomAirFlow_sup = roomObj.getVsup()
                roomAirFlow_eta = roomObj.getVeta()
                roomAirFlow_trans = roomObj.getVtrans()
            
            try:
                ventUnitName = roomObj.VentUnitName
                ventSystemName = roomObj.VentSystemName
            except:
                ventUnitName = '97ud-Default HRV unit'
                ventSystemName = 'Vent-1'
            
            # Get the Ventilation Schedule from the room if it has any
            try:
                speed_high = roomObj.phppVentSched.speed_high
                time_high = roomObj.phppVentSched.time_high
                speed_med = roomObj.phppVentSched.speed_med
                time_med = roomObj.phppVentSched.time_med
                speed_low = roomObj.phppVentSched.speed_low
                time_low = roomObj.phppVentSched.time_low
            except:
                speed_high = 1
                time_high = 1
                speed_med = None
                time_med = None
                speed_low = None
                time_low = None
            
            address_Amount = '{}{}'.format('D', roomRowStart + i)
            address_Name = '{}{}'.format('E', roomRowStart + i)
            address_VentAllocation = '{}{}'.format('F', roomRowStart + i)
            address_Area = '{}{}'.format('G', roomRowStart + i)
            address_RoomHeight = '{}{}'.format('H', roomRowStart + i)
            address_SupplyAirFlow = '{}{}'.format('J', roomRowStart + i)
            address_ExractAirFlow = '{}{}'.format('K', roomRowStart + i)
            address_TransferAirFlow = '{}{}'.format('L', roomRowStart + i)
            address_Util_hrs = '{}{}'.format('N', roomRowStart + i)
            address_Util_days = '{}{}'.format('O', roomRowStart + i)
            address_Holidays = '{}{}'.format('P', roomRowStart + i)
            
            address_ventSpeed_high = '{}{}'.format('Q', roomRowStart + i)
            address_ventTime_high = '{}{}'.format('R', roomRowStart + i) 
            address_ventSpeed_med = '{}{}'.format('S', roomRowStart + i)
            address_ventTime_med = '{}{}'.format('T', roomRowStart + i)
            address_ventSpeed_low = '{}{}'.format('U', roomRowStart + i)
            address_ventTime_low = '{}{}'.format('V', roomRowStart + i)
            
            ventMatchFormula = '=MATCH("{}",E{}:E{},0)'.format(ventSystemName, ventUnitRowStart, ventUnitRowStart+9)
            
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Amount, 1 ))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Name, '{}-{}'.format(roomObj.RoomNumber, roomObj.RoomName )))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_VentAllocation, ventMatchFormula ))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Area, roomObj.FloorArea_TFA ))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_RoomHeight, roomObj.RoomClearHeight ))
            
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_SupplyAirFlow, roomAirFlow_sup ))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ExractAirFlow, roomAirFlow_eta ))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_TransferAirFlow, roomAirFlow_trans ))
            
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Util_hrs, '24'))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Util_days, '7'))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Holidays,'0'))
            
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventSpeed_high, speed_high if speed_high else 1))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventTime_high, time_high if time_high else 1))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventSpeed_med,speed_med if speed_med else 1))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventTime_med, time_med if time_med else 0))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventSpeed_low,speed_low if speed_low else 0))
            addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventTime_low, time_low if time_low else 0))
            
            # Keep track of the names of the Vent units used
            ventUnitsUsed.append( ventUnitName )
    
    # Include any Exhaust Ventilation Objects that are found in any of the included Vent Systems
    rowCount = i+1
    for ventSystemDict in _ventSystems:
        for ventSystem in ventSystemDict.values():
            if ventSystem.SystemName in ventSystemsInlcuded:
                for exhaustVentObj in ventSystem.ExhaustObjs:
                    for mode in ['on', 'off']:
                        
                        address_Amount = '{}{}'.format('D', roomRowStart + rowCount)
                        address_Name = '{}{}'.format('E', roomRowStart + rowCount)
                        address_VentAllocation = '{}{}'.format('F', roomRowStart + rowCount)
                        address_Area = '{}{}'.format('G', roomRowStart + rowCount)
                        address_RoomHeight = '{}{}'.format('H', roomRowStart + rowCount)
                        address_SupplyAirFlow = '{}{}'.format('J', roomRowStart + rowCount)
                        address_ExractAirFlow = '{}{}'.format('K', roomRowStart + rowCount)
                        address_TransferAirFlow = '{}{}'.format('L', roomRowStart + rowCount)
                        address_Util_hrs = '{}{}'.format('N', roomRowStart + rowCount)
                        address_Util_days = '{}{}'.format('O', roomRowStart + rowCount)
                        address_Holidays = '{}{}'.format('P', roomRowStart + rowCount)
                            
                        address_ventSpeed_high = '{}{}'.format('Q', roomRowStart + rowCount)
                        address_ventTime_high = '{}{}'.format('R', roomRowStart + rowCount) 
                        address_ventSpeed_med = '{}{}'.format('S', roomRowStart + rowCount)
                        address_ventTime_med = '{}{}'.format('T', roomRowStart + rowCount)
                        address_ventSpeed_low = '{}{}'.format('U', roomRowStart + rowCount)
                        address_ventTime_low = '{}{}'.format('V', roomRowStart + rowCount)
                        
                        ventMatchFormula = '=MATCH("{}",E{}:E{},0)'.format(exhaustVentObj.Name, ventUnitRowStart, ventUnitRowStart+9)
                        
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Amount, 1 ))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Name, exhaustVentObj.Name +' [ON]' if mode=='on' else exhaustVentObj.Name +' [OFF]'))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_VentAllocation, ventMatchFormula ))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Area, '10'))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_RoomHeight, '2.5' ))
                        
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_SupplyAirFlow, exhaustVentObj.FlowRate_On if mode=='on' else exhaustVentObj.FlowRate_Off))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ExractAirFlow, exhaustVentObj.FlowRate_On if mode=='on' else exhaustVentObj.FlowRate_Off))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_TransferAirFlow, '0' ))
                        
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Util_hrs, exhaustVentObj.HrsPerDay_On if mode=='on' else 24 - float(exhaustVentObj.HrsPerDay_On)))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Util_days, exhaustVentObj.DaysPerWeek_On if mode=='on' else 7))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_Holidays, exhaustVentObj.Holidays))
                        
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventSpeed_high, 1))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventTime_high, 1))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventSpeed_med,0))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventTime_med, 0))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventSpeed_low, 0))
                        addnlVentRooms.append( PHPP_XL_Obj('Additional Vent', address_ventTime_low, 0))
                        
                        rowCount += 1
    
    return addnlVentRooms, ventUnitsUsed

def getAddnlVentSystems(_inputBranch, _ventUnitsUsed, _startRows):
    # Go through each Ventilation System passed in
    vent = []
    ventCompoRowStart = _startRows.get('Components').get('Ventilator')
    ventUnitRowStart = _startRows.get('Additional Ventilation').get('Vent Unit Selection')
    ventDuctsRowStart = _startRows.get('Additional Ventilation').get('Vent Ducts')
    ventCount = 0
    ductsCount = 0
    ductColCount = ord('Q')
    
    if len(_inputBranch)>0:
        print "Creating 'Additional Ventilation' Systems..."
        vent.append( PHPP_XL_Obj('Ventilation', 'H42', 'x') ) # Turn on Additional Vent
        vent.append( PHPP_XL_Obj('Additional Vent', 'F'+str(ventDuctsRowStart-11) , "=AVERAGE(Climate!E24, Climate!F24, Climate!N24, Climate!O24, Climate!P24") ) # External Average Temp
        
        for key in _inputBranch[0].keys():
            ventSystem = _inputBranch[0][key] 
            
            # Test to see if the Vent System should be included in the output
            ventIncluded = False
            for ventUnitName in _ventUnitsUsed:
                if ventSystem.Unit_Name == ventUnitName:
                    ventIncluded = True
                    break
                else:
                    ventIncluded = False
            
            # Basic Ventialtion
            if ventIncluded:
                # Create the Vent Unit in the Components Worksheet
                vent.append( PHPP_XL_Obj('Components', 'JH{}'.format(ventCompoRowStart + ventCount), ventSystem.Unit_Name if ventSystem else 'Default_Name' )) #  Create the Vent Unit
                vent.append( PHPP_XL_Obj('Components', 'JI{}'.format(ventCompoRowStart + ventCount), ventSystem.Unit_HR if ventSystem else 0.75 )) #  Vent Heat Recovery
                vent.append( PHPP_XL_Obj('Components', 'JJ{}'.format(ventCompoRowStart + ventCount), ventSystem.Unit_MR if ventSystem else 0 )) #  Vent Moisture Recovery
                vent.append( PHPP_XL_Obj('Components', 'JK{}'.format(ventCompoRowStart + ventCount), ventSystem.Unit_ElecEff if ventSystem else 0.45 )) #  Vent Elec Efficiency
                vent.append( PHPP_XL_Obj('Components', 'JL{}'.format(ventCompoRowStart + ventCount), 1)) #  DEFAULT MIN FLOW
                vent.append( PHPP_XL_Obj('Components', 'JM{}'.format(ventCompoRowStart + ventCount), 10000 )) #  DEFAULT MAX FLOW
                
                # Set the Vent Unit Type
                vent.append( PHPP_XL_Obj('Ventilation', 'L12', ventSystem.SystemType) ) 
                
                # Set the UD name for access in 'Addnl-Vent' dropdown list
                setattr(ventSystem, 'Unit_Name_UD', '{:02d}ud-{}'.format(ventCount+1, ventSystem.Unit_Name))
                
                # Build the Vent Unit
                vent.append(  PHPP_XL_Obj('Additional Vent',  'D{}'.format(ventUnitRowStart + ventCount) ,  1) ) # Quantity
                vent.append(  PHPP_XL_Obj('Additional Vent',  'E{}'.format(ventUnitRowStart + ventCount) ,  ventSystem.SystemName  if ventSystem.SystemName else '') ) # System Name
                #vent.append(  PHPP_XL_Obj('Additional Vent',  'E{}'.format(ventUnitRowStart + ventCount) ,  '=RIGHT(F{0},LEN(F{0})-5)'.format(ventUnitRowStart + ventCount))) # Description Formula to pull ERV Name
                vent.append(  PHPP_XL_Obj('Additional Vent',  'F{}'.format(ventUnitRowStart + ventCount) ,  ventSystem.Unit_Name_UD if ventSystem else '') ) # Vent Conpmonent UD Name
                vent.append(  PHPP_XL_Obj('Additional Vent',  'Q{}'.format(ventUnitRowStart + ventCount) ,  ventSystem.Exterior if ventSystem else '') ) # Exterior Installation?
                vent.append(  PHPP_XL_Obj('Additional Vent',  'X{}'.format(ventUnitRowStart + ventCount) ,  '2-Elec.') ) # Frost Protection Type
                vent.append(  PHPP_XL_Obj('Additional Vent',  'Y{}'.format(ventUnitRowStart + ventCount) ,  ventSystem.FrostTemp if ventSystem else '-5') ) # Frost Protection Temp
                
                # Build the Vent Unit Ducting
                vent.append( PHPP_XL_Obj('Additional Vent',  'D{}'.format(ventDuctsRowStart + ductsCount) , 1)) # Quantity
                vent.append( PHPP_XL_Obj('Additional Vent',  'E{}'.format(ventDuctsRowStart + ductsCount) , ventSystem.Duct01.DuctWidth if ventSystem else 104 ))
                vent.append( PHPP_XL_Obj('Additional Vent',  'H{}'.format(ventDuctsRowStart + ductsCount) , ventSystem.Duct01.InsulationThickness if ventSystem else 52 ))
                vent.append( PHPP_XL_Obj('Additional Vent',  'I{}'.format(ventDuctsRowStart + ductsCount) , ventSystem.Duct01.InsulationLambda if ventSystem else 0.04 ))
                vent.append( PHPP_XL_Obj('Additional Vent',  'J{}'.format(ventDuctsRowStart + ductsCount) , 'x' ))# Reflective
                vent.append( PHPP_XL_Obj('Additional Vent',  'L{}'.format(ventDuctsRowStart + ductsCount) , ventSystem.Duct01.DuctLength if ventSystem else 5 ))
                vent.append( PHPP_XL_Obj('Additional Vent',  'M{}'.format(ventDuctsRowStart + ductsCount) , '1'))
                
                vent.append( PHPP_XL_Obj('Additional Vent',  'D{}'.format(ventDuctsRowStart + ductsCount+1) , 1)) # Quantity
                vent.append( PHPP_XL_Obj('Additional Vent',  'E{}'.format(ventDuctsRowStart + ductsCount+1) , ventSystem.Duct02.DuctWidth if ventSystem else 104 ))
                vent.append( PHPP_XL_Obj('Additional Vent',  'H{}'.format(ventDuctsRowStart + ductsCount+1) , ventSystem.Duct02.InsulationThickness if ventSystem else 52 ))
                vent.append( PHPP_XL_Obj('Additional Vent',  'I{}'.format(ventDuctsRowStart + ductsCount+1) , ventSystem.Duct02.InsulationLambda if ventSystem else 0.04 ))
                vent.append( PHPP_XL_Obj('Additional Vent',  'J{}'.format(ventDuctsRowStart + ductsCount+1) , 'x' ))# Reflective
                vent.append( PHPP_XL_Obj('Additional Vent',  'L{}'.format(ventDuctsRowStart + ductsCount+1) , ventSystem.Duct02.DuctLength if ventSystem else 5 ))
                vent.append( PHPP_XL_Obj('Additional Vent',  'N{}'.format(ventDuctsRowStart + ductsCount+1) , '1'))
                
                vent.append( PHPP_XL_Obj('Additional Vent',  '{}{}'.format(chr(ductColCount), ventDuctsRowStart + ductsCount) , 1)) # Assign Duct to Vent
                vent.append( PHPP_XL_Obj('Additional Vent',  '{}{}'.format(chr(ductColCount), ventDuctsRowStart + ductsCount+1) , 1)) # Assign Duct to Vent
                
                ductColCount+=1
                ductsCount+=2
                ventCount+=1
                
            # Exhaust Ventialtion Objects
            if ventIncluded:
                # Add in any 'Exhaust Only' ventilation objects (kitchen hoods, etc...)
                for exhaustSystem in ventSystem.ExhaustObjs:
                    # Build the Vent in the Components Worksheet
                    vent.append( PHPP_XL_Obj('Components', 'JH{}'.format(ventCompoRowStart + ventCount), exhaustSystem.Name if exhaustSystem.Name else 'Exhaust' )) #  Create the Vent Unit
                    vent.append( PHPP_XL_Obj('Components', 'JI{}'.format(ventCompoRowStart + ventCount), 0 )) #  Vent Heat Recovery
                    vent.append( PHPP_XL_Obj('Components', 'JJ{}'.format(ventCompoRowStart + ventCount), 0 )) #  Vent Moisture Recovery
                    vent.append( PHPP_XL_Obj('Components', 'JK{}'.format(ventCompoRowStart + ventCount), 0.25 )) #  Vent Elec Efficiency
                    vent.append( PHPP_XL_Obj('Components', 'JL{}'.format(ventCompoRowStart + ventCount), 1)) #  DEFAULT MIN FLOW
                    vent.append( PHPP_XL_Obj('Components', 'JM{}'.format(ventCompoRowStart + ventCount), 10000 )) #  DEFAULT MAX FLOW
                    
                    # Set the UD name for access in 'Addnl-Vent' dropdown list
                    setattr(exhaustSystem, 'Unit_Name_UD', '{:02d}ud-{}'.format(ventCount+1, exhaustSystem.Name))
                    
                    # Build the Vent Unit
                    vent.append(  PHPP_XL_Obj('Additional Vent',  'D{}'.format(ventUnitRowStart + ventCount) ,  1) ) # Quantity
                    vent.append(  PHPP_XL_Obj('Additional Vent',  'E{}'.format(ventUnitRowStart + ventCount) ,  exhaustSystem.Name  if exhaustSystem.Name else 'Exhaust_Unit') )
                    vent.append(  PHPP_XL_Obj('Additional Vent',  'F{}'.format(ventUnitRowStart + ventCount) ,  exhaustSystem.Unit_Name_UD) ) # Vent Component UD Name
                    vent.append(  PHPP_XL_Obj('Additional Vent',  'Q{}'.format(ventUnitRowStart + ventCount) ,  '') ) # Exterior Installation?
                    vent.append(  PHPP_XL_Obj('Additional Vent',  'X{}'.format(ventUnitRowStart + ventCount) ,  '1-No') ) # Frost Protection Type
                    vent.append(  PHPP_XL_Obj('Additional Vent',  'Y{}'.format(ventUnitRowStart + ventCount) ,  '-5') ) # Frost Protection Temp
                    
                    # Build the Vent Unit Ducting
                    vent.append( PHPP_XL_Obj('Additional Vent',  'D{}'.format(ventDuctsRowStart + ductsCount) , 1)) # Quantity
                    vent.append( PHPP_XL_Obj('Additional Vent',  'E{}'.format(ventDuctsRowStart + ductsCount) , exhaustSystem.Duct01.DuctWidth if exhaustSystem else 104 ))
                    vent.append( PHPP_XL_Obj('Additional Vent',  'H{}'.format(ventDuctsRowStart + ductsCount) , exhaustSystem.Duct01.InsulationThickness if exhaustSystem else 52 ))
                    vent.append( PHPP_XL_Obj('Additional Vent',  'I{}'.format(ventDuctsRowStart + ductsCount) , exhaustSystem.Duct01.InsulationLambda if exhaustSystem else 0.04 ))
                    vent.append( PHPP_XL_Obj('Additional Vent',  'J{}'.format(ventDuctsRowStart + ductsCount) , 'x' ))# Reflective
                    vent.append( PHPP_XL_Obj('Additional Vent',  'L{}'.format(ventDuctsRowStart + ductsCount) , exhaustSystem.Duct01.DuctLength if exhaustSystem else 5 ))
                    vent.append( PHPP_XL_Obj('Additional Vent',  'M{}'.format(ventDuctsRowStart + ductsCount) , '1'))
                    
                    vent.append( PHPP_XL_Obj('Additional Vent',  'D{}'.format(ventDuctsRowStart + ductsCount+1) , 1)) # Quantity
                    vent.append( PHPP_XL_Obj('Additional Vent',  'E{}'.format(ventDuctsRowStart + ductsCount+1) , exhaustSystem.Duct02.DuctWidth if exhaustSystem else 104 ))
                    vent.append( PHPP_XL_Obj('Additional Vent',  'H{}'.format(ventDuctsRowStart + ductsCount+1) , exhaustSystem.Duct02.InsulationThickness if exhaustSystem else 52 ))
                    vent.append( PHPP_XL_Obj('Additional Vent',  'I{}'.format(ventDuctsRowStart + ductsCount+1) , exhaustSystem.Duct02.InsulationLambda if exhaustSystem else 0.04 ))
                    vent.append( PHPP_XL_Obj('Additional Vent',  'J{}'.format(ventDuctsRowStart + ductsCount+1) , 'x' ))# Reflective
                    vent.append( PHPP_XL_Obj('Additional Vent',  'L{}'.format(ventDuctsRowStart + ductsCount+1) , exhaustSystem.Duct02.DuctLength if exhaustSystem else 5 ))
                    vent.append( PHPP_XL_Obj('Additional Vent',  'N{}'.format(ventDuctsRowStart + ductsCount+1) , '1'))
                    
                    vent.append( PHPP_XL_Obj('Additional Vent',  '{}{}'.format(chr(ductColCount), ventDuctsRowStart + ductsCount) , 1)) # Assign Duct to Vent
                    vent.append( PHPP_XL_Obj('Additional Vent',  '{}{}'.format(chr(ductColCount), ventDuctsRowStart + ductsCount+1) , 1)) # Assign Duct to Vent
                    
                    ductColCount+=1
                    ductsCount+=2
                    ventCount+=1
    
    return vent

def getNonResRoomData(_inputBranch, _zones, _startRows):
    print "Creating 'Electricity non-res' Objects ... "
    elecNonRes = []
    rowStart_Lighting = _startRows.get('Electricity non-res').get('Lighting', 19)
    rowStart_OfficeEquip = _startRows.get('Electricity non-res').get('Office Equip', 62)
    rowStart_Kitchen = _startRows.get('Electricity non-res').get('Kitchen', 77)
    
    for i, roomObj in enumerate(_inputBranch):
        # First, see if the Room should be included in the output
        for zoneName in _zones:
            if roomObj.HostZoneName != zoneName:
                includeRoom = False
            elif getattr(roomObj, 'NonRes_RoomUse', '-') == '-':
                includeRoom = False
            else:
                includeRoom = True
                break
        
        # If the Room is to be included, write out the Excel objects
        if includeRoom:
            # Try and sort out the Room's Ventilation airflow and schedule if there is any
            if getattr(roomObj, 'NonRes_RoomLightingControl', None):
                roomID = '{}-{}'.format(getattr(roomObj, 'RoomNumber', None), getattr(roomObj, 'RoomName', None) )
                lightingControlNum = getattr(roomObj, 'NonRes_RoomLightingControl', '1-').split('-')[0]
                
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'C{}'.format(rowStart_Lighting+i), roomID))
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'D{}'.format(rowStart_Lighting+i), getattr(roomObj, 'FloorArea_Gross', None) ))
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'F{}'.format(rowStart_Lighting+i), getattr(roomObj, 'NonRes_RoomUse', None) ))
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'H{}'.format(rowStart_Lighting+i), 0)) # Deviation From North=0
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'J{}'.format(rowStart_Lighting+i), 0.69)) # Triple Glazing
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'M{}'.format(rowStart_Lighting+i), getattr(roomObj, 'RoomDepth', None)   ))
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'N{}'.format(rowStart_Lighting+i), '=D{}/M{}'.format(rowStart_Lighting+i, rowStart_Lighting+i)  ))
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'O{}'.format(rowStart_Lighting+i), getattr(roomObj, 'RoomClearHeight', None)   ))
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'P{}'.format(rowStart_Lighting+i), 1  )) # Lintel Height
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'Q{}'.format(rowStart_Lighting+i), 0  )) # Window Width                
                elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'W{}'.format(rowStart_Lighting+i), lightingControlNum ))
                
                if getattr(roomObj, 'NonRes_RoomMotionControl', 'No')=='Yes':
                    elecNonRes.append( PHPP_XL_Obj('Electricity non-res', 'X{}'.format(rowStart_Lighting+i), 'x' ))
    
    return elecNonRes

def getInfiltration(_inputBranch, _zonesToInclude):
    ##########################################
    ######   Envelope Airtightness    ########
    
    # Defaults
    Coef_E = 0.07
    Coef_F  = 15
    bldgWeightedACH = None
    bldgVn50 = None
    
    # Find the Floor-Area Weighted Average ACH of the Zones
    zonesVn50 = []
    zonesFloorArea = []
    zonesWeightedACH = []
    
    for zoneNametoInclude in _zonesToInclude:
        for zoneObj in _inputBranch:
            if zoneObj.ZoneName == zoneNametoInclude:
                try:
                    zonesFloorArea.append(zoneObj.FloorArea_Gross if zoneObj.FloorArea_Gross else False)
                    zonesWeightedACH.append(zoneObj.InfiltrationACH50 * zoneObj.FloorArea_Gross)
                    zonesVn50.append(zoneObj.Volume_Vn50)
                except:
                    pass
    
    if sum(zonesFloorArea)!= 0:
        bldgWeightedACH = sum(zonesWeightedACH) / sum(zonesFloorArea)
        bldgVn50 = sum(zonesVn50)
    
    airtightness = []
    print("Creating the Airtightness Objects...")
    airtightness.append(PHPP_XL_Obj('Ventilation', 'N25', Coef_E if Coef_E else float(0.07) ))# Wind protection E
    airtightness.append(PHPP_XL_Obj('Ventilation', 'N26', Coef_F if Coef_F else float(15) ))# Wind protection F
    airtightness.append(PHPP_XL_Obj('Ventilation', 'N27', bldgWeightedACH if bldgWeightedACH else float(0.6) ))# ACH50
    airtightness.append(PHPP_XL_Obj('Ventilation', 'P27', bldgVn50 if bldgVn50 else '=N9*1.2' ))#  Internal Reference Volume
    
    return airtightness

def updateStartRows(_startRowDict, _udIn):
    """Takes in the dictionary of start rows and any user-determined inputs
    modifies the dict values based on iputs. This is useful if the user has
    modified the PHPP for some reason and the start rows no longer align with 
    the normal ones. This happens esp. if the user adds more rows for an XXL
    size PHPP. (more rooms, more areas, etc...)"""
    try:
        for each in _udIn:
            parsed = each.split(':')
            newRowStart = int(parsed[1])
            worksheet, startItem = (parsed[0].split(','))
            _startRowDict[worksheet.lstrip().rstrip()][startItem.lstrip().rstrip()] = newRowStart
    except:
        udRowsMsg = "Couldn't read the udRowStarts_ input? Make sure it has dict keys separated by a comma and a semicolon before the value."
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, udRowsMsg)
    
    return _startRowDict

def filterName(_zoneName, _zonesNamesToFilter):
    flag = True
    for each in _zonesNamesToFilter:
        if each in _zoneName:
            flag = False
    
    return flag

def getGround(_floorElements, _zones):
    ground = []
    
    colLetter = {
        0: {'col0':'C', 'col1':'H', 'col2':'P'},
        1: {'col0':'W', 'col1':'AB', 'col2':'AJ'},
        2: {'col0':'AQ', 'col1':'AV', 'col2':'BD'}
        }
    
    if len(_floorElements) == 0:
        return ''
    
    if len(_floorElements) > 3:
        FloorElementsWarning= 'Warning: (grndFloorElements_) PHPP accepts only up to 3 unique \n'\
        'ground contact Floor Elements. Please simplify / consolidate your Floor Elements\n'\
        'before proceeding with export. For now only the first three Floor Elements\n'\
        'will be exported to PHPP.'
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, FloorElementsWarning)
        _floorElements = _floorElements[0:3]
    
    for i, floorElement in enumerate(_floorElements):
        if floorElement == None:
           continue
        
        # Filter for UD zone exclusions
        if floorElement.Zone not in _zones:
            continue
        
        col0 = colLetter[i]['col0']
        col1 = colLetter[i]['col1']
        col2 = colLetter[i]['col2']
        
        ground.append(PHPP_XL_Obj('Ground', col1+'9', floorElement.soilThermalConductivity ))
        ground.append(PHPP_XL_Obj('Ground', col1+'10', floorElement.soilHeatCapacity ))
        ground.append(PHPP_XL_Obj('Ground', col1+'18', floorElement.FloorArea ))
        ground.append(PHPP_XL_Obj('Ground', col1+'19', floorElement.PerimLen ))
        ground.append(PHPP_XL_Obj('Ground', col2+'17', floorElement.FloorUvalue ))
        ground.append(PHPP_XL_Obj('Ground', col2+'18', floorElement.PerimPsixLen ))
        ground.append(PHPP_XL_Obj('Ground', col1+'49', floorElement.groundWaterDepth ))
        ground.append(PHPP_XL_Obj('Ground', col1+'50', floorElement.groundWaterFlowrate ))
        
        if '1' in floorElement.Type or 'SLAB' in floorElement.Type.upper():
            # Slab on Grade Type
            ground.append(PHPP_XL_Obj('Ground', col0+'24', 'x' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'29', '' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'31', '' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'38', '' ))
            ground.append(PHPP_XL_Obj('Ground', col1+'25', floorElement.perimInsulDepth ))
            ground.append(PHPP_XL_Obj('Ground', col1+'26', floorElement.perimInsulThick ))
            ground.append(PHPP_XL_Obj('Ground', col1+'27', floorElement.perimInsulConductivity ))
            if 'V' in floorElement.perimInsulOrientation.upper():
                ground.append(PHPP_XL_Obj('Ground', col2+'25', '' ))
            else:
                ground.append(PHPP_XL_Obj('Ground', col2+'25', 'x' ))
        elif '2' in floorElement.Type or 'HEATED' in floorElement.Type.upper():
            # Heated Basement
            ground.append(PHPP_XL_Obj('Ground', col0+'24', '' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'29', 'x' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'31', '' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'38', '' ))
            ground.append(PHPP_XL_Obj('Ground', col1+'30', floorElement.WallHeight_BG ))
            ground.append(PHPP_XL_Obj('Ground', col2+'30', floorElement.WallU_BG ))
            
        elif '3' in floorElement.Type or 'UNHEATED' in floorElement.Type.upper():
            # Unheated Basement
            ground.append(PHPP_XL_Obj('Ground', col0+'24', '' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'29', '' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'31', 'x' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'38', '' ))
            ground.append(PHPP_XL_Obj('Ground', col1+'33', floorElement.WallHeight_AG ))
            ground.append(PHPP_XL_Obj('Ground', col2+'33', floorElement.WallU_AG ))
            ground.append(PHPP_XL_Obj('Ground', col1+'34', floorElement.WallHeight_BG ))
            ground.append(PHPP_XL_Obj('Ground', col2+'34', floorElement.WallU_BG ))
            ground.append(PHPP_XL_Obj('Ground', col2+'35', floorElement.FloorU ))
            ground.append(PHPP_XL_Obj('Ground', col1+'35', floorElement.ACH ))
            ground.append(PHPP_XL_Obj('Ground', col1+'36', floorElement.Volume ))
            
        elif '4' in floorElement.Type or 'CRAWL' in floorElement.Type.upper():
            # Suspended Floor overCrawlspace
            ground.append(PHPP_XL_Obj('Ground', col0+'24', '' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'29', '' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'31', '' ))
            ground.append(PHPP_XL_Obj('Ground', col0+'38', 'x' ))
            ground.append(PHPP_XL_Obj('Ground', col1+'39', floorElement.CrawlU ))
            ground.append(PHPP_XL_Obj('Ground', col1+'40', floorElement.WallHeight ))
            ground.append(PHPP_XL_Obj('Ground', col1+'41', floorElement.WallU ))
            ground.append(PHPP_XL_Obj('Ground', col2+'39', floorElement.VentOpeningArea ))
            ground.append(PHPP_XL_Obj('Ground', col2+'40', floorElement.windVelocity ))
            ground.append(PHPP_XL_Obj('Ground', col2+'41', floorElement.windFactor ))
            
    return ground

def getDHWSystem(_dhwSystems, _zones):
    ##########################################
    # If more that one system are to be used, combine them into a single system
    
    dhwSystems = defaultdict()
    
    for dhwSystem in _dhwSystems:
        dhwZonesServing = dhwSystem.ZonesAssigned
        for zoneName in dhwZonesServing:
            if zoneName in _zones:
                dhwSystems[dhwSystem.SystemName] = dhwSystem
   
    if len(dhwSystems.keys()) == 0:
        dhw_ = None
    elif len(list(set(dhwSystems.keys())))>1:
        dhw_ = combineDHWSystems(dhwSystems)
    else:
        dhw_ = _dhwSystems[0]
    
    ##########################################
    # DHW System Excel Objs
    dhwSystem = []
    if dhw_:
        print("Creating the 'DHW' Objects...")
        dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J146', dhw_.forwardTemp ) )
        dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'P145', 0 ) )
        dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'P29', 0 ) )
        
        # Usage Volume
        if dhw_.usage != None:
            if dhw_.usage.UsageType == 'Res':
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J47', dhw_.usage.demand_showers ) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J48', dhw_.usage.demand_others ) )
            elif dhw_.usage.UsageType == 'NonRes':
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J47', '=Q57'))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J48', '=Q58'))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J58', getattr(dhw_.usage, 'use_daysPerYear') ) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J62', 'x' if getattr(dhw_.usage, 'useShowers') != 'False' else '' ))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J63', 'x' if getattr(dhw_.usage, 'useHandWashing') != 'False' else '' ))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J64', 'x' if getattr(dhw_.usage, 'useWashStand') != 'False' else '' ))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J65', 'x' if getattr(dhw_.usage, 'useBidets') != 'False' else '' ))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J66', 'x' if getattr(dhw_.usage, 'useBathing') != 'False' else '' ))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J67', 'x' if getattr(dhw_.usage, 'useToothBrushing') != 'False' else '' ))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J68', 'x' if getattr(dhw_.usage, 'useCooking') != 'False' else '' ))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J74', 'x' if getattr(dhw_.usage, 'useDishwashing') != 'False' else '' ))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J75', 'x' if getattr(dhw_.usage, 'useCleanKitchen') != 'False' else '' ))
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J76', 'x' if getattr(dhw_.usage, 'useCleanRooms') != 'False' else '' ))
        
        # Recirc Piping
        if len(dhw_.circulation_piping)>0:
            dhwSystem.append( PHPP_XL_Obj('Aux Electricity', 'H29', 1 ) ) # Circulator Pump
            
        for colNum, recirc_line in enumerate(dhw_.circulation_piping):
            col = chr(ord('J') + colNum)
            
            if ord(col) <= ord('N'):
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 149), recirc_line.length ) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 150), recirc_line.diam ) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 151), recirc_line.insulThck ) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 152), 'x' if recirc_line.insulRefl=='Yes' else '' ) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 153), recirc_line.insulCond ) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 155), recirc_line.quality ) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 159), recirc_line.period ) )
            else:
                dhwRecircWarning = "Too many recirculation loops. PHPP only allows up to 5 loops to be entered.\nConsolidate the loops before moving forward"
                ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, dhwRecircWarning)
        
        # Branch Piping
        for colNum, branch_line in enumerate(dhw_.branch_piping):
            col = chr(ord('J') + colNum)
            
            if ord(col) <= ord('N'):
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 167), branch_line.diameter) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 168), branch_line.totalLength) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 169), branch_line.totalTapPoints) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 171), branch_line.tapOpenings) )
                dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', '{}{}'.format(col, 172), branch_line.utilisation) )
            else:
                dhwRecircWarning = "Too many branch piping sets. PHPP only allows up to 5 sets to be entered.\nConsolidate the piping sets before moving forward"
                ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, dhwRecircWarning)
        
        # Tanks
        if dhw_.tank1:
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J186', dhw_.tank1.type) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J189', 'x' if dhw_.tank1.solar==True else '') )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J191', dhw_.tank1.hl_rate) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J192', dhw_.tank1.vol) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J193', dhw_.tank1.stndbyFrac) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J195', dhw_.tank1.loction) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'J198', dhw_.tank1.locaton_t) )
        if dhw_.tank2:
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'M186', dhw_.tank2.type) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'M189', 'x' if dhw_.tank2.solar==True else '') )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'M191', dhw_.tank2.hl_rate) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'M192', dhw_.tank2.vol) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'M193', dhw_.tank2.stndbyFrac) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'M195', dhw_.tank2.loction) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'M198', dhw_.tank2.locaton_t) )
        if dhw_.tank_buffer:
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'P186', dhw_.tank_buffer.type) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'P191', dhw_.tank_buffer.hl_rate) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'P192', dhw_.tank_buffer.vol) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'P195', dhw_.tank_buffer.loction) )
            dhwSystem.append( PHPP_XL_Obj('DHW+Distribution', 'P198', dhw_.tank_buffer.locaton_t) )
        
    return dhwSystem

def combineDHWSystems(_dhwSystems):
    def getBranchPipeAttr(_dhwSystems, _attrName, _branchOrRecirc, _resultType):
        # Combine Elements accross the Systems
        results = DataTree[Object]()
        for sysNum, dhwSystem in enumerate(_dhwSystems.values()):
            pipingObj = getattr(dhwSystem, _branchOrRecirc)
            
            for i in range(0, 4):
                try:
                    temp = getattr(pipingObj[i], _attrName)
                    results.Add(temp, GH_Path(i) )
                except: 
                    pass
        
        output = []
        if _resultType == 'Sum':
            for i in range(results.BranchCount):
                output.append( sum( results.Branch(i) ) )
        elif _resultType == 'Average':
            for i in range(results.BranchCount):
                output.append( statistics.mean( results.Branch(i) ) )
        
        return output
    
    def combineTank(_dhwSystems, _tankName):
        # Combine Tank 1s
        hasTank = False
        tankObj = {'type':[], 'solar':[], 'hl_rate':[], 'vol':[],
        'stndbyFrac':[], 'loction':[], 'locaton_t':[]}
        for v in _dhwSystems.values():
            vTankObj = getattr(v, _tankName)
            
            if vTankObj != None:
                hasTank = True
                tankObj['type'].append( getattr(vTankObj, 'type') )
                tankObj['solar'].append( getattr(vTankObj, 'solar') )
                tankObj['hl_rate'].append( getattr(vTankObj, 'hl_rate') )
                tankObj['vol'].append( getattr(vTankObj, 'vol') )
                tankObj['stndbyFrac'].append( getattr(vTankObj, 'stndbyFrac') )
                tankObj['loction'].append( getattr(vTankObj, 'loction') )
                tankObj['locaton_t'].append( getattr(vTankObj, 'locaton_t') )
        
        if hasTank:
            return PHPP_DHW_tank(
                                _type = tankObj['type'][0] if len(tankObj['type']) != 0 else None,
                                _solar = tankObj['solar'][0] if len(tankObj['solar']) != 0 else None,
                                _hl_rate = statistics.mean(tankObj['hl_rate']) if len(tankObj['hl_rate']) != 0 else None,
                                _vol = statistics.mean(tankObj['vol']) if len(tankObj['vol']) != 0 else None,
                                _stndby_frac = statistics.mean(tankObj['stndbyFrac']) if len(tankObj['stndbyFrac']) != 0 else None,
                                _loc = tankObj['loction'][0] if len(tankObj['loction']) != 0 else None,
                                _loc_T = tankObj['locaton_t'][0] if len(tankObj['locaton_t']) != 0 else None,
                                )
        else:
            return None
    
    print 'Combining together DHW Systems...'
    #Combine Usages
    showers = []
    other = []
    for v in _dhwSystems.values():
        if isinstance(v.usage, PHPP_DHW_usage):
            showers.append( getattr(v.usage, 'demand_showers' ) )
            other.append( getattr(v.usage, 'demand_others' ) )
    
    dhwInputs = {'showers_demand_': sum(showers)/len(showers),
    'other_demand_': sum(other)/len(other)}
    print dhwInputs
    combinedUsage = PHPP_DHW_usage( dhwInputs )
    
    # Combine Branch Pipings
    combined_Diams = getBranchPipeAttr(_dhwSystems, 'diameter', 'branch_piping', 'Average')
    combined_Lens = getBranchPipeAttr(_dhwSystems, 'totalLength', 'branch_piping', 'Sum')
    combined_Taps =  getBranchPipeAttr(_dhwSystems, 'totalTapPoints', 'branch_piping',  'Sum')
    combined_Opens =  getBranchPipeAttr(_dhwSystems, 'tapOpenings', 'branch_piping',  'Average')
    combined_Utils =  getBranchPipeAttr(_dhwSystems, 'utilisation', 'branch_piping',  'Average')
    
    combined_BranchPipings = []
    for i in range(len(combined_Diams)):
        combinedBranchPiping = PHPP_DHW_branch_piping(
                                    combined_Diams[i],
                                    combined_Lens[i],
                                    combined_Taps[i],
                                    combined_Opens[i],
                                    combined_Utils[i]
                                   )
        combined_BranchPipings.append( combinedBranchPiping )
    
    # Combine Recirculation Pipings
    hasRecircPiping = False
    recircObj = {'length':[], 'diam':[], 'insulThck':[],
    'insulCond':[], 'insulRefl':[], 'quality':[], 'period':[],}
    for v in _dhwSystems.values():
        circObjs = v.circulation_piping
        if len(circObjs) != 0:
            hasRecircPiping = True
            try:
                recircObj['length'].append( getattr(circObjs[0], 'length'  ) )
                recircObj['diam'].append( getattr(circObjs[0], 'diam'  )) 
                recircObj['insulThck'].append( getattr(circObjs[0], 'insulThck'  ) )
                recircObj['insulCond'].append( getattr(circObjs[0], 'insulCond'  ) )
                recircObj['insulRefl'].append( getattr(circObjs[0], 'insulRefl'  ) )
                recircObj['quality'].append( getattr(circObjs[0], 'quality'  ) )
                recircObj['period'].append( getattr(circObjs[0], 'period'  ) )
            except:
                pass
    
    if hasRecircPiping:
        combined_RecircPipings = [PHPP_DHW_RecircPipe(
                        sum(recircObj['length']),
                        statistics.mean(recircObj['diam']) if len(recircObj['diam']) != 0 else None,
                        statistics.mean(recircObj['insulThck']) if len(recircObj['insulThck']) != 0 else None,
                        statistics.mean(recircObj['insulCond']) if len(recircObj['insulCond']) != 0 else None,
                        recircObj['insulRefl'][0] if len(recircObj['insulRefl']) != 0 else None,
                        recircObj['quality'][0] if len(recircObj['quality']) != 0 else None,
                        recircObj['period'][0] if len(recircObj['period']) != 0 else None
                        )]
    else:
        combined_RecircPipings = []
    
    # Combine the Tanks
    tank1 = combineTank(_dhwSystems, 'tank1')
    tank2 = combineTank(_dhwSystems, 'tank2')
    tank_buffer = combineTank(_dhwSystems, 'tank_buffer')
    
    # Build the combined / Averaged DHW System
    combinedDHWSys = PHPP_DHW_System(
                         _name='Combined', 
                         _usage=combinedUsage,
                         _fwdT=60, 
                         _pCirc=combined_RecircPipings, 
                         _pBran=combined_BranchPipings, 
                         _t1=tank1, 
                         _t2=tank2, 
                         _tBf=tank_buffer
                         )
    
    return combinedDHWSys

def getLocation(_locationObjs):
    climate = []
    
    if len(_locationObjs) == 0:
        return climate
    
    loc = _locationObjs[0]
    print("Creating the 'Climate' Objeects...")
    climate.append( PHPP_XL_Obj('Climate', 'D9', loc.Country if loc else 'US-United States of America' )) # Climate Data Set Name (Dropdown)
    climate.append( PHPP_XL_Obj('Climate', 'D10', loc.Region if loc else 'New York' )) # Climate Data Set Name (Dropdown)
    climate.append( PHPP_XL_Obj('Climate', 'D12', loc.DataSet if loc else 'US0055b-New York' )) # Climate Data Set Name (Dropdown)
    climate.append( PHPP_XL_Obj('Climate', 'D18', loc.Altitude if loc else '=D17' )) # Altitude
    
    return climate

##########################################
# Figure out the right Rows to start writing
# Modify values based on user input (if any)
startRows = {'Additional Ventilation': 
                {'Rooms':56,
                'Vent Unit Selection':97,
                'Vent Ducts':127 },
            'Components':
                {'Ventilator':15},
            'Areas':
                {'TB':145, 'Surfaces':41},
            'Electricity non-res':
                {'Lighting': 19,
                'Office Equip': 62,
                'Kitchen':77},
            }

if len(udRowStarts_)>0:
    startRows = updateStartRows(startRows, udRowStarts_)

##########################################
# Sort out which zones to include in the output
# Filter out any zones not to include
if _PHPPObjs.BranchCount>0:
    zones = [x.ZoneName for x in _PHPPObjs.Branch(8)]
    if len(zonesInclude_)>0: 
        zones = [x for x in zones if not filterName(x, zonesInclude_)]
    if len(zoneExclude_)>0:
        zones = [x for x in zones if filterName(x, zoneExclude_)]
    print 'Inlcuding Zones {} in the Export'.format(zones)

##########################################
# Construct the Excel-Ready Write Objects
toPHPP_Geom_ = DataTree[Object]() # Master tree to hold all the results
if _PHPPObjs.BranchCount != 0:
    uValuesList, uValueUID_Names    = getUvalues( _PHPPObjs.Branch(1) )
    winComponentsList               = getComponents( _PHPPObjs.Branch(5) )
    areasList, surfacesIncluded     = getAreas( _PHPPObjs.Branch(4), zones )
    tb_List                         = getThermalBridges( thermalBridges_, startRows)
    winSurfacesList                 = getWindows( _PHPPObjs.Branch(5), surfacesIncluded, _PHPPObjs.Branch(4) )   
    shadingList                     = getShading( _PHPPObjs.Branch(5), surfacesIncluded )
    tfa                             = getTFA(tfa_, _PHPPObjs.Branch(6), zones)
    addnlVentRooms, ventUnitsUsed   = getAddnlVentRooms( _PHPPObjs.Branch(6), _PHPPObjs.Branch(7), zones, startRows )
    vent                            = getAddnlVentSystems( _PHPPObjs.Branch(7), ventUnitsUsed, startRows )
    airtightness                    = getInfiltration( _PHPPObjs.Branch(8), zones)
    ground                          = getGround( grndFloorElements_ if len(grndFloorElements_)>0 else _PHPPObjs.Branch(11), zones )
    dhw                             = getDHWSystem( _PHPPObjs.Branch(10), zones )
    nonRes_Elec                     = getNonResRoomData( _PHPPObjs.Branch(6), zones, startRows )
    location                        = getLocation( _PHPPObjs.Branch(12) )
    
    ##########################################
    # Add all the Excel-Ready Objects to a master Tree for outputting / passing
    toPHPP_Geom_.AddRange(uValuesList, GH_Path(0))
    toPHPP_Geom_.AddRange(winComponentsList, GH_Path(1))
    toPHPP_Geom_.AddRange(areasList, GH_Path(2))
    toPHPP_Geom_.AddRange(winSurfacesList, GH_Path(3))
    toPHPP_Geom_.AddRange(shadingList, GH_Path(4))
    toPHPP_Geom_.AddRange(tfa, GH_Path(5))
    toPHPP_Geom_.AddRange(tb_List, GH_Path(6))
    toPHPP_Geom_.AddRange(addnlVentRooms, GH_Path(7))
    toPHPP_Geom_.AddRange(vent, GH_Path(8))  
    toPHPP_Geom_.AddRange(airtightness, GH_Path(9))  
    toPHPP_Geom_.AddRange(ground, GH_Path(10)) 
    toPHPP_Geom_.AddRange(dhw, GH_Path(11)) 
    toPHPP_Geom_.AddRange(nonRes_Elec, GH_Path(12))
    toPHPP_Geom_.AddRange(location, GH_Path(13))
    
    ##########################################
    # Give Warnings
    if len(toPHPP_Geom_.Branch(2))/10 > 100:
        AreasWarning = 'Warning: It looks like you have {:.0f} surfaces in the model. By Default\n'\
        'the PHPP can only hold 100 surfaces. Before writing out to the PHPP be sure to\n '\
        'add more lines to the "Areas" worksheet of your excel file.\n'\
        'After adding lines to the PHPP, be sure to input the correct Start Rows into\n'\
        'the "udRowStarts_" of this component.'.format(len(toPHPP_Geom_.Branch(2))/10)
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, AreasWarning)
    
    if len(toPHPP_Geom_.Branch(7))/17 > 30:
        VentWarning = 'Warning: It looks like you have {:.0f} rooms in the model. By Default\n'\
        'the PHPP can only hold 30 different rooms in the Additional Ventilation worksheet.\n'\
        'Before writing out to the PHPP be sure to add more lines to the\n'\
        '"Additional Ventilation" worksheet in the "Dimensionsing of Air Quantities" section.\n'\
        'After adding lines to the PHPP, be sure to input the correct Start Rows into\n'\
        'the "udRowStarts_" of this component.'.format(len(toPHPP_Geom_.Branch(7))/17)
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, VentWarning)
    
    if len(toPHPP_Geom_.Branch(12))/8 > 22:
        NonResWarning = 'Warning: It looks like you have {:.0f} Non-Residential Rooms in the model. By Default\n'\
        'the PHPP can only hold 22 different rooms in the "Electricity non-res" worksheet.\n'\
        'Before writing out to the PHPP be sure to add more lines to the \n '\
        '"Electricity non-res" worksheet in the "Lighting/non-residential" section.'.format(len(toPHPP_Geom_.Branch(12))/8)
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, NonResWarning)
