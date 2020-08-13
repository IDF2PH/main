#
# IDF2PHPP: A Plugin for exporting an EnergyPlus IDF file to the
# Passive House Planning Package (PHPP). Created by blgdtyp, llc
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
Used to set and import data from the Library files(s) that hold the component info.
Will open up a new dialogue window where the user can input the right file-paths for
the main 'Component Library' as well as a second 'Thermal Bridge' library. When the
User clicks 'Refresh' this will add all the components (assemblies, frames, glass, TBs)
to the Document UserText. To see this library - go to ...File/Properties..../Document User Text
Each component is stored as a unique dictionary with all its parameters. When the 'Refresh'
button is pressed, this will delete all existing Document User Text entries and
rewrite new ones based on the excel library files designated. 
-
EM June. 02, 2020
"""

import Eto
import rhinoscriptsyntax as rs
import clr
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel
import os
from System import Array
from System.IO import File
from shutil import copyfile
import random
import scriptcontext as sc
from datetime import datetime
from datetime import timedelta
import json
import System.Windows.Forms.DialogResult
import System.Drawing.Image
import Rhino

__commandname__ = "PHPP_Library"

class Dialog_Libaries(Eto.Forms.Dialog):
    
    def __init__(self, _libPath_compo, _libPath_tb, _libPath_psiInstalls):
        
        # Set variables based on any existing params in the file
        self.LibPath_compo = _libPath_compo
        self.LibPath_tb = _libPath_tb
        self.LibPath_psiInstalls = _libPath_psiInstalls
        
        # Set up the ETO Dialogue window
        self.Title = "Set PHPP Library Files:"
        self.Resizable = False
        
        # Create the Priamary Controls for the Dialog
        self.CompoLib_txtBox = Eto.Forms.TextBox( Text = str(self.LibPath_compo) )
        self.CompoLib_txtBox.Width=450
        self.CompoLib_loadButton = Eto.Forms.Button(Text = 'Load File...')
        self.CompoLib_loadButton.Click += self.OnLoadButton_compo
        self.CompoLib_refreshButton = Eto.Forms.Button(Text = 'Refresh Lib.')
        self.CompoLib_refreshButton.Click += self.OnRefreshButtonClick_compo
        
        self.TBLib_txtBox = Eto.Forms.TextBox( Text = str(self.LibPath_tb) )
        self.TBLib_txtBox.Width=450
        self.TBLib_loadButton = Eto.Forms.Button(Text = 'Load File...')
        self.TBLib_loadButton.Click += self.OnLoadButton_tb
        self.TBLib_refreshButton = Eto.Forms.Button(Text = 'Refresh Lib.')
        self.TBLib_refreshButton.Click += self.OnRefreshButtonClick_tb
        
        self.PsiLib_txtBox = Eto.Forms.TextBox( Text = str(self.LibPath_psiInstalls) )
        self.PsiLib_txtBox.Width=450
        self.PsiLib_loadButton = Eto.Forms.Button(Text = 'Load File...')
        self.PsiLib_loadButton.Click += self.OnLoadButton_psi
        self.PsiLib_refreshButton = Eto.Forms.Button(Text = 'Refresh Lib.')
        self.PsiLib_refreshButton.Click += self.OnRefreshButtonClick_psi
        
        # Create the OK / Cancel Buttons
        self.Button_OK = Eto.Forms.Button(Text = 'OK')
        self.Button_OK.Click += self.OnOKButtonClick
        self.Button_Cancel = Eto.Forms.Button(Text = 'Cancel')
        self.Button_Cancel.Click += self.OnCancelButtonClick
        
        ## Layout
        layout = Eto.Forms.DynamicLayout()
        layout.Spacing = Eto.Drawing.Size(5,10) # Space between (X, Y) all objs in the main window
        layout.Padding = Eto.Drawing.Padding(10) # Overall inset from main window
        
        # Group: Main Component Library
        self.groupbox_Main = Eto.Forms.GroupBox(Text = 'Main Component Library File:')
        layout_Group_Main = Eto.Forms.DynamicLayout()
        layout_Group_Main.Padding = Eto.Drawing.Padding(5) # Offfset from the outside of the Group Edge
        layout_Group_Main.Spacing = Eto.Drawing.Size(5,5) # Spacing between elements
        layout_Group_Main.AddRow(self.CompoLib_txtBox, self.CompoLib_loadButton, self.CompoLib_refreshButton)
        
        self.groupbox_Main.Content = layout_Group_Main
        layout.Add(self.groupbox_Main)
        
        # Group: Thermal Bridges
        self.groupbox_TB = Eto.Forms.GroupBox(Text = 'Thermal Bridges Library File:')
        layout_Group_TB = Eto.Forms.DynamicLayout()
        layout_Group_TB.Padding = Eto.Drawing.Padding(5) # Offfset from the outside of the Group Edge
        layout_Group_TB.Spacing = Eto.Drawing.Size(5,5) # Spacing between elements
        layout_Group_TB.AddRow(self.TBLib_txtBox, self.TBLib_loadButton, self.TBLib_refreshButton)
        
        self.groupbox_TB.Content = layout_Group_TB
        layout.Add(self.groupbox_TB)
        
        # Group: Psi-Installs
        self.groupbox_Psi = Eto.Forms.GroupBox(Text = 'Window Psi-Installs Library File:')
        layout_Group_Psi = Eto.Forms.DynamicLayout()
        layout_Group_Psi.Padding = Eto.Drawing.Padding(5) # Offfset from the outside of the Group Edge
        layout_Group_Psi.Spacing = Eto.Drawing.Size(5,5) # Spacing between elements
        layout_Group_Psi.AddRow(self.PsiLib_txtBox, self.PsiLib_loadButton, self.PsiLib_refreshButton)
        
        self.groupbox_Psi.Content = layout_Group_Psi
        layout.Add(self.groupbox_Psi)
        
        # Buttons
        vert = layout.BeginVertical()
        vert.Spacing = Eto.Drawing.Size(10,5)
        layout.AddRow(None, self.Button_Cancel, self.Button_OK, None)
        layout.EndVertical()
        
        # Set the dialog window Content
        self.Content = layout
    
    def LoadLibrary(self):
        """ Opens a dialogue window so the use can select a file
        """
        self.fd = Rhino.UI.OpenFileDialog()
        self.fd.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
        
        if self.fd.ShowDialog()!= System.Windows.Forms.DialogResult.OK:
            print 'Canceled...'
            return None
        else:
            return self.fd.FileName
    
    def OnLoadButton_compo(self, sender, e):
        print('Setting the Primary Component Library Path')
        self.LibPath_compo = self.LoadLibrary()
        self.CompoLib_txtBox.Text = self.LibPath_compo
    
    def OnLoadButton_tb(self, sender, e):
        print('Setting the Thermal Bridge Library Path')
        self.LibPath_tb = self.LoadLibrary()
        self.TBLib_txtBox.Text = self.LibPath_tb
    
    def OnLoadButton_psi(self, sender, e):
        print('Setting the Thermal Bridge Library Path')
        self.LibPath_psiInstalls = self.LoadLibrary()
        self.PsiLib_txtBox.Text = self.LibPath_psiInstalls
    
    def OnCancelButtonClick(self, sender, e):
        print('Canceled...')
        self.Update = False
        self.Close()
    
    def OnOKButtonClick(self, sender, e):
        print('Adding the Library Paths to the file Properties/DocumentUserText  Selected')
        self.Update = True
        self.Close()
    
    def OnRefreshButtonClick_compo(self, sender, e):
        print('Refreshing the Main Component Library...')
        self.refresh_DocumentUserText_compo()
        self.Update = True
    
    def OnRefreshButtonClick_tb(self, sender, e):
        print('Refreshing the Thermal Bridge Library...')
        self.refresh_DocumentUserText_TB()
        self.Update = True
        
    def OnRefreshButtonClick_psi(self, sender, e):
        print('Refreshing the Window Psi-Install Library...')
        self.refresh_DocumentUserText_psi()
        self.Update = True
    
    def refresh_DocumentUserText_compo(self):
        """ Removes all existing Frame, Glazing, and Assmbly values from the 
        DocumentUserText dictionary. Reads in all the values from the designated 
        XLSX Libary file and puts into the DocumentUserText dictionary
        """
        t1 = datetime.today()
        succesReading = False
        
        #####################################################################################
        #################              EXCEL-READ PART                      #################
        #####################################################################################
        try:
            if self.LibPath_compo != None:
                # If a Library File is set in the file...
                if os.path.exists(self.LibPath_compo):
                    print 'Reading the Main Component Library File....'
                    # Make a Temporary copy
                    saveDir = os.path.split(self.LibPath_compo)[0]
                    tempFile = '{}_temp.xlsx'.format(random.randint(0,1000))
                    tempFilePath = os.path.join(saveDir, tempFile)
                    copyfile(self.LibPath_compo, tempFilePath) # create a copy of the file to read from
                    
                    # Open the Excel Instance and File
                    ex = Excel.ApplicationClass()   
                    ex.Visible = False  # False means excel is hidden as it works
                    ex.DisplayAlerts = False
                    workbook = ex.Workbooks.Open(tempFilePath)
                    
                    # Find the Windows Worksheet
                    worksheets = workbook.Worksheets
                    try:
                        wsComponents = worksheets['Components']
                    except:
                        print "Could not find the 'Components' Worksheet in the taget file?"
                    
                    # Read in the Components from Excel Worksheet
                    # Come in as 2D Arrays..... grrr.....
                    xlArrayGlazing =  wsComponents.Range['IE15:IG113'].Value2
                    xlArrayFrames = wsComponents.Range['IL15:JC113'].Value2
                    xlArrayAssemblies = wsComponents.Range['E15:H113'].Value2
                    
                    workbook.Close()  # Close the worbook itself
                    ex.Quit()  # Close out the instance of Excel
                    os.remove(tempFilePath) # Remove the temporary read-file
                    
                    # Build the Glazing Library
                    lib_Glazing = []
                    xlListGlazing = list(xlArrayGlazing)
                    for i in range(0,  len(xlListGlazing), 3 ):
                        if xlListGlazing[i] != None:
                            newGlazing = [xlListGlazing[i],
                                        xlListGlazing[i+1],
                                        xlListGlazing[i+2]
                                        ]
                            lib_Glazing.append(newGlazing)
                    
                    # Build the Frame Library
                    lib_Frames = []
                    xlListFrames = list(xlArrayFrames)
                    for i in range(0,  len(xlListFrames), 18 ):
                        newFrame = []
                        if xlListFrames[i] != None:
                            for k in range(i, i+18):
                                newFrame.append(xlListFrames[k])
                            lib_Frames.append(newFrame)
                    
                    # Build the Assembly Library
                    lib_Assemblies = []
                    xlListAssemblies = list(xlArrayAssemblies)
                    for i in range(0, len(xlListAssemblies), 4):
                        newAssembly = [xlListAssemblies[i],
                                xlListAssemblies[i+1],
                                xlListAssemblies[i+2],
                                xlListAssemblies[i+3]
                        ]
                        lib_Assemblies.append(newAssembly)
                        
                succesReading = True
            else:
                print('No Library filepath set.')
                succesReading = False
        except:
            print('Woops... something went wrong reading from the Excel file?')
            succesReading = False
        
        #####################################################################################
        #################          Write the Rhino Doc Library              #################
        #####################################################################################
        if succesReading==True:
            # If there is any existing 'Document User-Text' Assemblies, Frames or Glazing Types, remove them
            print('Clearing out the old values from the Document UserText....')
            if rs.IsDocumentUserText():
                try:
                    for eachKey in rs.GetDocumentUserText():
                        if 'PHPP_lib_Glazing_' in eachKey or 'PHPP_lib_Frame_' in eachKey or 'PHPP_lib_Assmbly_' in eachKey:
                            print 'Deleting: Key:', eachKey
                            rs.SetDocumentUserText(eachKey) # If no values are passed, deletes the Key/Value
                except:
                    print 'Error reading the Library File for some reason.'
            
            print('Writing New Library elements to the Document UserText....')
            # Write the new Assemblies to the Document's User-Text
            if lib_Assemblies:
                for eachAssembly in lib_Assemblies:
                    if eachAssembly[0] != None and len(eachAssembly[0])>0: # Filter our Null values
                        newAssembly = {"Name":eachAssembly[0],
                                        "Thickness":eachAssembly[1],
                                        "uValue":eachAssembly[2],
                                        "intInsulation":eachAssembly[3]
                                        }
                        rs.SetDocumentUserText("PHPP_lib_Assmbly_{}".format(newAssembly["Name"]), json.dumps(newAssembly) )
            
            # Write the new Glazings to the Document's User-Text
            if lib_Glazing:
                for eachGlazing in lib_Glazing:
                    newGlazingType = {"Name" : eachGlazing[0],
                                        "gValue" : eachGlazing[1],
                                        "uValue" : eachGlazing[2]
                                        }
                    rs.SetDocumentUserText("PHPP_lib_Glazing_{}".format(newGlazingType["Name"]), json.dumps(newGlazingType) )
            
            # Write the new Frames to the Document's User-Text
            if lib_Frames:
                for eachFrame in lib_Frames:
                    newFrameType = {"Name" : eachFrame[0],
                                        "uFrame_L" : eachFrame[1],
                                        "uFrame_R" : eachFrame[2],
                                        "uFrame_B" : eachFrame[3],
                                        "uFrame_T" : eachFrame[4],
                                        
                                        "wFrame_L" : eachFrame[5],
                                        "wFrame_R" : eachFrame[6],
                                        "wFrame_B" : eachFrame[7],
                                        "wFrame_T" : eachFrame[8],
                                        
                                        "psiG_L" : eachFrame[9],
                                        "psiG_R" : eachFrame[10],
                                        "psiG_B" : eachFrame[11],
                                        "psiG_T" : eachFrame[12],
                                        
                                        "psiInst_L" : eachFrame[13],
                                        "psiInst_R" : eachFrame[14],
                                        "psiInst_B" : eachFrame[15],
                                        "psiInst_T" : eachFrame[16],
                                        
                                        "chi" : eachFrame[17]
                                        }
                    rs.SetDocumentUserText("PHPP_lib_Frame_{}".format(newFrameType["Name"]), json.dumps(newFrameType) )
        
        else:
            print 'Did not modify the Rhino User-Text Attributes.'
        
        print 'This ran in:', datetime.today() - t1, 'seconds'
        
        return 1
    
    def refresh_DocumentUserText_TB(self):
        """ Removes all existing Thermal Bridge values from the 
        DocumentUserText dictionary. Reads in all the values from the designated 
        XLSX Libary file and puts into the DocumentUserText dictionary
        """
        t1 = datetime.today()
        succesReading = False
        
        #####################################################################################
        #################              EXCEL-READ PART                      #################
        #####################################################################################
        try:
            if self.LibPath_tb != None:
                # If a Library File is set in the file...
                if os.path.exists(self.LibPath_tb):
                    print 'Reading the Thermal Bridge Library File....'
                    # Make a Temporary copy
                    saveDir = os.path.split(self.LibPath_tb)[0]
                    tempFile = '{}_temp.xlsx'.format(random.randint(0,1000))
                    tempFilePath = os.path.join(saveDir, tempFile)
                    copyfile(self.LibPath_tb, tempFilePath) # create a copy of the file to read from
                    
                    # Open the Excel Instance and File
                    ex = Excel.ApplicationClass()   
                    ex.Visible = False  # False means excel is hidden as it works
                    ex.DisplayAlerts = False
                    workbook = ex.Workbooks.Open(tempFilePath)
                    
                    
                    
                    
                    
                    # Find the Windows Worksheet
                    worksheets = workbook.Worksheets
                    try:
                        ws_TBs = worksheets['Thermal Bridges']
                    except:
                        print "Could not find any worksheet names 'Thermal Bridges' in the taget file?"
                    
                    # Read in the Thermal Bridges from an Excel Worksheet
                    # Come in as 2D Arrays..... grrr.....
                    xlArray_TBs = ws_TBs.Range['A2:C100'].Value2
                    
                    workbook.Close()  # Close the worbook itself
                    ex.Quit()  # Close out the instance of Excel
                    os.remove(tempFilePath) # Remove the temporary read-file
                    
                    # Build the Thermal Bridge Library
                    lib_TBs = []
                    xlList_TBs = list(xlArray_TBs)
                    for i in range(0, len(xlList_TBs), 3):
                        newAssembly = [xlList_TBs[i],
                                xlList_TBs[i+1],
                                xlList_TBs[i+2],
                        ]
                        lib_TBs.append(newAssembly)
                        
                succesReading = True
            else:
                print('No Library filepath set.')
                succesReading = False
        except:
            print('Woops... something went wrong reading from the Excel file?')
            succesReading = False
        
        #####################################################################################
        #################          Write the Rhino Doc Library              #################
        #####################################################################################
        if succesReading==True:
            # If there are any existing 'Document User-Text' Thermal Bridge objects, remove them
            print('Clearing out the old values from the Document UserText....')
            if rs.IsDocumentUserText():
                try:
                    for eachKey in rs.GetDocumentUserText():
                        if 'PHPP_lib_TB_' in eachKey:
                            print 'Deleting: Key:', eachKey
                            rs.SetDocumentUserText(eachKey) # If no values are passed, deletes the Key/Value
                except:
                    print 'Error reading the Library File for some reason.'
            
            print('Writing New Library elements to the Document UserText....')
            # Write the new Assemblies to the Document's User-Text
            if lib_TBs:
                for eachAssembly in lib_TBs:
                    if eachAssembly[0] != None and len(eachAssembly[0])>0: # Filter our Null values
                        newAssembly = {"Name":eachAssembly[0],
                                        "Psi-Value":eachAssembly[1],
                                        "fRsi":eachAssembly[2],
                                        }
                        rs.SetDocumentUserText("PHPP_lib_TB_{}".format(newAssembly["Name"]), json.dumps(newAssembly) )
        else:
            print 'Did not modify the Rhino User-Text Attributes.'
        
        print 'This ran in:', datetime.today() - t1, 'seconds'
        
        return 1
    
    def refresh_DocumentUserText_psi(self):
        """Reads Psi-Install Excel and adds entries to the file's DocumentUserText
        
        Removes all existing Window Pis-Install values from the
        DocumentUserText dictionary. Reads in all the values from the designated 
        XLSX Libary file and puts into the DocumentUserText dictionary
        
        Will look in the self.LibPath_psiInstalls (path) Excel file for:
            Worksheet named 'Psi-Installs'
            Reads from Cells B3:F103 (5 columns):
                Col 1: Type Name (str)
                Col 2: Left Psi-Install (float)
                Col 3: Right Psi-Install (float)
                Col 4: Bottom Psi-Install (float)
                Col 5: Top Psi-Install (float)
        """
        
        t1 = datetime.today()
        succesReading = False
        
        #####################################################################################
        #################              EXCEL-READ PART                      #################
        #####################################################################################
        try:
            if self.LibPath_psiInstalls != None:
                # If a Library File is set in the file...
                if os.path.exists(self.LibPath_psiInstalls):
                    print 'Reading the Window Psi-Installs Library File....'
                    # Make a Temporary copy
                    saveDir = os.path.split(self.LibPath_psiInstalls)[0]
                    tempFile = '{}_tempPsiInst.xlsx'.format(random.randint(0,1000))
                    tempFilePath = os.path.join(saveDir, tempFile)
                    copyfile(self.LibPath_psiInstalls, tempFilePath) # create a copy of the file to read from
                    
                    # Open the Excel Instance and File
                    ex = Excel.ApplicationClass()   
                    ex.Visible = False  # False means excel is hidden as it works
                    ex.DisplayAlerts = False
                    workbook = ex.Workbooks.Open(tempFilePath)
                    
                    # Find the Windows Worksheet
                    worksheets = workbook.Worksheets
                    try:
                        ws_PsiInst = worksheets['Psi-Installs']
                    except:
                        try:
                            ws_PsiInst = worksheets['Psi-Install']
                        except:
                            print "Could not find any worksheet names 'Psi-Installs' in the taget file?"
                    
                    # Read in the Thermal Bridges from an Excel Worksheet
                    # Come in as 2D Arrays..... grrr.....
                    xlArray_Psi = ws_PsiInst.Range['B3:F103'].Value2
                    
                    workbook.Close()  # Close the worbook itself
                    ex.Quit()  # Close out the instance of Excel
                    os.remove(tempFilePath) # Remove the temporary read-file
                    
                    # Build the Thermal Bridge Library
                    lib_PsiInst = []
                    xlList_Psi = list(xlArray_Psi)
                    for i in range(0, len(xlList_Psi), 5):
                        if xlList_Psi[i] != None:
                            newPsiInstall = [
                                    xlList_Psi[i],
                                    xlList_Psi[i+1],
                                    xlList_Psi[i+2],
                                    xlList_Psi[i+3],
                                    xlList_Psi[i+4]
                            ]
                            lib_PsiInst.append(newPsiInstall)
                        
                succesReading = True
            else:
                print('No Library filepath set.')
                succesReading = False
        except:
            print('Woops... something went wrong reading from the Excel file?')
            succesReading = False
        print lib_PsiInst
        #####################################################################################
        #################          Write the Rhino Doc Library              #################
        #####################################################################################
        if succesReading==True:
            # If there are any existing 'Document User-Text' Psi-Install objects, remove them
            print('Clearing out the old values from the Document UserText....')
            if rs.IsDocumentUserText():
                try:
                    for eachKey in rs.GetDocumentUserText():
                        if 'PHPP_lib_PsiInstall_' in eachKey:
                            print 'Deleting: Key:', eachKey
                            rs.SetDocumentUserText(eachKey) # If no values are passed, deletes the Key/Value
                except:
                    print 'Error reading the Library File for some reason.'
            
            print('Writing New Library elements to the Document UserText....')
            # Write the new Psi-Installs to the Document's User-Text
            if lib_PsiInst:
                for eachPsiInstall in lib_PsiInst:
                    if eachPsiInstall[0] != None and len(eachPsiInstall[0])>0: # Filter our Null values
                        newPsiInstall = {"Typename":eachPsiInstall[0],
                                        "Left":eachPsiInstall[1],
                                        "Right":eachPsiInstall[2],
                                        "Bottom":eachPsiInstall[3],
                                        "Top":eachPsiInstall[4],
                                        }
                        rs.SetDocumentUserText("PHPP_lib_PsiInstall_{}".format(newPsiInstall["Typename"]), json.dumps(newPsiInstall) )
        else:
            print 'Did not modify the Rhino User-Text Attributes.'
        
        print 'This ran in:', datetime.today() - t1, 'seconds'
        
        return 1

def RunCommand( is_interactive ):
    # First, get any existing library paths in the Document User Text
    libPath_compo = rs.GetDocumentUserText('PHPP_Component_Lib')
    libPath_tb = rs.GetDocumentUserText('PHPP_TB_Lib')
    libPath_psiInstalls = rs.GetDocumentUserText('PHPP_PsiInstalls_Lib')
    
    # Apply defaults if nothing already loaded into the file
    libPath_compo = libPath_compo if libPath_compo else 'Load Main Library File...'
    libPath_tb = libPath_tb if libPath_tb else 'Load TB Library File ...'
    libPath_psiInstalls = libPath_psiInstalls if libPath_psiInstalls else 'Load Window Psi-Installs Library File...'
    
    # Call the Dialog Window
    dialog = Dialog_Libaries(libPath_compo, libPath_tb, libPath_psiInstalls)
    rc = dialog.ShowModal(Rhino.UI.RhinoEtoApp.MainWindow)
    
    # If 'OK' button, then update the DocumentUserText
    if dialog.Update:
        if dialog.LibPath_compo:
            rs.SetDocumentUserText('PHPP_Component_Lib', dialog.LibPath_compo)
            print 'Setting PHPP Component Library for the Rhino Document to: {}'.format(dialog.LibPath_compo)
        
        if dialog.LibPath_tb:
            rs.SetDocumentUserText('PHPP_TB_Lib', dialog.LibPath_tb)
            print 'Setting PHPP Thermal Bridge Library for the Rhino Document to: {}'.format(dialog.LibPath_tb)
    
        if dialog.LibPath_psiInstalls:
            rs.SetDocumentUserText('PHPP_PsiInstalls_Lib', dialog.LibPath_psiInstalls)
            print 'Setting Window Psi-Installs Library for the Rhino Document to: {}'.format(dialog.LibPath_psiInstalls)
    
    return 1

# temp for debuggin in editor
# RunCommand(True) 