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
Enables the COM interface connecting Rhino to Excel on the user's computer.
> This should identify a running excel instance or start excel if it is not already running.
> Copies a file if the target file name doesn't already exist. 
> Passes the full file name out, along with whether or not a copy was made.
-
Component by Jack Hymowitz, July 31, 2020

    Args:
        _run: Set to true to enable the excel application, false saves the open sheet and stops the application
        visibl_e: Set to true to show the Excel application on the screen. Default true.
        useUserWorkbook_: Set to true to look for and use an open excel interface instead of starting a new one. For now, should not be used (set to false or disconnected)
        oldFilename: The full file path to the source file
        newDirectory: The location to put the new copied file, or a blank string ("") if newFilename should be a relative location or if newFilename is the entire filepath.
        newFilename: The destination name to copy the file to (can be the full path if newDirectory is a blank string)
    Returns:
        excel: The Excel COM interface created, or None if not running
"""

ghenv.Component.Name = "BT_OpenXLWorkbook"
ghenv.Component.NickName = "Open XL Workbook"
ghenv.Component.Message = 'JUL_31_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "02 | IDF2PHPP"

from ghpythonlib.componentbase import executingcomponent as component
import Grasshopper, GhPython
import System
import Rhino
import rhinoscriptsyntax as rs

import scriptcontext as sc
import clr
import os
clr.AddReferenceByName('Microsoft.Office.Interop.Excel')#, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from System.Runtime.InteropServices import Marshal
from Microsoft.Office.Interop import Excel

from shutil import copyfile
import Grasshopper.Kernel as ghK
import inspect


class ExcelInstance:
    """A holder for the methods we use to interact with the Excel COM interface"""
    
    #Run once on startup, defines the variables that we will use
    def __init__(self):
        self.ex=None
        self.activeWorkbook=None
        self.activeWorkbookName=""
        self.sheetsDict={}
        self.userOpened=False
    
    #Starts a brand new excel instance
    def startNewInstance(self):
        """Start a new Excel Application instance"""
        self.ex = Excel.ApplicationClass()
        self.ex.DisplayAlerts = False
        self.ex.EnableEvents = False
        self.userOpened=False
    
    #Not working right now 100%, but should find an already running excel window
    def findExistingInstance(self):
        """Find an Excel instance that is already running
        Returns True if an already running instance was found, False otherwise"""
        #try:
        self.userOpened=True
        self.activeWorkbookName=="NEW_USER_FILE"
        self.ex=Marshal.GetActiveObject("Excel.Application")
        print(self.ex)
        if self.ex==None:
            return None
        print(self.ex.Workbooks)
        print(dir(self.ex.Workbooks))
        print(self.ex.Workbooks.Count)
        if(self.ex.Workbooks.Count>0):
            for workbook in self.ex.Workbooks:
                self.activeWorkbook=workbook
                break
                #there is definitely a better way of doing this
        else:
            return None
        print(self.activeWorkbook==None)
        if self.activeWorkbook==None:
            return None
        self.userOpened=True
        self.activeWorkbookName="NEW_USER_FILE"
        self.loadSheets()
        return True
        #except:
        #    return False
        #Define later...
    
    def openWorkbook(self,filename):
        """Open a workbook in the running excel interface, closing the open one
        Args:
            filename: The workbook to open
        """
        if self.activeWorkbookName==filename:
            return False
        self.close(True)
        self.activeWorkbookName=filename
        self.activeWorkbook=self.ex.Workbooks.Open(filename)
        self.loadSheets()
        return True
    
    #Finds the sheets in the open workbook
    def loadSheets(self):
        self.sheetsDict={}
        for sheet in self.activeWorkbook.Worksheets:
            self.sheetsDict[sheet.Name]=sheet
            sheet.Unprotect()
        
    def saveAndQuit(self,closeIfUser):
        """Close the running excel instance, and save first
        Args:
            closeIfUser (boolean): Set to true if this method is to run even if this was a user-opened file, 
                false to run only if was an instance created here."""

        self.save()
        self.close(closeIfUser)
        self.quit(closeIfUser)
        
    def save(self):
        try:
            self.ex.activeWorkbook.Save()
            return True
        except:
            return False
            
    def close(self, closeIfUser):
        """Close the open excel workbook
        Args:
            closeIfUser (boolean): Set to true if this method is to run even if this was a user-opened file, 
                false to run only if was an instance created here."""
        if self.activeWorkbook!=None or (not closeIfUser and self.userOpened):
            return
        try:
            self.ex.activeWorkbook.Close()
        except:
            pass
    def quit(self, closeIfUser):
        """Stop the running excel instance
        Args:
            closeIfUser (boolean): Set to true if this method is to run even if this was a user-opened file, 
                false to run only if was an instance created here."""
        self.activeWorkbook=None
        self.activeWorkbookName=""
        if self.ex!=None:
            self.ex.Quit()
            ex=None
        
    def __unicode__(self):
        return u"Excel Instance | Active Worksheet: {self.activeWorkbookName}".format(self=self)
    def __str__(self):
        return unicode(self).encode("utf-8")
    def __repr__(self):
        return "{}( _nm={!r}, _ex={!r}, _activeWorkbook={!r}, _name={!r}".format(
               self.__class__.__name__,
               self.ex,
               self.activeWorkbook,
               self.activeWorkbookName )

class MyComponent(component):
    
    def StartExcel(self,useUserWorkbook):
        if not "excel" in sc.sticky: #Excel interface isn't already running, so start it
            excel = ExcelInstance()
            if useUserWorkbook:
                excel.findExistingInstance()    #Try looking for a running instance
            else:
                excel.startNewInstance()            #If didn't find one, make one
            sc.sticky["excel"]=excel
        else:
            excel = sc.sticky["excel"]
        return excel
        
    def StopExcel(self):
        if "excel" in sc.sticky:
            print("Quitting")
            excel=sc.sticky["excel"]
            try:
                excel.saveAndQuit(False)
            except:
                pass
            del sc.sticky["excel"]
    
    def doCopy(self, oldFilename,newDirectory,newFilename):
        
        if not os.path.exists(newDirectory):
            os.mkdir(newDirectory)
        dest=newDirectory+os.path.sep+newFilename+".xlsx"
        if os.path.exists(dest):
            msg1 = "File already copied, not overwritten"
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Remark, msg1)
            return dest
        try:
            src=os.path.abspath(oldFilename)
        except:
            msg1 = "Source file not found"
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg1)
            return None
        try:
            copyfile(src,dest)
        except:
            msg1 = "Unable to copy file"
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, msg1)
            return dest
        return dest
    
    def OpenWorkbook(self,excel,useUserWorkbookDefault,oldFilename,newDirectory,newFilename):
        if excel:
            """if useUserWorkbookDefault and excel.userOpened and excel.activeWorkbookName=="NEW_USER_FILE": #If a sheet is already open, set it up
                if "XLSdata" in sc.sticky: 
                    del sc.sticky["XLSdata"]
                excel.loadSheets()
                excel.activeWorkbookName="USER_FILE"
            elif useUserWorkbookDefault and excel.userOpened:   #If we already set it up, we are OK
                pass"""
            if newFilename:
                filename=self.doCopy(oldFilename,newDirectory,newFilename)
                if excel.openWorkbook(filename): #If we need to open a new sheet, set it up
                    if "XLSdata" in sc.sticky: 
                        del sc.sticky["XLSdata"]
                    excel.loadSheets()
            else:
                return False
        else:
            return False
        return True
    def RunScript(self,run,visible,useUserWorkbook,oldFilename,newDirectory,newFilename):
        if not run:
            self.StopExcel()
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, "Excel Not Started")
            return None
        excel=self.StartExcel(useUserWorkbook)
        if excel==None:
            return None
        if not useUserWorkbook:
            if visible==None:
                visible=True
            excel.ex.Visible = visible
            excel.ex.ScreenUpdating = visible
        if not excel:
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Error, "Excel Unable to Start")
            return None
        if not (useUserWorkbook or oldFilename and newFilename):
            ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Warning, "Not running, either run is false or filename not set.")
            return None
        #if useUserWorkbookDefault:
            
        #outputName=self.doCopy(oldFilename,newDirectory,newFilename)
        #if not outputName and not useUserWorkbookDefault:
        #    return (None,"Unable to copy file.")
        if self.OpenWorkbook(excel,useUserWorkbook,oldFilename,newDirectory,newFilename):
            return excel
        ghenv.Component.AddRuntimeMessage(ghK.GH_RuntimeMessageLevel.Error, "Unable to open workbook")
        return None
