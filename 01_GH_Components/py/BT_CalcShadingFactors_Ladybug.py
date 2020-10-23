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
This component is used when calculating 'Shading Factors' using Ladbug 'radiationAnalysis' components. Use this after the radiationAnalysis in order to determine the shading factors for each window. Pass the results of this component on to an IDF2PH 'Apply Win. Shading Factors' component.
Note that when using the Ladybug 'radiationAnalysis' components, all surfaces are treated as opaque and there are no bounces taken into account (I think). 
-
EM October 23, 2020
    Args:
        _windows_mesh: (Tree) A Tree where each branch is one of the window objects. Connect this to "analysisMesh" output on a Ladybug "radiationAnalysis" component.
        _windows_radiation: (Tree) A Tree where each branch is one of the window objects. Connect this to "radiationResult" output on a Ladybug "radiationAnalysis" component.
        _sphere_testVec: (Tree) A Tree where each branch is one of the window objects. Connect this to "testVec" output on a Ladybug "radiationAnalysis" component.
        _sphere_radiation: (Tree) A Tree where each branch is one of the window objects. Connect this to "radiationResult" output on a Ladybug "radiationAnalysis" component.
    Returns:
        win_radiation_: (kWh/m2) The average solar radiation for each window over the analysis period. 
        sphere_segment_radiation_: (kWh/m2) The solar radiation for the sphere segment which is closest to the each window. The order of the elements in the list matches the windows input. 
        win_shading_factor_: (%) The calculated 'Shading Factor' for each window over the analysis period. The order of this list should match the input window order. Pass these results on to an IDF2PH 'Apply Win. Shading Factors' component.
        max_error_: (%) The largest 'error' between a window surface normal and the sphere segments. Try and keep below 3-5% I think.
"""

ghenv.Component.Name = "BT_CalcShadingFactors_Ladybug"
ghenv.Component.NickName = "Calc Ladybug Shading Factors"
ghenv.Component.Message = 'OCT_23_2020'
ghenv.Component.IconDisplayMode = ghenv.Component.IconDisplayMode.application
ghenv.Component.Category = "BT"
ghenv.Component.SubCategory = "01 | Model"

import ghpythonlib.components as ghc
import Rhino
from collections import namedtuple
import scriptcontext as sc
import math
import rhinoscriptsyntax as rs

def calc_mesh_face_rad(_face, _verts, _rad):
    face_area = calc_mesh_face_area(_face, _verts)
    return face_area * rad

def calc_mesh_face_area(_face, _verts):
    pt_a = _verts[_face.A]
    pt_b = _verts[_face.B]
    pt_c = _verts[_face.C]
    
    a = pt_a.DistanceTo(pt_b)
    b = pt_b.DistanceTo(pt_c)
    c = pt_c.DistanceTo(pt_a)
    p = 0.5 * (a + b + c)
    area1 = math.sqrt(p * (p-a) * (p-b) * (p-c))
    
    if _face.IsQuad:
        pt_d = _verts[_face.D]
        a = pt_a.DistanceTo(pt_c)
        b = pt_c.DistanceTo(pt_d)
        c = pt_d.DistanceTo(pt_a)
        p = 0.5 * (a + b + c)
        area2 = math.sqrt(p * (p-a) * (p-b) * (p-c))
    else:
        area2 = 0
    
    return area1 + area2

# Calc the window Radiation / m2
# ------------------------------------------------------------------------------
Window = namedtuple('Window', ['radiation', 'area', 'srfc_normal'])
windows = []

for branch_msh, branch_rad in zip(_windows_mesh.Branches, _windows_radiation.Branches):
    for msh in branch_msh:
        win_total_rad = sum([calc_mesh_face_rad(face, msh.Vertices, rad) for face, rad in zip(msh.Faces, branch_rad)])
        
        win_area = Rhino.Geometry.AreaMassProperties.Compute(msh).Area
        win_rad_per_m2 = win_total_rad / win_area
        win_srfc_normal = ghc.Average( msh.FaceNormals)
        
        # Window Objects
        windows.append(Window(win_rad_per_m2, win_area, win_srfc_normal))


# Build the sphere radiation objects
# ------------------------------------------------------------------------------
OrientationRad = namedtuple('OrientationRad', ['radiation', 'srfc_normal'])
northVec = Rhino.Geometry.Vector3d(0,1,0)
objs = []

for v_list, r_list in zip(_sphere_testVec.Branches, _sphere_radiation.Branches):
    for v, r in zip( v_list, r_list):
        objs.append(OrientationRad(r, v))

# Find the closest Sphere segment to the window (based on srfc normal) 
# Calc shading factor
# ------------------------------------------------------------------------------
win_radiation_ = []
sphere_segment_radiation_ = []
win_shading_factor_ = []
max_error = 0

for window in windows:
    
    smallest_angle = 360
    for obj in objs:
        angle_diff = rs.VectorAngle(obj.srfc_normal, window.srfc_normal)
        
        if angle_diff < smallest_angle:
            closestRadVal = obj.radiation
            smallest_angle = angle_diff
    
    win_radiation_.append(window.radiation)
    sphere_segment_radiation_.append(closestRadVal)   
    win_shading_factor_.append(window.radiation / closestRadVal )
    
    if smallest_angle > max_error:
        max_error = smallest_angle
