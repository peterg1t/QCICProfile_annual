#############################START LICENSE##########################################
# Copyright (C) 2019 Pedro Martinez
#
# # This program is free software: you can redistribute it and/or modify
# # it under the terms of the GNU Affero General Public License as published
# # by the Free Software Foundation, either version 3 of the License, or
# # (at your option) any later version (the "AGPL-3.0+").
#
# # This program is distributed in the hope that it will be useful,
# # but WITHOUT ANY WARRANTY; without even the implied warranty of
# # MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# # GNU Affero General Public License and the additional terms for more
# # details.
#
# # You should have received a copy of the GNU Affero General Public License
# # along with this program. If not, see <http://www.gnu.org/licenses/>.
#
# # ADDITIONAL TERMS are also included as allowed by Section 7 of the GNU
# # Affero General Public License. These additional terms are Sections 1, 5,
# # 6, 7, 8, and 9 from the Apache License, Version 2.0 (the "Apache-2.0")
# # where all references to the definition "License" are instead defined to
# # mean the AGPL-3.0+.
#
# # You should have received a copy of the Apache-2.0 along with this
# # program. If not, see <http://www.apache.org/licenses/LICENSE-2.0>.
#############################END LICENSE##########################################


###########################################################################################
#
#   Script name: qc-ICProfile
#
#   Description: Tool for batch processing and report generation of ICProfile files
#
#   Example usage: python qc-ICProfile "/folder/"
#
#   Author: Pedro Martinez
#   pedro.enrique.83@gmail.com
#   5877000722
#   Date:2019-04-09
#
###########################################################################################



import os
import sys
import pydicom
import re
import argparse
import linecache
import tokenize
from PIL import *
import subprocess
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.backends.backend_pdf import PdfPages
from tqdm import tqdm
import numpy as np
import pandas as pd
from openpyxl import Workbook, cell, load_workbook
from math import *
import scipy.integrate as integrate

reffolder = '/home/peter/Dropbox/MScMedPhysi/scripts-medphys/QCICProfile_annual/2015_Method/'


def area_calc(profile,coord):
    # print(profile,coord)
    area = np.trapz(profile,coord)
    return area


def xls_cell_spec(energy, gantry, coll, wedge,slope,XdiffAve,XdiffStd,XCI,XdiffMax,XdiffMin,YdiffAve,YdiffStd,YCI,YdiffMax,YdiffMin,central_value):
    print(energy)
    print(gantry)
    print(coll, wedge,slope)
    print(type(energy), type(gantry), type(coll), type(wedge),type(slope))
    CellChange = {}
    Attributes=[]

    # cell specification for 6MV
    #Open field
    if (energy == '6 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '0' and slope == 0):
        energy = 6
        CellChange['B21'] = XdiffAve
        CellChange['C21'] = XdiffStd
        CellChange['D21'] = XCI
        CellChange['E21'] = XdiffMax
        CellChange['F21'] = XdiffMin
        CellChange['B29'] = YdiffAve
        CellChange['C29'] = YdiffStd
        CellChange['D29'] = YCI
        CellChange['E29'] = YdiffMax
        CellChange['F29'] = YdiffMin
        CellChange['B65'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '6 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '15' and slope > 0):
        energy = 6
        CellChange['B22'] = XdiffAve
        CellChange['C22'] = XdiffStd
        CellChange['D22'] = XCI
        CellChange['E22'] = XdiffMax
        CellChange['F22'] = XdiffMin
        CellChange['B30'] = YdiffAve
        CellChange['C30'] = YdiffStd
        CellChange['D30'] = YCI
        CellChange['E30'] = YdiffMax
        CellChange['F30'] = YdiffMin
        CellChange['C65'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '6 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '60' and slope > 0):
        energy = 6
        CellChange['B23'] = XdiffAve
        CellChange['C23'] = XdiffStd
        CellChange['D23'] = XCI
        CellChange['E23'] = XdiffMax
        CellChange['F23'] = XdiffMin
        CellChange['B31'] = YdiffAve
        CellChange['C31'] = YdiffStd
        CellChange['D31'] = YCI
        CellChange['E31'] = YdiffMax
        CellChange['F31'] = YdiffMin
        CellChange['E65'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '6 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '15' and slope < 0):
        energy = 6
        CellChange['B24'] = XdiffAve
        CellChange['C24'] = XdiffStd
        CellChange['D24'] = XCI
        CellChange['E24'] = XdiffMax
        CellChange['F24'] = XdiffMin
        CellChange['B32'] = YdiffAve
        CellChange['C32'] = YdiffStd
        CellChange['D32'] = YCI
        CellChange['E32'] = YdiffMax
        CellChange['F32'] = YdiffMin
        CellChange['D65'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '6 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '60' and slope < 0):
        energy = 6
        CellChange['B25'] = XdiffAve
        CellChange['C25'] = XdiffStd
        CellChange['D25'] = XCI
        CellChange['E25'] = XdiffMax
        CellChange['F25'] = XdiffMin
        CellChange['B33'] = YdiffAve
        CellChange['C33'] = YdiffStd
        CellChange['D33'] = YCI
        CellChange['E33'] = YdiffMax
        CellChange['F33'] = YdiffMin
        CellChange['F65'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    
    
    
    
    
    
    
    
    
    
    
    
    # cell specification for 10MV
    # Open cell
    if (energy == '10 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '0' and slope == 0):
        energy = 10
        CellChange['I21'] = XdiffAve
        CellChange['J21'] = XdiffStd
        CellChange['K21'] = XCI
        CellChange['L21'] = XdiffMax
        CellChange['M21'] = XdiffMin
        CellChange['I29'] = YdiffAve
        CellChange['J29'] = YdiffStd
        CellChange['K29'] = YCI
        CellChange['L29'] = YdiffMax
        CellChange['M29'] = YdiffMin
        CellChange['B74'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '10 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '15' and slope > 0):
        energy = 10
        CellChange['I22'] = XdiffAve
        CellChange['J22'] = XdiffStd
        CellChange['K22'] = XCI
        CellChange['L22'] = XdiffMax
        CellChange['M22'] = XdiffMin
        CellChange['I30'] = YdiffAve
        CellChange['J30'] = YdiffStd
        CellChange['K30'] = YCI
        CellChange['L30'] = YdiffMax
        CellChange['M30'] = YdiffMin
        CellChange['C74'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '10 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '60' and slope > 0):
        energy = 10
        CellChange['I23'] = XdiffAve
        CellChange['J23'] = XdiffStd
        CellChange['K23'] = XCI
        CellChange['L23'] = XdiffMax
        CellChange['M23'] = XdiffMin
        CellChange['I31'] = YdiffAve
        CellChange['J31'] = YdiffStd
        CellChange['K31'] = YCI
        CellChange['L31'] = YdiffMax
        CellChange['M31'] = YdiffMin
        CellChange['E74'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '10 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '15' and slope < 0):
        energy = 10
        CellChange['I24'] = XdiffAve
        CellChange['J24'] = XdiffStd
        CellChange['K24'] = XCI
        CellChange['L24'] = XdiffMax
        CellChange['M24'] = XdiffMin
        CellChange['I32'] = YdiffAve
        CellChange['J32'] = YdiffStd
        CellChange['K32'] = YCI
        CellChange['L32'] = YdiffMax
        CellChange['M32'] = YdiffMin
        CellChange['D74'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '10 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '60' and slope < 0):
        energy = 10
        CellChange['I25'] = XdiffAve
        CellChange['J25'] = XdiffStd
        CellChange['K25'] = XCI
        CellChange['L25'] = XdiffMax
        CellChange['M25'] = XdiffMin
        CellChange['I33'] = YdiffAve
        CellChange['J33'] = YdiffStd
        CellChange['K33'] = YCI
        CellChange['L33'] = YdiffMax
        CellChange['M33'] = YdiffMin
        CellChange['F74'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']
















        # cell specification for 15MV
        #Open cell
    if (energy == '15 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '0' and slope ==0):
        energy = 15
        CellChange['P21'] = XdiffAve
        CellChange['Q21'] = XdiffStd
        CellChange['R21'] = XCI
        CellChange['S21'] = XdiffMax
        CellChange['T21'] = XdiffMin
        CellChange['P29'] = YdiffAve
        CellChange['Q29'] = YdiffStd
        CellChange['R29'] = YCI
        CellChange['S29'] = YdiffMax
        CellChange['T29'] = YdiffMin
        CellChange['B83'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']


    if (energy == '15 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '15' and slope > 0):
        energy = 15
        CellChange['P22'] = XdiffAve
        CellChange['Q22'] = XdiffStd
        CellChange['R22'] = XCI
        CellChange['S22'] = XdiffMax
        CellChange['T22'] = XdiffMin
        CellChange['P30'] = YdiffAve
        CellChange['Q30'] = YdiffStd
        CellChange['R30'] = YCI
        CellChange['S30'] = YdiffMax
        CellChange['T30'] = YdiffMin
        CellChange['C83'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '15 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '60' and slope > 0):
        energy = 15
        CellChange['P23'] = XdiffAve
        CellChange['Q23'] = XdiffStd
        CellChange['R23'] = XCI
        CellChange['S23'] = XdiffMax
        CellChange['T23'] = XdiffMin
        CellChange['P31'] = YdiffAve
        CellChange['Q31'] = YdiffStd
        CellChange['R31'] = YCI
        CellChange['S31'] = YdiffMax
        CellChange['T31'] = YdiffMin
        CellChange['E83'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '15 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '15' and slope < 0):
        energy = 15
        CellChange['P24'] = XdiffAve
        CellChange['Q24'] = XdiffStd
        CellChange['R24'] = XCI
        CellChange['S24'] = XdiffMax
        CellChange['T24'] = XdiffMin
        CellChange['P32'] = YdiffAve
        CellChange['Q32'] = YdiffStd
        CellChange['R32'] = YCI
        CellChange['S32'] = YdiffMax
        CellChange['T32'] = YdiffMin
        CellChange['D83'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '15 MV' and gantry == '0 deg' and coll == '0 deg' and wedge == '60' and slope < 0):
        energy = 15
        CellChange['P25'] = XdiffAve
        CellChange['Q25'] = XdiffStd
        CellChange['R25'] = XCI
        CellChange['S25'] = XdiffMax
        CellChange['T25'] = XdiffMin
        CellChange['P33'] = YdiffAve
        CellChange['Q33'] = YdiffStd
        CellChange['R33'] = YCI
        CellChange['S33'] = YdiffMax
        CellChange['T33'] = YdiffMin
        CellChange['F83'] = central_value
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']














    if (energy == '6 MV' and gantry == '90 deg' and coll == '90 deg' and wedge == '60' and slope > 0):
        energy = 6
        CellChange['B40'] = XdiffAve
        CellChange['C40'] = XdiffStd
        CellChange['D40'] = XCI
        CellChange['E40'] = XdiffMax
        CellChange['F40'] = XdiffMin
        CellChange['B45'] = YdiffAve
        CellChange['C45'] = YdiffStd
        CellChange['D45'] = YCI
        CellChange['E45'] = YdiffMax
        CellChange['F45'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '6 MV' and gantry == '90 deg' and coll == '90 deg' and wedge == '60' and slope < 0):
        energy = 6
        CellChange['B41'] = XdiffAve
        CellChange['C41'] = XdiffStd
        CellChange['D41'] = XCI
        CellChange['E41'] = XdiffMax
        CellChange['F41'] = XdiffMin
        CellChange['B46'] = YdiffAve
        CellChange['C46'] = YdiffStd
        CellChange['D46'] = YCI
        CellChange['E46'] = YdiffMax
        CellChange['F46'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '6 MV' and gantry == '90 deg' and coll == '270 deg' and wedge == '60' and slope > 0):
        energy = 6
        CellChange['B53'] = XdiffAve
        CellChange['C53'] = XdiffStd
        CellChange['D53'] = XCI
        CellChange['E53'] = XdiffMax
        CellChange['F53'] = XdiffMin
        CellChange['B58'] = YdiffAve
        CellChange['C58'] = YdiffStd
        CellChange['D58'] = YCI
        CellChange['E58'] = YdiffMax
        CellChange['F58'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '6 MV' and gantry == '90 deg' and coll == '270 deg' and wedge == '60' and slope < 0):
        energy = 6
        CellChange['B54'] = XdiffAve
        CellChange['C54'] = XdiffStd
        CellChange['D54'] = XCI
        CellChange['E54'] = XdiffMax
        CellChange['F54'] = XdiffMin
        CellChange['B59'] = YdiffAve
        CellChange['C59'] = YdiffStd
        CellChange['D59'] = YCI
        CellChange['E59'] = YdiffMax
        CellChange['F59'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']














    if (energy == '10 MV' and gantry == '90 deg' and coll == '90 deg' and wedge == '60' and slope > 0):
        energy = 10
        CellChange['I40'] = XdiffAve
        CellChange['J40'] = XdiffStd
        CellChange['K40'] = XCI
        CellChange['L40'] = XdiffMax
        CellChange['M40'] = XdiffMin
        CellChange['I45'] = YdiffAve
        CellChange['J45'] = YdiffStd
        CellChange['K45'] = YCI
        CellChange['L45'] = YdiffMax
        CellChange['M45'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '10 MV' and gantry == '90 deg' and coll == '90 deg' and wedge == '60' and slope < 0):
        energy = 10
        CellChange['I41'] = XdiffAve
        CellChange['J41'] = XdiffStd
        CellChange['K41'] = XCI
        CellChange['L41'] = XdiffMax
        CellChange['M41'] = XdiffMin
        CellChange['I46'] = YdiffAve
        CellChange['J46'] = YdiffStd
        CellChange['K46'] = YCI
        CellChange['L46'] = YdiffMax
        CellChange['M46'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '10 MV' and gantry == '90 deg' and coll == '270 deg' and wedge == '60' and slope > 0):
        energy = 10
        CellChange['I53'] = XdiffAve
        CellChange['J53'] = XdiffStd
        CellChange['K53'] = XCI
        CellChange['L53'] = XdiffMax
        CellChange['M53'] = XdiffMin
        CellChange['I58'] = YdiffAve
        CellChange['J58'] = YdiffStd
        CellChange['K58'] = YCI
        CellChange['L58'] = YdiffMax
        CellChange['M58'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '10 MV' and gantry == '90 deg' and coll == '270 deg' and wedge == '60' and slope < 0):
        energy = 10
        CellChange['I54'] = XdiffAve
        CellChange['J54'] = XdiffStd
        CellChange['K54'] = XCI
        CellChange['L54'] = XdiffMax
        CellChange['M54'] = XdiffMin
        CellChange['I59'] = YdiffAve
        CellChange['J59'] = YdiffStd
        CellChange['K59'] = YCI
        CellChange['L59'] = YdiffMax
        CellChange['M59'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']




















    if (energy == '15 MV' and gantry == '90 deg' and coll == '90 deg' and wedge == '60' and slope > 0):
        energy = 15
        CellChange['P40'] = XdiffAve
        CellChange['Q40'] = XdiffStd
        CellChange['R40'] = XCI
        CellChange['S40'] = XdiffMax
        CellChange['T40'] = XdiffMin
        CellChange['P45'] = YdiffAve
        CellChange['Q45'] = YdiffStd
        CellChange['R45'] = YCI
        CellChange['S45'] = YdiffMax
        CellChange['T45'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '15 MV' and gantry == '90 deg' and coll == '90 deg' and wedge == '60' and slope < 0):
        energy = 15
        CellChange['P41'] = XdiffAve
        CellChange['Q41'] = XdiffStd
        CellChange['R41'] = XCI
        CellChange['S41'] = XdiffMax
        CellChange['T41'] = XdiffMin
        CellChange['P46'] = YdiffAve
        CellChange['Q46'] = YdiffStd
        CellChange['R46'] = YCI
        CellChange['S46'] = YdiffMax
        CellChange['T46'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '15 MV' and gantry == '90 deg' and coll == '270 deg' and wedge == '60' and slope > 0):
        energy = 15
        CellChange['P53'] = XdiffAve
        CellChange['Q53'] = XdiffStd
        CellChange['R53'] = XCI
        CellChange['S53'] = XdiffMax
        CellChange['T53'] = XdiffMin
        CellChange['P58'] = YdiffAve
        CellChange['Q58'] = YdiffStd
        CellChange['R58'] = YCI
        CellChange['S58'] = YdiffMax
        CellChange['T58'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    if (energy == '15 MV' and gantry == '90 deg' and coll == '270 deg' and wedge == '60' and slope < 0):
        energy = 15
        CellChange['P54'] = XdiffAve
        CellChange['Q54'] = XdiffStd
        CellChange['R54'] = XCI
        CellChange['S54'] = XdiffMax
        CellChange['T54'] = XdiffMin
        CellChange['P59'] = YdiffAve
        CellChange['Q59'] = YdiffStd
        CellChange['R59'] = YCI
        CellChange['S59'] = YdiffMax
        CellChange['T59'] = YdiffMin
        Attributes = ['XdiffAve', 'XdiffStd', 'XCI', 'XdiffMax', 'XdiffMin','XdiffAve', 'YdiffStd', 'YCI', 'YdiffMax', 'YdiffMin','central_value']

    return CellChange, Attributes





def reffileselect(energy,wedge,slopeY):
    # print(energy, wedge, slopeY, type(energy), type(wedge), type(slopeY))
    if energy == '6 MV' and int(wedge) == 15 and slopeY > 0:
        print('Ref File','6X15WY1IN')
        refname = reffolder+'6X/15W Y1IN.txt'
    elif energy == '6 MV' and int(wedge) == 15 and slopeY < 0:
        print('Ref File','6X15WY2OUT')
        refname = reffolder + '6X/15W Y2OUT.txt'
    elif energy == '6 MV' and int(wedge) == 60 and slopeY > 0:
        print('Ref File','6X60WY1IN')
        refname = reffolder + '6X/60W Y1IN.txt'
    elif energy == '6 MV' and int(wedge) == 60 and slopeY < 0:
        print('Ref File','6X60WY2OUT')
        refname = reffolder + '6X/60W Y2OUT.txt'
    elif energy == '6 MV' and int(wedge) == 0 and slopeY == 0:
        print('Ref File','20X20 Open')
        refname = reffolder + '6X/20x20 Open.txt'
    elif energy == '10 MV' and int(wedge) == 15 and slopeY > 0:
        print('Ref File','10X15WY1IN')
        refname = reffolder + '10X/15W Y1IN.txt'
    elif energy == '10 MV' and int(wedge) == 15 and slopeY < 0:
        print('Ref File','10X15WY2OUT')
        refname = reffolder + '10X/15W Y2OUT.txt'
    elif energy == '10 MV' and int(wedge) == 60 and slopeY > 0:
        print('Ref File','10X60WY1IN')
        refname = reffolder + '10X/60W Y1IN.txt'
    elif energy == '10 MV' and int(wedge) == 60 and slopeY < 0:
        print('Ref File','10X60WY2OUT')
        refname = reffolder + '10X/60W Y2OUT.txt'
    elif energy == '10 MV' and int(wedge) == 0 and slopeY == 0:
        print('Ref File','20X20 Open')
        refname = reffolder + '10X/20x20 Open.txt'
    elif energy == '15 MV' and int(wedge) == 15 and slopeY > 0:
        print('Ref File','15X15WY1IN')
        refname = reffolder + '15X/15W Y1IN.txt'
    elif energy == '15 MV' and int(wedge) == 15 and slopeY < 0:
        print('Ref File','15X15WY2OUT')
        refname = reffolder + '15X/15W Y2OUT.txt'
    elif energy == '15 MV' and int(wedge) == 60 and slopeY > 0:
        print('Ref File','15X60WY1IN')
        refname = reffolder + '15X/60W Y1IN.txt'
    elif energy == '15 MV' and int(wedge) == 60 and slopeY < 0:
        print('Ref File','15X60WY2OUT')
        refname = reffolder + '15X/60W Y2OUT.txt'
    elif energy == '15 MV' and int(wedge) == 0 and slopeY == 0:
        print('Ref File','20X20 Open')
        refname = reffolder + '15X/20x20 Open.txt'
    return refname




def CIcalc(data1,data2,li,ri):
    diff = 100 * (data2[li:ri]-data1[li:ri])/data2[li:ri]
    diffAve = np.mean(np.abs(diff))
    diffStd = np.std(np.abs(diff),ddof=1)
    CI = diffAve/diffStd
    diffMax = np.max(diff)
    diffMin = np.min(diff)
    # print('data1',data1,len(data1))
    # print('data1',data1[li:ri],len(data1))
    # print('data2',data2,len(data1))
    # print('data2',data2[li:ri],len(data1))

    # print('diff',diff)

    return diffAve, diffStd, CI, diffMax, diffMin





def read_icp(dirname,excel):
# this section reads the header and detects to what cells to write
    gainlist=[]
    modelist=[]
    energylist=[]
    wedgelist=[]
    gantrylist=[]
    colllist=[]
    central_cham_list=[]


    if excel != None:
        print(excel)
        wb2 = load_workbook(excel)
        ws1 = wb2['EDW']

    for subdir, dirs, files in os.walk(dirname):
        for file in files:
            print('File',file)
            print('Start header processing')
            f = open(dirname+file, mode='r', encoding='ISO-8859-1') #,encoding='utf-8-sig')
            lines = f.readlines()
            f.close()
            calname = 'N:'+lines[5].rstrip().split('N:')[1]
            # gain = int(re.findall(r'\d+',lines[20])[0])  # using regex
            gain = int(lines[20].rstrip().split('\t')[1])
            mode = lines[29].rstrip().split('\t')[1]
            energy = lines[29].rstrip().split('\t')[3]
            wedge = lines[31].rstrip().split('\t')[3]
            gantry = lines[33].rstrip().split('\t')[1]
            coll = lines[33].rstrip().split('\t')[3]





            print('Calibration file name = ',calname)
            print('Gain = ',gain)
            print('Mode = ',mode)
            print('Energy = ',energy)
            print('Wedge = ',wedge)
            print('Gantry = ',gantry)
            print('Collimator = ',coll)


            my_dict = {}
            my_list = []


            if mode=='X-Ray FFF' and gain!=2:
                print('Error, gain was set incorrectly')
                exit(0)







            print('Start data processing')
            # reading measurement file
            df = pd.read_csv(dirname+file,skiprows=106,delimiter='\t')
            tblen = df.shape[0]  #length of the table

            #These vectors will hold the inline and crossline data
            RawCountXvect = []
            CorrCountXvect=[] # correcting for leakage using the expression # for Detector(n) = {RawCount(n) - TimeTic * LeakRate(n)} * cf(n)
            CorrCountXleak = []
            RawCountYvect = []
            CorrCountYvect=[]
            CorrCountYleak = []
            RawCountPDvect = [] # positive diagonal
            CorrCountPDvect=[]
            CorrCountPDleak = []
            RawCountNDvect = [] # negative diagonal
            CorrCountNDvect=[]
            CorrCountNDleak = []
            BiasX=[]
            CalibX=[]
            BiasY=[]
            CalibY=[]
            BiasPD=[]
            CalibPD=[]
            BiasND=[]
            CalibND=[]
            Timetic=df['TIMETIC'][3]
            # These vectors will have the location of the sensors in the x, y and diagonal directions
            Y = (np.linspace(1,65,65)-33)/2
            X= np.delete(np.delete(Y,31),32)
            PD = np.delete(np.delete((np.linspace(1,65,65)-33)/2,31),32)
            ND=PD

            PDX = PD/ np.cos(pi / 4)
            PDY = PD/ np.sin(pi / 4)
            NDX = ND/ np.cos(pi / 4 - pi / 2)
            NDY = ND/ np.sin(pi / 4 - pi / 2)



            QuadWedgeCal=[0.5096,0,0,0,0,0,0,0,0] # 6xqw,15xqw,6fffqw,10fffqw,6eqw,9eqw,12eqw,16eqw,20eqw




            # figs = [] #in this list we will hold all the figures
            # print('Timetic=',Timetic*1e-6,df['TIMETIC'][0],df['TIMETIC'][1],df['TIMETIC'][2],df['TIMETIC'][3]) # duration of the measurement



            for column in df.columns[5:68]: #this section records the X axis (-)
                CorrCountXvect.append((df[column][3] - Timetic*df[column][0])*df[column][1]/gain)#the corrected data for leakage = Timetic*Bias*Calibration/Gain
                # CorrCountXvect.append((df[column][3])*df[column][1])#the corrected data for leakage = Timetic*Bias*Calibration/Gain
                central_value_Xvect = CorrCountXvect[len(CorrCountXvect) // 2]
                BiasX.append(df[column][0]) # already used in the formula above but saving them just in case
                CalibX.append(df[column][1])
                RawCountXvect.append(df[column][3])

            for column in df.columns[68:133]: #this section records the Y axis (|)
                CorrCountYvect.append((df[column][3] - Timetic*df[column][0])*df[column][1]/gain)
                # CorrCountYvect.append((df[column][3])*df[column][1])
                BiasY.append(df[column][0]) # already used in the formula above but saving them just in case
                CalibY.append(df[column][1])
                RawCountYvect.append(df[column][3])

            for column in df.columns[133:196]: #this section records the D1 axis  (/)
                CorrCountPDvect.append((df[column][3] - Timetic*df[column][0])*df[column][1]/gain)
                # CorrCountPDvect.append((df[column][3])*df[column][1])
                BiasPD.append(df[column][0]) # already used in the formula above but saving them just in case
                CalibPD.append(df[column][1])
                RawCountPDvect.append(df[column][3])

            for column in df.columns[196:259]: #this section records the D2 axis  (\)
                CorrCountNDvect.append((df[column][3] - Timetic*df[column][0])*df[column][1]/gain)
                # CorrCountNDvect.append((df[column][3])*df[column][1])
                BiasND.append(df[column][0]) # already used in the formula above but saving them just in case
                CalibND.append(df[column][1])
                RawCountNDvect.append(df[column][3])



            #Let's calculate the slope around the center of the Y axis to determine if it was Y1IN or Y2OUT
            slopeY = (CorrCountYvect[32+1] - CorrCountYvect[32-1])/(Y[32+1]-Y[32-1])
            if wedge == '0':
                slopeY=0
                central_value_Y_openfield=CorrCountYvect[32]
            else:
                central_cham_list.append(CorrCountYvect[32])
            #if slope > 0 Y1IN if slope < 0 Y2OUT


            refname = reffileselect(energy,wedge,slopeY)

            # BEGIN FIGURE
            # #Temporary figure placement
            # fig=plt.figure()
            # ax=Axes3D(fig)
            # ax.scatter(X,np.zeros(len(X)),CorrCountXvect/np.max(CorrCountXvect),label='X profile')
            # ax.scatter(np.zeros(len(Y)),Y,CorrCountYvect/np.max(CorrCountYvect),label='Y profile')
            # ax.scatter(PDX,PDY,CorrCountPDvect/np.max(CorrCountPDvect),label='PD profile')
            # ax.scatter(NDX,NDY,CorrCountNDvect/np.max(CorrCountNDvect),label='ND profile')





            #now we load the following reference file to establish the comparison and calculate the conformity index


            # refdata =np.loadtxt(refname,dtype=float,delimiter='\t')
            print(refname)
            refdata =np.loadtxt(refname,dtype=float,delimiter='\t')
            CorrCountXref = refdata[0:63]
            CorrCountYref = refdata[63:128]
            CorrCountNDref = refdata[128:191]
            CorrCountPDref = refdata[191:255]


            # ax.scatter(X, np.zeros(len(X)), CorrCountXref / np.max(CorrCountXref), label='Xref profile')
            # ax.scatter(np.zeros(len(Y)), Y, CorrCountYref / np.max(CorrCountYref), label='Yref profile')
            # ax.scatter(PDX, PDY, CorrCountPDref / np.max(CorrCountPDref), label='PDref profile')
            # ax.scatter(NDX, NDY, CorrCountNDref / np.max(CorrCountNDref), label='NDref profile')
            # ax.set_xlabel('X distance [cm]')
            # ax.set_ylabel('Y distance [cm]')
            # ax.legend(loc='upper left')
            # # ax.set_title('Energy Mode = '+str(energy)+beamtype)
            # ax.set_title(filename)
            # plt.show()
            #END FIGURE

            # We use these central values
            central_value_X = CorrCountXvect[len(CorrCountXvect) // 2]
            central_value_Y = CorrCountYvect[len(CorrCountYvect) // 2]
            central_value_Xref = CorrCountXref[len(CorrCountXref) // 2]
            central_value_Yref = CorrCountYref[len(CorrCountYref) // 2]


            # print(central_value_X,central_value_Y,central_value_Xref,central_value_Yref)
            # XdiffAve, XdiffStd, XCI, XdiffMax, XdiffMin = CIcalc(CorrCountXvect/np.max(CorrCountXvect),CorrCountXref/np.max(CorrCountXref))
            XdiffAve, XdiffStd, XCI, XdiffMax, XdiffMin = CIcalc(CorrCountXvect/central_value_X,CorrCountXref/central_value_Xref,16,47)
            YdiffAve, YdiffStd, YCI, YdiffMax, YdiffMin = CIcalc(CorrCountYvect/central_value_Y,CorrCountYref/central_value_Yref,16,49)
            print('X','Diff_Ave', 'Diff_Std', 'CI', 'Diff_Max', 'Diff_Min')
            print('X',XdiffAve, XdiffStd, XCI, XdiffMax, XdiffMin)
            print('Y', 'Diff_Ave', 'Diff_Std', 'CI', 'Diff_Max', 'Diff_Min')
            print('Y', YdiffAve, YdiffStd, YCI, YdiffMax, YdiffMin)

            CellChange, Attributes = xls_cell_spec(energy, gantry, coll, wedge,slopeY,XdiffAve,XdiffStd,XCI,XdiffMax,XdiffMin,YdiffAve,YdiffStd,YCI,YdiffMax,YdiffMin,central_value_Xvect)
            print(CellChange,Attributes)
            for i in range(0, len(CellChange)):
                print(ws1[list(CellChange.keys())[i]],list(CellChange.values())[i])
                ws1[list(CellChange.keys())[i]]=list(CellChange.values())[i]

    wb2.save(excel)
    exit(0)






    return CellChange,energy, Attributes








if __name__ == "__main__":
    parser = argparse.ArgumentParser()  # pylint: disable = invalid-name
    parser.add_argument("directory", help="path to directory")
    parser.add_argument("-x", "--excel", help="path to excel file")
    # parser.add_argument("-r", "--reference", help="path to reference file")
    args = parser.parse_args()  # pylint: disable = invalid-name

    if args.directory:
        dirname = args.directory  # pylint: disable = invalid-name
        if args.excel:
            excelfilename = args.excel  # pylint: disable = invalid-name
            read_icp(dirname,excelfilename)
        else:
            read_icp(dirname,None)













