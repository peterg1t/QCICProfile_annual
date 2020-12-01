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




def area_calc(profile,coord):
    # print(profile,coord)
    area = np.trapz(profile,coord)
    return area


def xls_cell_spec(mode, energy, gantry, central_value_Xvect, flatness_x, flatness_y, flatness_pd, flatness_nd, symmetry_X, symmetry_Y, symmetry_PD, symmetry_ND):

    print('mode',mode)
    print('energy',energy)
    print('gantry',gantry)
    CellChange = {}
    Attributes = []

    # cell specification for 6MV
    if mode == 'X-Ray' and energy == '6 MV' and gantry == 0:
        energy = 6
        beamtype = 'X'
        CellChange['D53'] = flatness_x  # "FlatnessX"
        CellChange['E53'] = symmetry_X  # "SymmetryX"
        CellChange['F53'] = flatness_y  # "FlatnessY"
        CellChange['G53'] = symmetry_Y  # "SymmetryY"
        CellChange['H53'] = flatness_pd  # "FlatnessPD"
        CellChange['I53'] = symmetry_PD  # "SymmetryPD"
        CellChange['J53'] = flatness_nd  # "FlatnessND"
        CellChange['K53'] = symmetry_ND  # "SymmetryND"
        CellChange['L53'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    # cell specification for 6MV
    elif mode == 'X-Ray' and energy == '6 MV' and gantry == 90:
        energy = 6
        beamtype = 'X'
        CellChange['D54'] = flatness_x  # "FlatnessX"
        CellChange['E54'] = symmetry_X  # "SymmetryX"
        CellChange['F54'] = flatness_y  # "FlatnessY"
        CellChange['G54'] = symmetry_Y  # "SymmetryY"
        CellChange['H54'] = flatness_pd  # "FlatnessPD"
        CellChange['I54'] = symmetry_PD  # "SymmetryPD"
        CellChange['J54'] = flatness_nd  # "FlatnessND"
        CellChange['K54'] = symmetry_ND  # "SymmetryND"
        CellChange['L54'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX', 'FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD', 'FlatnessND', 'SymmetryND', 'Central Reading']

        # cell specification for 6MV
    elif mode == 'X-Ray' and energy == '6 MV' and gantry == 180:
        energy = 6
        beamtype = 'X'
        CellChange['D55'] = flatness_x  # "FlatnessX"
        CellChange['E55'] = symmetry_X  # "SymmetryX"
        CellChange['F55'] = flatness_y  # "FlatnessY"
        CellChange['G55'] = symmetry_Y  # "SymmetryY"
        CellChange['H55'] = flatness_pd  # "FlatnessPD"
        CellChange['I55'] = symmetry_PD  # "SymmetryPD"
        CellChange['J55'] = flatness_nd  # "FlatnessND"
        CellChange['K55'] = symmetry_ND  # "SymmetryND"
        CellChange['L55'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX', 'FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD', 'FlatnessND',
                      'SymmetryND', 'Central Reading']

        # cell specification for 6MV
    elif mode == 'X-Ray' and energy == '6 MV' and gantry == 270:
        energy = 6
        beamtype = 'X'
        CellChange['D56'] = flatness_x  # "FlatnessX"
        CellChange['E56'] = symmetry_X  # "SymmetryX"
        CellChange['F56'] = flatness_y  # "FlatnessY"
        CellChange['G56'] = symmetry_Y  # "SymmetryY"
        CellChange['H56'] = flatness_pd  # "FlatnessPD"
        CellChange['I56'] = symmetry_PD  # "SymmetryPD"
        CellChange['J56'] = flatness_nd  # "FlatnessND"
        CellChange['K56'] = symmetry_ND  # "SymmetryND"
        CellChange['L56'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX', 'FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD', 'FlatnessND',
                      'SymmetryND', 'Central Reading']



        # cell specification for 10MV
    elif mode == 'X-Ray' and energy == '10 MV' and gantry == 0:
        energy = 10
        beamtype = 'X'
        CellChange['D57'] = flatness_x  # "FlatnessX"
        CellChange['E57'] = symmetry_X  # "SymmetryX"
        CellChange['F57'] = flatness_y  # "FlatnessY"
        CellChange['G57'] = symmetry_Y  # "SymmetryY"
        CellChange['H57'] = flatness_pd  # "FlatnessPD"
        CellChange['I57'] = symmetry_PD  # "SymmetryPD"
        CellChange['J57'] = flatness_nd  # "FlatnessND"
        CellChange['K57'] = symmetry_ND  # "SymmetryND"
        CellChange['L57'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']
        


    elif mode == 'X-Ray' and energy == '10 MV' and gantry == 90:
        energy = 10
        beamtype = 'X'
        CellChange['D58'] = flatness_x  # "FlatnessX"
        CellChange['E58'] = symmetry_X  # "SymmetryX"
        CellChange['F58'] = flatness_y  # "FlatnessY"
        CellChange['G58'] = symmetry_Y  # "SymmetryY"
        CellChange['H58'] = flatness_pd  # "FlatnessPD"
        CellChange['I58'] = symmetry_PD  # "SymmetryPD"
        CellChange['J58'] = flatness_nd  # "FlatnessND"
        CellChange['K58'] = symmetry_ND  # "SymmetryND"
        CellChange['L58'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'X-Ray' and energy == '10 MV' and gantry == 180:
        energy = 10
        beamtype = 'X'
        CellChange['D59'] = flatness_x  # "FlatnessX"
        CellChange['E59'] = symmetry_X  # "SymmetryX"
        CellChange['F59'] = flatness_y  # "FlatnessY"
        CellChange['G59'] = symmetry_Y  # "SymmetryY"
        CellChange['H59'] = flatness_pd  # "FlatnessPD"
        CellChange['I59'] = symmetry_PD  # "SymmetryPD"
        CellChange['J59'] = flatness_nd  # "FlatnessND"
        CellChange['K59'] = symmetry_ND  # "SymmetryND"
        CellChange['L59'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'X-Ray' and energy == '10 MV' and gantry == 270:
        energy = 10
        beamtype = 'X'
        CellChange['D60'] = flatness_x  # "FlatnessX"
        CellChange['E60'] = symmetry_X  # "SymmetryX"
        CellChange['F60'] = flatness_y  # "FlatnessY"
        CellChange['G60'] = symmetry_Y  # "SymmetryY"
        CellChange['H60'] = flatness_pd  # "FlatnessPD"
        CellChange['I60'] = symmetry_PD  # "SymmetryPD"
        CellChange['J60'] = flatness_nd  # "FlatnessND"
        CellChange['K60'] = symmetry_ND  # "SymmetryND"
        CellChange['L60'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']



    # cell specification for 15MV
    elif mode == 'X-Ray' and energy == '15 MV' and gantry == 0:
        energy = 15
        beamtype = 'X'
        CellChange['D61'] = flatness_x  # "FlatnessX"
        CellChange['E61'] = symmetry_X  # "SymmetryX"
        CellChange['F61'] = flatness_y  # "FlatnessY"
        CellChange['G61'] = symmetry_Y  # "SymmetryY"
        CellChange['H61'] = flatness_pd  # "FlatnessPD"
        CellChange['I61'] = symmetry_PD  # "SymmetryPD"
        CellChange['J61'] = flatness_nd  # "FlatnessND"
        CellChange['K61'] = symmetry_ND  # "SymmetryND"
        CellChange['L61'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'X-Ray' and energy == '15 MV' and gantry == 90:
        energy = 15
        beamtype = 'X'
        CellChange['D62'] = flatness_x  # "FlatnessX"
        CellChange['E62'] = symmetry_X  # "SymmetryX"
        CellChange['F62'] = flatness_y  # "FlatnessY"
        CellChange['G62'] = symmetry_Y  # "SymmetryY"
        CellChange['H62'] = flatness_pd  # "FlatnessPD"
        CellChange['I62'] = symmetry_PD  # "SymmetryPD"
        CellChange['J62'] = flatness_nd  # "FlatnessND"
        CellChange['K62'] = symmetry_ND  # "SymmetryND"
        CellChange['L62'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'X-Ray' and energy == '15 MV' and gantry == 180:
        energy = 15
        beamtype = 'X'
        CellChange['D63'] = flatness_x  # "FlatnessX"
        CellChange['E63'] = symmetry_X  # "SymmetryX"
        CellChange['F63'] = flatness_y  # "FlatnessY"
        CellChange['G63'] = symmetry_Y  # "SymmetryY"
        CellChange['H63'] = flatness_pd  # "FlatnessPD"
        CellChange['I63'] = symmetry_PD  # "SymmetryPD"
        CellChange['J63'] = flatness_nd  # "FlatnessND"
        CellChange['K63'] = symmetry_ND  # "SymmetryND"
        CellChange['L63'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'X-Ray' and energy == '15 MV' and gantry == 270:
        energy = 15
        beamtype = 'X'
        CellChange['D64'] = flatness_x  # "FlatnessX"
        CellChange['E64'] = symmetry_X  # "SymmetryX"
        CellChange['F64'] = flatness_y  # "FlatnessY"
        CellChange['G64'] = symmetry_Y  # "SymmetryY"
        CellChange['H64'] = flatness_pd  # "FlatnessPD"
        CellChange['I64'] = symmetry_PD  # "SymmetryPD"
        CellChange['J64'] = flatness_nd  # "FlatnessND"
        CellChange['K64'] = symmetry_ND  # "SymmetryND"
        CellChange['L64'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']



    # cell specification for 6FFF
    elif mode == 'X-Ray FFF' and energy == '6 MV' and gantry == 0:
        energy = 6
        beamtype = 'FFF'
        CellChange['E65'] = symmetry_X  # "SymmetryX"
        CellChange['G65'] = symmetry_Y  # "SymmetryY"
        CellChange['I65'] = symmetry_PD  # "SymmetryPD"
        CellChange['K65'] = symmetry_ND  # "SymmetryND"
        CellChange['L65'] = central_value_Xvect  # "CentralReading"
        Attributes = ['SymmetryX', 'SymmetryY', 'SymmetryPD', 'SymmetryND', 'Central Reading']


    # cell specification for 6FFF
    elif mode == 'X-Ray FFF' and energy == '6 MV' and gantry == 90:
        energy = 6
        beamtype = 'FFF'
        CellChange['E66'] = symmetry_X  # "SymmetryX"
        CellChange['G66'] = symmetry_Y  # "SymmetryY"
        CellChange['I66'] = symmetry_PD  # "SymmetryPD"
        CellChange['K66'] = symmetry_ND  # "SymmetryND"
        CellChange['L66'] = central_value_Xvect  # "CentralReading"
        Attributes = ['SymmetryX', 'SymmetryY', 'SymmetryPD', 'SymmetryND', 'Central Reading']


    # cell specification for 6FFF
    elif mode == 'X-Ray FFF' and energy == '6 MV' and gantry == 180:
        energy = 6
        beamtype = 'FFF'
        CellChange['E67'] = symmetry_X  # "SymmetryX"
        CellChange['G67'] = symmetry_Y  # "SymmetryY"
        CellChange['I67'] = symmetry_PD  # "SymmetryPD"
        CellChange['K67'] = symmetry_ND  # "SymmetryND"
        CellChange['L67'] = central_value_Xvect  # "CentralReading"
        Attributes = ['SymmetryX', 'SymmetryY', 'SymmetryPD', 'SymmetryND', 'Central Reading']

    # cell specification for 6FFF
    elif mode == 'X-Ray FFF' and energy == '6 MV' and gantry == 270:
        energy = 6
        beamtype = 'FFF'
        CellChange['E68'] = symmetry_X  # "SymmetryX"
        CellChange['G68'] = symmetry_Y  # "SymmetryY"
        CellChange['I68'] = symmetry_PD  # "SymmetryPD"
        CellChange['K68'] = symmetry_ND  # "SymmetryND"
        CellChange['L68'] = central_value_Xvect  # "CentralReading"
        Attributes = ['SymmetryX', 'SymmetryY', 'SymmetryPD', 'SymmetryND', 'Central Reading']



    # cell specification for 10FFF
    elif mode == 'X-Ray FFF' and energy == '10 MV' and gantry == 0:
        energy = 10
        beamtype = 'FFF'
        CellChange['E69'] = symmetry_X  # "SymmetryX"
        CellChange['G69'] = symmetry_Y  # "SymmetryY"
        CellChange['I69'] = symmetry_PD  # "SymmetryPD"
        CellChange['K69'] = symmetry_ND  # "SymmetryND"
        CellChange['L69'] = central_value_Xvect  # "CentralReading"
        Attributes = ['SymmetryX', 'SymmetryY', 'SymmetryPD', 'SymmetryND', 'Central Reading']


    # cell specification for 10FFF
    elif mode == 'X-Ray FFF' and energy == '10 MV' and gantry == 90:
        energy = 10
        beamtype = 'FFF'
        CellChange['E70'] = symmetry_X  # "SymmetryX"
        CellChange['G70'] = symmetry_Y  # "SymmetryY"
        CellChange['I70'] = symmetry_PD  # "SymmetryPD"
        CellChange['K70'] = symmetry_ND  # "SymmetryND"
        CellChange['L70'] = central_value_Xvect  # "CentralReading"
        Attributes = ['SymmetryX', 'SymmetryY', 'SymmetryPD', 'SymmetryND', 'Central Reading']

    # cell specification for 10FFF
    elif mode == 'X-Ray FFF' and energy == '10 MV' and gantry == 180:
        energy = 10
        beamtype = 'FFF'
        CellChange['E71'] = symmetry_X  # "SymmetryX"
        CellChange['G71'] = symmetry_Y  # "SymmetryY"
        CellChange['I71'] = symmetry_PD  # "SymmetryPD"
        CellChange['K71'] = symmetry_ND  # "SymmetryND"
        CellChange['L71'] = central_value_Xvect  # "CentralReading"
        Attributes = ['SymmetryX', 'SymmetryY', 'SymmetryPD', 'SymmetryND', 'Central Reading']

    # cell specification for 10FFF
    elif mode == 'X-Ray FFF' and energy == '10 MV' and gantry == 270:
        energy = 10
        beamtype = 'FFF'
        CellChange['E72'] = symmetry_X  # "SymmetryX"
        CellChange['G72'] = symmetry_Y  # "SymmetryY"
        CellChange['I72'] = symmetry_PD  # "SymmetryPD"
        CellChange['K72'] = symmetry_ND  # "SymmetryND"
        CellChange['L72'] = central_value_Xvect  # "CentralReading"
        Attributes = ['SymmetryX', 'SymmetryY', 'SymmetryPD', 'SymmetryND', 'Central Reading']






    # cell specification for 6MeV
    elif mode == 'Electron' and energy == '6 MeV' and gantry == 0:
        energy = 6
        beamtype = 'e'
        CellChange['D29'] = flatness_x  # "FlatnessX"
        CellChange['E29'] = symmetry_X  # "SymmetryX"
        CellChange['F29'] = flatness_y  # "FlatnessY"
        CellChange['G29'] = symmetry_Y  # "SymmetryY"
        CellChange['H29'] = flatness_pd  # "FlatnessPD"
        CellChange['I29'] = symmetry_PD  # "SymmetryPD"
        CellChange['J29'] = flatness_nd  # "FlatnessND"
        CellChange['K29'] = symmetry_ND  # "SymmetryND"
        CellChange['L29'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']


    elif mode == 'Electron' and energy == '6 MeV' and gantry == 90:
        energy = 6
        beamtype = 'e'
        CellChange['D30'] = flatness_x  # "FlatnessX"
        CellChange['E30'] = symmetry_X  # "SymmetryX"
        CellChange['F30'] = flatness_y  # "FlatnessY"
        CellChange['G30'] = symmetry_Y  # "SymmetryY"
        CellChange['H30'] = flatness_pd  # "FlatnessPD"
        CellChange['I30'] = symmetry_PD  # "SymmetryPD"
        CellChange['J30'] = flatness_nd  # "FlatnessND"
        CellChange['K30'] = symmetry_ND  # "SymmetryND"
        CellChange['L30'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '6 MeV' and gantry == 180:
        energy = 6
        beamtype = 'e'
        CellChange['D31'] = flatness_x  # "FlatnessX"
        CellChange['E31'] = symmetry_X  # "SymmetryX"
        CellChange['F31'] = flatness_y  # "FlatnessY"
        CellChange['G31'] = symmetry_Y  # "SymmetryY"
        CellChange['H31'] = flatness_pd  # "FlatnessPD"
        CellChange['I31'] = symmetry_PD  # "SymmetryPD"
        CellChange['J31'] = flatness_nd  # "FlatnessND"
        CellChange['K31'] = symmetry_ND  # "SymmetryND"
        CellChange['L31'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '6 MeV' and gantry == 270:
        energy = 6
        beamtype = 'e'
        CellChange['D32'] = flatness_x  # "FlatnessX"
        CellChange['E32'] = symmetry_X  # "SymmetryX"
        CellChange['F32'] = flatness_y  # "FlatnessY"
        CellChange['G32'] = symmetry_Y  # "SymmetryY"
        CellChange['H32'] = flatness_pd  # "FlatnessPD"
        CellChange['I32'] = symmetry_PD  # "SymmetryPD"
        CellChange['J32'] = flatness_nd  # "FlatnessND"
        CellChange['K32'] = symmetry_ND  # "SymmetryND"
        CellChange['L32'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']







    # cell specification for 9MeV
    elif mode == 'Electron' and energy == '9 MeV' and gantry == 0:
        energy = 9
        beamtype = 'e'
        CellChange['D33'] = flatness_x  # "FlatnessX"
        CellChange['E33'] = symmetry_X  # "SymmetryX"
        CellChange['F33'] = flatness_y  # "FlatnessY"
        CellChange['G33'] = symmetry_Y  # "SymmetryY"
        CellChange['H33'] = flatness_pd  # "FlatnessPD"
        CellChange['I33'] = symmetry_PD  # "SymmetryPD"
        CellChange['J33'] = flatness_nd  # "FlatnessND"
        CellChange['K33'] = symmetry_ND  # "SymmetryND"
        CellChange['L33'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']


    elif mode == 'Electron' and energy == '9 MeV' and gantry == 90:
        energy = 9
        beamtype = 'e'
        CellChange['D34'] = flatness_x  # "FlatnessX"
        CellChange['E34'] = symmetry_X  # "SymmetryX"
        CellChange['F34'] = flatness_y  # "FlatnessY"
        CellChange['G34'] = symmetry_Y  # "SymmetryY"
        CellChange['H34'] = flatness_pd  # "FlatnessPD"
        CellChange['I34'] = symmetry_PD  # "SymmetryPD"
        CellChange['J34'] = flatness_nd  # "FlatnessND"
        CellChange['K34'] = symmetry_ND  # "SymmetryND"
        CellChange['L34'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '9 MeV' and gantry == 180:
        energy = 9
        beamtype = 'e'
        CellChange['D35'] = flatness_x  # "FlatnessX"
        CellChange['E35'] = symmetry_X  # "SymmetryX"
        CellChange['F35'] = flatness_y  # "FlatnessY"
        CellChange['G35'] = symmetry_Y  # "SymmetryY"
        CellChange['H35'] = flatness_pd  # "FlatnessPD"
        CellChange['I35'] = symmetry_PD  # "SymmetryPD"
        CellChange['J35'] = flatness_nd  # "FlatnessND"
        CellChange['K35'] = symmetry_ND  # "SymmetryND"
        CellChange['L35'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '9 MeV' and gantry == 270:
        energy = 9
        beamtype = 'e'
        CellChange['D36'] = flatness_x  # "FlatnessX"
        CellChange['E36'] = symmetry_X  # "SymmetryX"
        CellChange['F36'] = flatness_y  # "FlatnessY"
        CellChange['G36'] = symmetry_Y  # "SymmetryY"
        CellChange['H36'] = flatness_pd  # "FlatnessPD"
        CellChange['I36'] = symmetry_PD  # "SymmetryPD"
        CellChange['J36'] = flatness_nd  # "FlatnessND"
        CellChange['K36'] = symmetry_ND  # "SymmetryND"
        CellChange['L36'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']










    # cell specification for 12MeV
    elif mode == 'Electron' and energy == '12 MeV' and gantry == 0:
        energy = 12
        beamtype = 'e'
        CellChange['D37'] = flatness_x  # "FlatnessX"
        CellChange['E37'] = symmetry_X  # "SymmetryX"
        CellChange['F37'] = flatness_y  # "FlatnessY"
        CellChange['G37'] = symmetry_Y  # "SymmetryY"
        CellChange['H37'] = flatness_pd  # "FlatnessPD"
        CellChange['I37'] = symmetry_PD  # "SymmetryPD"
        CellChange['J37'] = flatness_nd  # "FlatnessND"
        CellChange['K37'] = symmetry_ND  # "SymmetryND"
        CellChange['L37'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']


    elif mode == 'Electron' and energy == '12 MeV' and gantry == 90:
        energy = 12
        beamtype = 'e'
        CellChange['D38'] = flatness_x  # "FlatnessX"
        CellChange['E38'] = symmetry_X  # "SymmetryX"
        CellChange['F38'] = flatness_y  # "FlatnessY"
        CellChange['G38'] = symmetry_Y  # "SymmetryY"
        CellChange['H38'] = flatness_pd  # "FlatnessPD"
        CellChange['I38'] = symmetry_PD  # "SymmetryPD"
        CellChange['J38'] = flatness_nd  # "FlatnessND"
        CellChange['K38'] = symmetry_ND  # "SymmetryND"
        CellChange['L38'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '12 MeV' and gantry == 180:
        energy = 12
        beamtype = 'e'
        CellChange['D39'] = flatness_x  # "FlatnessX"
        CellChange['E39'] = symmetry_X  # "SymmetryX"
        CellChange['F39'] = flatness_y  # "FlatnessY"
        CellChange['G39'] = symmetry_Y  # "SymmetryY"
        CellChange['H39'] = flatness_pd  # "FlatnessPD"
        CellChange['I39'] = symmetry_PD  # "SymmetryPD"
        CellChange['J39'] = flatness_nd  # "FlatnessND"
        CellChange['K39'] = symmetry_ND  # "SymmetryND"
        CellChange['L39'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '12 MeV' and gantry == 270:
        energy = 12
        beamtype = 'e'
        CellChange['D40'] = flatness_x  # "FlatnessX"
        CellChange['E40'] = symmetry_X  # "SymmetryX"
        CellChange['F40'] = flatness_y  # "FlatnessY"
        CellChange['G40'] = symmetry_Y  # "SymmetryY"
        CellChange['H40'] = flatness_pd  # "FlatnessPD"
        CellChange['I40'] = symmetry_PD  # "SymmetryPD"
        CellChange['J40'] = flatness_nd  # "FlatnessND"
        CellChange['K40'] = symmetry_ND  # "SymmetryND"
        CellChange['L40'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']









    # cell specification for 16MeV
    elif mode == 'Electron' and energy == '16 MeV' and gantry == 0:
        energy = 16
        beamtype = 'e'
        CellChange['D41'] = flatness_x  # "FlatnessX"
        CellChange['E41'] = symmetry_X  # "SymmetryX"
        CellChange['F41'] = flatness_y  # "FlatnessY"
        CellChange['G41'] = symmetry_Y  # "SymmetryY"
        CellChange['H41'] = flatness_pd  # "FlatnessPD"
        CellChange['I41'] = symmetry_PD  # "SymmetryPD"
        CellChange['J41'] = flatness_nd  # "FlatnessND"
        CellChange['K41'] = symmetry_ND  # "SymmetryND"
        CellChange['L41'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']


    elif mode == 'Electron' and energy == '16 MeV' and gantry == 90:
        energy = 16
        beamtype = 'e'
        CellChange['D42'] = flatness_x  # "FlatnessX"
        CellChange['E42'] = symmetry_X  # "SymmetryX"
        CellChange['F42'] = flatness_y  # "FlatnessY"
        CellChange['G42'] = symmetry_Y  # "SymmetryY"
        CellChange['H42'] = flatness_pd  # "FlatnessPD"
        CellChange['I42'] = symmetry_PD  # "SymmetryPD"
        CellChange['J42'] = flatness_nd  # "FlatnessND"
        CellChange['K42'] = symmetry_ND  # "SymmetryND"
        CellChange['L42'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '16 MeV' and gantry == 180:
        energy = 16
        beamtype = 'e'
        CellChange['D43'] = flatness_x  # "FlatnessX"
        CellChange['E43'] = symmetry_X  # "SymmetryX"
        CellChange['F43'] = flatness_y  # "FlatnessY"
        CellChange['G43'] = symmetry_Y  # "SymmetryY"
        CellChange['H43'] = flatness_pd  # "FlatnessPD"
        CellChange['I43'] = symmetry_PD  # "SymmetryPD"
        CellChange['J43'] = flatness_nd  # "FlatnessND"
        CellChange['K43'] = symmetry_ND  # "SymmetryND"
        CellChange['L43'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '16 MeV' and gantry == 270:
        energy = 16
        beamtype = 'e'
        CellChange['D44'] = flatness_x  # "FlatnessX"
        CellChange['E44'] = symmetry_X  # "SymmetryX"
        CellChange['F44'] = flatness_y  # "FlatnessY"
        CellChange['G44'] = symmetry_Y  # "SymmetryY"
        CellChange['H44'] = flatness_pd  # "FlatnessPD"
        CellChange['I44'] = symmetry_PD  # "SymmetryPD"
        CellChange['J44'] = flatness_nd  # "FlatnessND"
        CellChange['K44'] = symmetry_ND  # "SymmetryND"
        CellChange['L44'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']






    # cell specification for 20MeV
    elif mode == 'Electron' and energy == '20 MeV' and gantry == 0:
        energy = 20
        beamtype = 'e'
        CellChange['D45'] = flatness_x  # "FlatnessX"
        CellChange['E45'] = symmetry_X  # "SymmetryX"
        CellChange['F45'] = flatness_y  # "FlatnessY"
        CellChange['G45'] = symmetry_Y  # "SymmetryY"
        CellChange['H45'] = flatness_pd  # "FlatnessPD"
        CellChange['I45'] = symmetry_PD  # "SymmetryPD"
        CellChange['J45'] = flatness_nd  # "FlatnessND"
        CellChange['K45'] = symmetry_ND  # "SymmetryND"
        CellChange['L45'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']


    elif mode == 'Electron' and energy == '20 MeV' and gantry == 90:
        energy = 20
        beamtype = 'e'
        CellChange['D46'] = flatness_x  # "FlatnessX"
        CellChange['E46'] = symmetry_X  # "SymmetryX"
        CellChange['F46'] = flatness_y  # "FlatnessY"
        CellChange['G46'] = symmetry_Y  # "SymmetryY"
        CellChange['H46'] = flatness_pd  # "FlatnessPD"
        CellChange['I46'] = symmetry_PD  # "SymmetryPD"
        CellChange['J46'] = flatness_nd  # "FlatnessND"
        CellChange['K46'] = symmetry_ND  # "SymmetryND"
        CellChange['L46'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '20 MeV' and gantry == 180:
        energy = 20
        beamtype = 'e'
        CellChange['D47'] = flatness_x  # "FlatnessX"
        CellChange['E47'] = symmetry_X  # "SymmetryX"
        CellChange['F47'] = flatness_y  # "FlatnessY"
        CellChange['G47'] = symmetry_Y  # "SymmetryY"
        CellChange['H47'] = flatness_pd  # "FlatnessPD"
        CellChange['I47'] = symmetry_PD  # "SymmetryPD"
        CellChange['J47'] = flatness_nd  # "FlatnessND"
        CellChange['K47'] = symmetry_ND  # "SymmetryND"
        CellChange['L47'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']

    elif mode == 'Electron' and energy == '20 MeV' and gantry == 270:
        energy = 20
        beamtype = 'e'
        CellChange['D48'] = flatness_x  # "FlatnessX"
        CellChange['E48'] = symmetry_X  # "SymmetryX"
        CellChange['F48'] = flatness_y  # "FlatnessY"
        CellChange['G48'] = symmetry_Y  # "SymmetryY"
        CellChange['H48'] = flatness_pd  # "FlatnessPD"
        CellChange['I48'] = symmetry_PD  # "SymmetryPD"
        CellChange['J48'] = flatness_nd  # "FlatnessND"
        CellChange['K48'] = symmetry_ND  # "SymmetryND"
        CellChange['L48'] = central_value_Xvect  # "CentralReading"
        Attributes = ['FlatnessX', 'SymmetryX','FlatnessY', 'SymmetryY', 'FlatnessPD', 'SymmetryPD','FlatnessND', 'SymmetryND', 'Central Reading']













    return CellChange, Attributes




def int_detc_indx(CorrCounts,FRGN):
    max_l = np.amax(CorrCounts[0:len(CorrCounts) // 2])
    max_r = np.amax(CorrCounts[len(CorrCounts) // 2:len(CorrCounts)])
    for i in range(0, len(CorrCounts) // 2):  # for the left side of the array
        if CorrCounts[i] <= max_l / 2 and CorrCounts[i + 1] > max_l / 2:
            lh = i + (max_l / 2 - CorrCounts[i]) / (CorrCounts[i + 1] - CorrCounts[i])

    for j in range(len(CorrCounts) // 2, len(CorrCounts)-1):  # for the right side of the array
        if CorrCounts[j] > max_r / 2 and CorrCounts[j + 1] <= max_r / 2:
            rh = j + (CorrCounts[j] - max_r / 2) / (CorrCounts[j] - CorrCounts[j + 1])

    CM = (lh + rh) / 2


    lFRGN = CM + (lh - CM) * FRGN / 100
    rFRGN = CM + (rh - CM) * FRGN / 100
    print("lFRGN","rFRGN")
    print(lFRGN,rFRGN)

    # lf = int(round(lFRGN)) + 1cd
    # rf = int(round(rFRGN))
    # lf = int(lFRGN)+1 # Although this is in the manual the +1 could be because of use of non-zero array start.
    lf = int(lFRGN)
    rf = int(rFRGN)

    return lf, rf, lFRGN, rFRGN, CM




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
        ws1 = wb2['Symmetry & Flatness with Gantry']

    # for subdir, dirs, files in os.walk(dirname):
    for entry in os.scandir(dirname):
        # for file in files:
        if entry.is_file():
            print('File',entry.path)
            print('Start header processing')
            f = open(entry.path, mode='r', encoding='ISO-8859-1') #,encoding='utf-8-sig')
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
            df = pd.read_csv(entry.path,skiprows=106,delimiter='\t')
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
            # print('Timetic=',Timetic*1e-6,df['TIMETIC']) # duration of the measurement
            # print('Backrate',df)



            for column in df.columns[5:68]: #this section records the X axis (-)
                CorrCountXvect.append((df[column][3] - Timetic*df[column][0])*df[column][1]/gain)#the corrected data for leakage = Timetic*Bias*Calibration/Gain
                central_value_Xvect = CorrCountXvect[len(CorrCountXvect) // 2]
                BiasX.append(df[column][0]) # already used in the formula above but saving them just in case
                CalibX.append(df[column][1])
                RawCountXvect.append(df[column][3])

            for column in df.columns[68:133]: #this section records the Y axis (|)
                CorrCountYvect.append((df[column][3] - Timetic*df[column][0])*df[column][1]/gain)
                BiasY.append(df[column][0]) # already used in the formula above but saving them just in case
                CalibY.append(df[column][1])
                RawCountYvect.append(df[column][3])

            for column in df.columns[133:196]: #this section records the D1 axis  (/)
                CorrCountPDvect.append((df[column][3] - Timetic*df[column][0])*df[column][1]/gain)
                BiasPD.append(df[column][0]) # already used in the formula above but saving them just in case
                CalibPD.append(df[column][1])
                RawCountPDvect.append(df[column][3])

            for column in df.columns[196:259]: #this section records the D2 axis  (\)
                CorrCountNDvect.append((df[column][3] - Timetic*df[column][0])*df[column][1]/gain)
                BiasND.append(df[column][0]) # already used in the formula above but saving them just in case
                CalibND.append(df[column][1])
                RawCountNDvect.append(df[column][3])




            # We use these central values
            central_value_X = CorrCountXvect[len(CorrCountXvect) // 2]
            central_value_Y = CorrCountYvect[len(CorrCountYvect) // 2]

            FRGN = 80
            xli, xri, xlFRGN, xrFRGN, CMX = int_detc_indx(CorrCountXvect, FRGN)
            yli, yri, ylFRGN, yrFRGN, CMY = int_detc_indx(CorrCountYvect, FRGN)
            pdli, pdri, pdlFRGN, pdrFRGN, CMPD = int_detc_indx(CorrCountPDvect, FRGN)
            ndli, ndri, ndlFRGN, ndrFRGN, CMND = int_detc_indx(CorrCountNDvect, FRGN)

            print('xli,xri, xlFRGN, xrFRGN,CMX')
            print(xli, xri, xlFRGN, xrFRGN, CMX)

            print('yli,yri, ylFRGN, yrFRGN,CMY')
            print(yli, yri, ylFRGN, yrFRGN, CMY)

            # here we calculate the unflatness
            central_value = float(CorrCountXvect[31])
            # print('these must be equal=',CorrCountXvect[31],CorrCountYvect[32])
            unflatness_x = float(2 * CorrCountXvect[len(CorrCountXvect) // 2] / (CorrCountXvect[8] + CorrCountXvect[
                54]))  # calculating unflatness in the Transverse - X direction (using -12 and 12)
            unflatness_y = float(2 * CorrCountYvect[len(CorrCountYvect) // 2] / (
                        CorrCountYvect[8] + CorrCountYvect[56]))  # calculating unflatness in the Radial - Y direction
            print('unflatness(x)=', unflatness_x, 'unflatness(y)=', unflatness_y)

            # flatness calculation by variance, remember these ranges are assuming a field size of 30X30
            # print(len(CorrCountXvect))
            # print(CorrCountXvect,np.amax(CorrCountXvect[8:54]),np.amin(CorrCountXvect[8:54]))
            flatness_x = 100 * (np.amax(CorrCountXvect[xli:xri + 1]) - np.amin(CorrCountXvect[xli:xri + 1])) / (
                        np.amax(CorrCountXvect[xli:xri + 1]) + np.amin(
                    CorrCountXvect[xli:xri + 1]))  # calculating flatness in the Transverse - X direction
            flatness_y = 100 * (np.amax(CorrCountYvect[yli:yri + 1]) - np.amin(CorrCountYvect[yli:yri + 1])) / (
                        np.amax(CorrCountYvect[yli:yri + 1]) + np.amin(CorrCountYvect[
                                                                       yli:yri + 1]))  # calculating flatness in the Radial - Y direction (It has a couple of more sensors in -0.5 and 0.5)
            flatness_pd = 100 * (np.amax(CorrCountPDvect[yli:yri + 1]) - np.amin(CorrCountPDvect[pdli:pdri + 1])) / (
                    np.amax(CorrCountPDvect[pdli:pdri + 1]) + np.amin(CorrCountPDvect[
                                                                   pdli:pdri + 1]))  # calculating flatness in the Radial - Y direction (It has a couple of more sensors in -0.5 and 0.5)
            flatness_nd = 100 * (np.amax(CorrCountNDvect[ndli:ndri + 1]) - np.amin(CorrCountNDvect[ndli:ndri + 1])) / (
                    np.amax(CorrCountNDvect[ndli:ndri + 1]) + np.amin(CorrCountNDvect[
                                                                   ndli:ndri + 1]))  # calculating flatness in the Radial - Y direction (It has a couple of more sensors in -0.5 and 0.5)

            print('flatness(x)=', flatness_x, 'flatness(y)=', flatness_y, 'flatness(pd)=', flatness_pd, 'flatness(nd)=', flatness_nd)

            Xi = []  # these two vectors hold the index location of each detector
            for i, v in enumerate(CorrCountXvect):
                # print(i, v)
                Xi.append(i)

            PDi = []
            for i, v in enumerate(CorrCountPDvect):
                # print(i, v)
                PDi.append(i)

            NDi = []
            for i, v in enumerate(CorrCountNDvect):
                # print(i, v)
                NDi.append(i)


            Yi = []
            for i, v in enumerate(CorrCountYvect):
                # print(i, v)
                Yi.append(i)





            # here we calculate the symmetry with the CAX Point Difference Symmetry


            for i in range(0,len(CorrCountXvect)):
                symmetry_X = (CorrCountXvect[i] - CorrCountXvect[len(CorrCountXvect)-i])/CorrCountXvect[len(CorrCountXvect)//2]*100
                print(len(CorrCountXvect)//2, i, len(CorrCountXvect)-i)
                print('symmetry_X=',symmetry_X)

            exit(0)





            # here we calculate the symmetry (This code is equivalent to SYMA - see documentation)
            # for the X
            area_R_X = area_calc(CorrCountXvect[int(CMX):xri + 1], Xi[int(CMX):xri + 1])
            area_L_X = area_calc(CorrCountXvect[xli:int(CMX) + 1], Xi[xli:int(CMX) + 1])
            mCMX = (CorrCountXvect[int(CMX) + 1] - CorrCountXvect[int(CMX)]) / (int(CMX) + 1 - int(CMX))
            fCMX = CorrCountXvect[int(CMX)] + (CMX - int(CMX)) * mCMX
            areaCMX = 1 / 2 * (fCMX + CorrCountXvect[int(CMX)]) * (CMX - int(CMX))
            ml = (CorrCountXvect[xli] - CorrCountXvect[xli - 1]) / (xli - (xli - 1))
            fxlFRGN = CorrCountXvect[xli - 1] + (xlFRGN - (xli - 1)) * ml
            areaExL = 1 / 2 * (fxlFRGN + CorrCountXvect[xli]) * (xli - xlFRGN)
            area_L_X = area_L_X + areaCMX + areaExL
            mr = (CorrCountXvect[xri + 1] - CorrCountXvect[xri]) / (xri + 1 - xri)
            fxrFRGN = CorrCountXvect[xri] + (xrFRGN - (xri)) * mr
            areaExR = 1 / 2 * (fxrFRGN + CorrCountXvect[xri]) * (xrFRGN - xri)
            area_R_X = area_R_X - areaCMX + areaExR


            symmetry_X = 200 * (area_R_X - area_L_X) / (area_L_X + area_R_X)
            print('Symmetry_X=', symmetry_X)






            # for the Y
            area_R_Y = area_calc(CorrCountYvect[int(CMY):yri + 1], Yi[int(CMY):yri + 1])
            area_L_Y = area_calc(CorrCountYvect[yli:int(CMY) + 1], Yi[yli:int(CMY) + 1])
            symmetry_Y = 200 * (area_R_Y - area_L_Y) / (area_L_Y + area_R_Y)
            mCMY = (CorrCountYvect[int(CMY) + 1] - CorrCountYvect[int(CMY)]) / (int(CMY) + 1 - int(CMY))
            fCMY = CorrCountYvect[int(CMY)] + (CMY - int(CMY)) * mCMY
            areaCMY = 1 / 2 * (fCMY + CorrCountYvect[int(CMY)]) * (CMY - int(CMY))
            ml = (CorrCountYvect[yli] - CorrCountYvect[yli - 1]) / (yli - (yli - 1))
            fylFRGN = CorrCountYvect[yli - 1] + (ylFRGN - (yli - 1)) * ml
            areaEyL = 1 / 2 * (fylFRGN + CorrCountYvect[yli]) * (yli - ylFRGN)
            area_L_Y = area_L_Y + areaCMY + areaEyL
            mr = (CorrCountYvect[yri + 1] - CorrCountYvect[yri]) / (yri + 1 - yri)
            fyrFRGN = CorrCountYvect[yri] + (yrFRGN - (yri)) * mr
            areaEyR = 1 / 2 * (fyrFRGN + CorrCountYvect[yri]) * (yrFRGN - yri)
            area_R_Y = area_R_Y - areaCMY + areaEyR


            symmetry_Y = 200 * (area_R_Y - area_L_Y) / (area_L_Y + area_R_Y)
            # symmetry_Y = 100*(CorrCountYvect[8]-CorrCountYvect[57])/CorrCountYvect[int(len(CorrCountYvect) / 2)]
            print('Symmetry_Y=', symmetry_Y)






            # for the PD
            area_R_PD = area_calc(CorrCountPDvect[int(CMPD):pdri + 1], PDi[int(CMPD):pdri + 1])
            area_L_PD = area_calc(CorrCountPDvect[pdli:int(CMPD) + 1], PDi[pdli:int(CMPD) + 1])
            symmetry_PD = 200 * (area_R_PD - area_L_PD) / (area_L_PD + area_R_PD)
            mCMPD = (CorrCountPDvect[int(CMPD) + 1] - CorrCountPDvect[int(CMPD)]) / (int(CMPD) + 1 - int(CMPD))
            fCMPD = CorrCountPDvect[int(CMPD)] + (CMPD - int(CMPD)) * mCMPD
            areaCMPD = 1 / 2 * (fCMPD + CorrCountPDvect[int(CMPD)]) * (CMPD - int(CMPD))
            ml = (CorrCountPDvect[pdli] - CorrCountPDvect[pdli - 1]) / (pdli - (pdli - 1))
            fpdlFRGN = CorrCountPDvect[pdli - 1] + (pdlFRGN - (pdli - 1)) * ml
            areaEpdL = 1 / 2 * (fpdlFRGN + CorrCountPDvect[pdli]) * (pdli - pdlFRGN)
            area_L_PD = area_L_PD + areaCMPD + areaEpdL
            mr = (CorrCountPDvect[pdri + 1] - CorrCountPDvect[pdri]) / (pdri + 1 - pdri)
            fpdrFRGN = CorrCountPDvect[pdri] + (pdrFRGN - (pdri)) * mr
            areaEpdR = 1 / 2 * (fpdrFRGN + CorrCountPDvect[pdri]) * (pdrFRGN - pdri)
            area_R_PD = area_R_PD - areaCMPD + areaEpdR

            symmetry_PD = 200 * (area_R_PD - area_L_PD) / (area_L_PD + area_R_PD)
            print('Symmetry_PD=', symmetry_PD)







            # for the ND
            area_R_ND = area_calc(CorrCountNDvect[int(CMND):ndri + 1], NDi[int(CMND):ndri + 1])
            area_L_ND = area_calc(CorrCountNDvect[ndli:int(CMND) + 1], NDi[ndli:int(CMND) + 1])
            symmetry_ND = 200 * (area_R_ND - area_L_ND) / (area_L_ND + area_R_ND)
            mCMND = (CorrCountNDvect[int(CMND) + 1] - CorrCountNDvect[int(CMND)]) / (int(CMND) + 1 - int(CMND))
            fCMND = CorrCountNDvect[int(CMND)] + (CMND - int(CMND)) * mCMND
            areaCMND = 1 / 2 * (fCMND + CorrCountNDvect[int(CMND)]) * (CMND - int(CMND))
            ml = (CorrCountNDvect[ndli] - CorrCountNDvect[ndli - 1]) / (ndli - (ndli - 1))
            fndlFRGN = CorrCountNDvect[ndli - 1] + (ndlFRGN - (ndli - 1)) * ml
            areaEndL = 1 / 2 * (fndlFRGN + CorrCountNDvect[ndli]) * (ndli - ndlFRGN)
            area_L_ND = area_L_ND + areaCMND + areaEndL
            mr = (CorrCountNDvect[ndri + 1] - CorrCountNDvect[ndri]) / (ndri + 1 - ndri)
            fndrFRGN = CorrCountNDvect[ndri] + (ndrFRGN - (ndri)) * mr
            areaEndR = 1 / 2 * (fndrFRGN + CorrCountNDvect[ndri]) * (ndrFRGN - ndri)
            area_R_ND = area_R_ND - areaCMND + areaEndR

            symmetry_ND = 200 * (area_R_ND - area_L_ND) / (area_L_ND + area_R_ND)
            print('Symmetry_ND=', symmetry_ND)

            print(central_value_Xvect)




            CellChange, Attributes = xls_cell_spec(mode, energy, int(gantry[:-3]), central_value_Xvect, flatness_x, flatness_y, flatness_pd, flatness_nd, symmetry_X, symmetry_Y, symmetry_PD, symmetry_ND)
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











