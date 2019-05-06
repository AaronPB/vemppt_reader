#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
    This file is part of vemppt_reader.

    vemppt_reader is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    vemppt_reader is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with vemppt_reader.  If not, see <https://www.gnu.org/licenses/>.
"""

import xlwt                     #Save data into a .xls file
from datetime import datetime   #Register local dates

#= Global defines =#
initial_date = datetime.now()
register_values = xlwt.Workbook(encoding="utf-8")

sheet = register_values.add_sheet(initial_date.strftime("MPPT_READ"),cell_overwrite_ok=False)

#= xls Styles =#
header_style = xlwt.easyxf("font: bold on; borders: bottom dashed")
title_style = xlwt.easyxf("font: bold on")

#= Functions =#
def xls_open(line):

    sheet.write(line, 0, "Read Time", style=header_style)
    sheet.write(line, 1, "Battery Voltage (V)", style=header_style)
    sheet.write(line, 2, "Current (A)", style=header_style)
    sheet.write(line, 3, "PV Voltage (V)", style=header_style)
    sheet.write(line, 4, "PV Power (W)", style=header_style)
    sheet.write(line, 5, "CS", style=header_style)
    sheet.write(line, 6, "LOAD", style=header_style)
    sheet.write(line, 7, "ERROR", style=header_style)
    sheet.write(line, 8, "MPPT", style=header_style)

    sheet.write(line, 9, "MPPT General Information", style=header_style)
    sheet.write(line+1, 9, "Product ID", style=title_style)
    sheet.write(line+2, 9, "Firmware version", style=title_style)
    sheet.write(line+3, 9, "Serial number", style=title_style)
    sheet.write(line+4, 9, "Yield total", style=title_style)
    sheet.write(line+4, 11, "kWh")
    sheet.write(line+5, 9, "Yield today", style=title_style)
    sheet.write(line+5, 11, "kWh")
    sheet.write(line+6, 9, "Maximum power today", style=title_style)
    sheet.write(line+6, 11, "W")
    sheet.write(line+7, 9, "Yield yesterday", style=title_style)
    sheet.write(line+7, 11, "kWh")
    sheet.write(line+8, 9, "Maximum Power yesterday", style=title_style)
    sheet.write(line+8, 11, "W")
    sheet.write(line+9, 9, "Day sequence number", style=title_style)

    line = line + 1
    return line

def xls_write(packet,column,line):
    try:
        sheet.write(line,column,packet)
    except:
        print("[!] Data not saved because it is overwriting other data!")
    return 0

def xls_close(date):
    register_values.save(date.strftime("MPPT Read %d_%m_%Y - %Hh %Mmin.xls"))
    return 0

"""
TO-DO
Asynchronous messages registration in a svg file
"""
