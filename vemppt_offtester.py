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

import vemppt_parser        #Text parse script
import vemppt_register      #Data register script

import logging                  #Log errors
import serial                   #Read VEDirect data via serial USB
import ctypes                   #To use MessageBoxW
import time                     #Reading periods
from datetime import datetime   #Register local dates

def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

def main():
    #Value reset
    line = 0                            #Initial line in vemppt_register
    initial_date = datetime.now()       #Date register for xls file

    #MPPT Tester values
    pid = b'PID\t0xA04B\r\n'
    fw = b'FW\t139\r\n'
    ser = b'SER#\tHQ1540TUFJ5\r\n'
    v = b'V\t13310\r\n'
    i = b'I\t1000\r\n'
    vpv = b'VPV\t79530\r\n'
    ppv = b'PPV\t26\r\n'
    cs = b'CS\t0\r\n'
    mppt = b'MPPT\t0\r\n'
    err = b'ERR\t0\r\n'
    load = b'LOAD\tON\r\n'
    h19 = b'H19\t22700\r\n'
    h20 = b'H20\t236\r\n'
    h21 = b'H21\t734\r\n'
    h22 = b'H22\t302\r\n'
    h23 = b'H23\t875\r\n'
    hsds = b'HSDS\t193\r\n'
    checksum = b'Checksum\t\t'

    #Main thread
    try:
        Mbox("MPPT Tester","Correct port communication\nData registration will begin",64)

        line = vemppt_register.xls_open(line)

        while True:
            line = vemppt_parser.parser(pid,line)
            line = vemppt_parser.parser(fw,line)
            line = vemppt_parser.parser(ser,line)
            line = vemppt_parser.parser(v,line)
            line = vemppt_parser.parser(i,line)
            line = vemppt_parser.parser(vpv,line)
            line = vemppt_parser.parser(ppv,line)
            line = vemppt_parser.parser(cs,line)
            line = vemppt_parser.parser(mppt,line)
            line = vemppt_parser.parser(err,line)
            line = vemppt_parser.parser(load,line)
            line = vemppt_parser.parser(h19,line)
            line = vemppt_parser.parser(h20,line)
            line = vemppt_parser.parser(h21,line)
            line = vemppt_parser.parser(h22,line)
            line = vemppt_parser.parser(h23,line)
            line = vemppt_parser.parser(hsds,line)
            line = vemppt_parser.parser(checksum,line)
        
    except Exception as e:
        logging.exception("An error ocurred")
        Mbox("An error ocurred","Something went wrong!",16)
    finally:
        vemppt_register.xls_close(initial_date)

if __name__ == '__main__':
    main()
