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

import vemppt_register      #Data register script

import time                     #Reading periods
from datetime import datetime   #Register local dates


cs = {
    0: 'Not charging',
    2: 'Fault',
    3: 'Bulk',
    4: 'Absorption',
    5: 'Float'
    }

err = {
    0: 'No error',
    2: 'Battery voltage too high',
    3: 'Remote temperature sensor failure',
    4: 'Remote temperature sensor failure',
    5: 'Remote temperature sensor failure (connection lost)',
    6: 'Remote battery voltage sense failure',
    7: 'Remote battery voltage sense failure',
    8: 'Remote battery voltage sense failure (connection lost)',
    17: 'Charger temperature too high',
    18: 'Charger over current',
    19: 'Charger current reversed',
    20: 'Bulk time limit exceeded',
    21: 'Current sensor issue (sensor bias/sensor broken)',
    26: 'Terminals overheated',
    28: 'Power stage issue',
    33: 'Input voltage too high (solar panel)',
    34: 'Input current too high (solar panel)',
    38: 'Input shutdown (due to excessive battery voltage)',
    39: 'Input shutdown',
    65: '[Info] Communication warning',
    66: '[Info] Incompatible device',
    67: 'BMS Connection lost',
    114: 'CPU temperature too high',
    116: 'Factory calibration data lost',
    117: 'Invalid/incompatible firmware',
    119: 'User settings invalid'
    }

mppt = {
    0: 'Off',
    2: 'Voltage or current limited',
    3: 'MPP Tracker active'
    }

def converter(packet,conv):
    try:
        if conv == "cs":
            return cs[int(packet)]
        if conv == "err":
            return err[int(packet)]
        if conv == "mppt":
            return mppt[int(packet)]
    except:
        print("[!] Unrecognised value type of",conv)
        return packet

def parser(parse_line,line):
    parse_str = parse_line.decode("utf-8")
    #= Asynchronous message =#
    if ":A" in parse_str:
        print("Asynchronous message")
        return line

    #= Product ID =#
    elif "PID" in parse_str:
        if line == 1:
            parse_str = parse_str.split("\t")
            packet = parse_str[1]
            vemppt_register.xls_write(packet,10,line)
        vemppt_register.xls_write(datetime.now().strftime("%H:%M:%S"),0,line)
        return line

    #= Firmware Version =#
    elif "FW" in parse_str:
        if line == 1:
            parse_str = parse_str.split("\t")
            packet = parse_str[1]
            vemppt_register.xls_write(packet,10,line+1)
        return line

    #= Serial number =#
    elif "SER" in parse_str:
        if line == 1:
            parse_str = parse_str.split("\t")
            packet = parse_str[1]
            vemppt_register.xls_write(packet,10,line+2)
        return line

    #= Battery voltage =#
    elif "V" in parse_str and "P" not in parse_str: #For VPV and PPV cases
        parse_str = parse_str.split("\t")
        packet = float(parse_str[1]) * 0.001
        vemppt_register.xls_write(packet,1,line)
        return line

    #= Current =#
    elif "I" in parse_str and "P" not in parse_str: #For PID cases
        parse_str = parse_str.split("\t")
        packet = float(parse_str[1]) * 0.001
        vemppt_register.xls_write(packet,2,line)
        return line

    #= PV Voltage =#
    elif "VPV" in parse_str:
        parse_str = parse_str.split("\t")
        packet = float(parse_str[1]) * 0.001
        vemppt_register.xls_write(packet,3,line)
        return line

    #= PV Power =#
    elif "PPV" in parse_str:
        parse_str = parse_str.split("\t")
        packet = int(parse_str[1])
        vemppt_register.xls_write(packet,4,line)
        return line

    #= State of operation =#
    elif "CS" in parse_str:
        parse_str = parse_str.split("\t")
        packet = converter(int(parse_str[1]),"cs")
        vemppt_register.xls_write(packet,5,line)
        return line

    #= Tracker operation mode =#
    elif "MPPT" in parse_str:
        parse_str = parse_str.split("\t")
        packet = converter(int(parse_str[1]),"mppt")
        vemppt_register.xls_write(packet,8,line)
        return line

    #= Error code =#
    elif "ERR" in parse_str:
        parse_str = parse_str.split("\t")
        packet = converter(int(parse_str[1]),"err")
        vemppt_register.xls_write(packet,7,line)
        return line

    #= Load otput state (ON/OFF) =#
    elif "LOAD" in parse_str:
        parse_str = parse_str.split("\t")
        packet = (parse_str[1])
        vemppt_register.xls_write(packet,6,line)
        return line

    #= Yield total =#
    elif "H19" in parse_str:
        if line == 1:
            parse_str = parse_str.split("\t")
            packet = float(parse_str[1]) * 0.01
            vemppt_register.xls_write(packet,10,line+3)
        return line

    #= Yield today =#
    elif "H20" in parse_str:
        if line == 1:
            parse_str = parse_str.split("\t")
            packet = float(parse_str[1]) * 0.01
            vemppt_register.xls_write(packet,10,line+4)
        return line

    #= Maximum power today =#
    elif "H21" in parse_str:
        if line == 1:
            parse_str = parse_str.split("\t")
            packet = float(parse_str[1])
            vemppt_register.xls_write(packet,10,line+5)
        return line

    #= Yield yesterday =#
    elif "H22" in parse_str:
        if line == 1:
            parse_str = parse_str.split("\t")
            packet = float(parse_str[1]) * 0.01
            vemppt_register.xls_write(packet,10,line+6)
        return line

    #= Maximum power yesterday =#
    elif "H23" in parse_str:
        if line == 1:
            parse_str = parse_str.split("\t")
            packet = float(parse_str[1])
            vemppt_register.xls_write(packet,10,line+7)
        return line

    #= Day sequence number (0..364) =#
    elif "HSDS" in parse_str:
        if line == 1:
            parse_str = parse_str.split("\t")
            packet = int(parse_str[1])
            vemppt_register.xls_write(packet,10,line+8)
        return line

    #= Checksum =#
    elif "Checksum" in parse_str:
        print("Checksum identified. Reader progress:",line)
        time.sleep(1)
        line = line + 1     #Go to next line in xls (Checksum is last value)
        return line

    #= NULL =#
    else:
        print("[!] Unrecognised data:",parse_line)
        return line

if __name__ == '__main__':
    print("[!] Wrong execution!\nPlease, run main.py or test the register with offtest.py")
