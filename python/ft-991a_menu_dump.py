import serial
import time
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import csv
import threading

# ================= CONFIGURATION =================
PORT = "COM5"          # CHANGE THIS to your Enhanced COM port
BAUDRATE = 38400       # Match Menu 031 CAT RATE
TIMEOUT = 1.0
TXT_OUTPUT = "ft991a_menu_dump.txt"
CSV_OUTPUT = "ft991a_menu_dump.csv"

# ================= Menu Mapping (from FT-991A Manual) =================
# Expanded from previous; add/edit as needed from manual
MENU_MAP = {
    '001': {'name': 'AGC FAST DELAY', 'description': 'AGC-FAST decay time', 'default': '300', 'unit': 'msec'},
    '002': {'name': 'AGC MID DELAY', 'description': 'AGC-MID decay time', 'default': '700', 'unit': 'msec'},
    '003': {'name': 'AGC SLOW DELAY', 'description': 'AGC-SLOW decay time', 'default': '3000', 'unit': 'msec'},
    '004': {'name': 'HOME FUNCTION', 'description': 'HOME screen display', 'default': '0', 'value_map': {'0': 'SCOPE', '1': 'FUNCTION'}},
    '005': {'name': 'MY CALL INDICATION', 'description': 'MY CALL display duration', 'default': '1', 'unit': 'sec', 'value_map': {'0': 'OFF'}},
    '006': {'name': 'DISPLAY COLOR', 'description': 'VFO-A frequency field color', 'default': '0', 'value_map': {'0': 'BLUE', '1': 'GRAY', '2': 'GREEN', '3': 'ORANGE', '4': 'PURPLE', '5': 'RED', '6': 'SKY BLUE'}},
    '007': {'name': 'DIMMER LED', 'description': 'Key LED brightness', 'default': '2'},
    '008': {'name': 'DIMMER TFT', 'description': 'TFT display brightness', 'default': '8'},
    '009': {'name': 'BAR MTR PEAK HOLD', 'description': 'Bar meter peak hold time', 'default': '0', 'value_map': {'0': 'OFF', '1': '0.5', '2': '1.0', '3': '2.0'}},
    '010': {'name': 'DVS RX OUT LEVEL', 'description': 'Voice memory monitoring level', 'default': '050'},
    '011': {'name': 'DVS TX OUT LEVEL', 'description': 'Mic output level for voice memory', 'default': '050'},
# ... (keep all previous entries from your last version; truncated here for brevity)
    '012': {'name': 'KEYER TYPE', 'description': 'Keyer mode', 'default': '3', 'value_map': {'0': 'OFF', '1': 'BUG', '2': 'ELEKEY-A', '3': 'ELEKEY-B', '4': 'ELEKEY-Y', '5': 'ACS'}},
    '013': {'name': 'KEYER DOT/DASH', 'description': 'Reverse CW paddle connections', 'default': '0', 'value_map': {'0': 'NOR', '1': 'REV'}},
    '014': {'name': 'CW WEIGHT', 'description': 'CW keying weight', 'default': '32', 'value_map': None},  # Raw is like '32' for 3.2
    '015': {'name': 'BEACON INTERVAL', 'description': 'Beacon repeat interval', 'default': '000', 'unit': 'sec', 'value_map': {'000': 'OFF'}},
    '016': {'name': 'NUMBER STYLE', 'description': 'Contest number format', 'default': '0', 'value_map': {'0': '1290', '1': 'AUNO', '2': 'AUNT', '3': 'A2NO', '4': 'A2NT', '5': '12NO', '6': '12NT'}},
    '017': {'name': 'CONTEST NUMBER', 'description': 'Contest number (Morse)', 'default': '0001', 'value_map': None},
    '018': {'name': 'CW MEMORY 1', 'description': 'Memory entry method', 'default': '1', 'value_map': {'0': 'TEXT', '1': 'MESSAGE'}},
    '019': {'name': 'CW MEMORY 2', 'description': 'Memory entry method', 'default': '1', 'value_map': {'0': 'TEXT', '1': 'MESSAGE'}},
    '020': {'name': 'CW MEMORY 3', 'description': 'Memory entry method', 'default': '1', 'value_map': {'0': 'TEXT', '1': 'MESSAGE'}},
    '021': {'name': 'CW MEMORY 4', 'description': 'Memory entry method', 'default': '1', 'value_map': {'0': 'TEXT', '1': 'MESSAGE'}},
    '022': {'name': 'CW MEMORY 5', 'description': 'Memory entry method', 'default': '1', 'value_map': {'0': 'TEXT', '1': 'MESSAGE'}},
    '023': {'name': 'NB WIDTH', 'description': 'NB pulse duration', 'default': '2', 'value_map': {'0': '1 msec', '1': '3 msec', '2': '10 msec'}},
    '024': {'name': 'NB REJECTION', 'description': 'Noise attenuation level', 'default': '1', 'value_map': {'0': '10 dB', '1': '30 dB', '2': '50 dB'}},
    '025': {'name': 'NB LEVEL', 'description': 'NB strength', 'default': '02', 'value_map': None},
    '026': {'name': 'BEEP LEVEL', 'description': 'Beep volume', 'default': '050', 'value_map': None},
    '027': {'name': 'TIME ZONE', 'description': 'UTC offset', 'default': '000', 'value_map': None},  # e.g., '-0700' = -07:00
    '028': {'name': 'GPS/232C SELECT', 'description': 'GPS/CAT jack function', 'default': '3', 'value_map': {'0': 'GPS1', '1': 'GPS2', '2': 'RS232C'}},
    '029': {'name': '232C RATE', 'description': 'RS-232C baud rate', 'default': '0', 'value_map': {'0': '4800', '1': '9600', '2': '19200', '3': '38400'}},
    '030': {'name': '232C TOT', 'description': 'RS-232C timeout', 'default': '0', 'value_map': {'0': '10 msec', '1': '100 msec', '2': '1000 msec', '3': '3000 msec'}},
    '031': {'name': 'CAT RATE', 'description': 'CAT baud rate', 'default': '3', 'value_map': {'0': '4800', '1': '9600', '2': '19200', '3': '38400'}},
    '032': {'name': 'CAT TOT', 'description': 'CAT timeout', 'default': '2', 'value_map': {'0': '10 msec', '1': '100 msec', '2': '1000 msec', '3': '3000 msec'}},
    '033': {'name': 'CAT RTS', 'description': 'CAT RTS flow control', 'default': '1', 'value_map': {'0': 'DISABLE', '1': 'ENABLE'}},
    '034': {'name': 'MEM GROUP', 'description': 'Memory group function', 'default': '0', 'value_map': {'0': 'DISABLE', '1': 'ENABLE'}},
    '035': {'name': 'QUICK SPLIT FREQ', 'description': 'Offset for Quick Split', 'default': '+05', 'unit': 'kHz', 'value_map': None},
    '036': {'name': 'TX TOT', 'description': 'Transmit timeout', 'default': '00', 'unit': 'min', 'value_map': {'00': 'OFF'}},
    '037': {'name': 'MIC SCAN', 'description': 'Mic auto-scan', 'default': '1', 'value_map': {'0': 'DISABLE', '1': 'ENABLE'}},
    '038': {'name': 'MIC SCAN RESUME', 'description': 'Scan resume mode', 'default': '1', 'value_map': {'0': 'PAUSE', '1': 'TIME'}},
    '039': {'name': 'REF FREQ ADJ', 'description': 'Reference oscillator adjustment', 'default': '+00', 'value_map': None},
    '040': {'name': 'CLAR MODE SELECT', 'description': 'Clarifier mode', 'default': '0', 'value_map': {'0': 'RX', '1': 'TX', '2': 'TRX'}},
    '041': {'name': 'AM LCUT FREQ', 'description': 'AM low-cut filter frequency', 'default': '00', 'value_map': {'00': 'OFF'}},
    '042': {'name': 'AM LCUT SLOPE', 'description': 'AM low-cut slope', 'default': '0', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '043': {'name': 'AM HCUT FREQ', 'description': 'AM high-cut filter frequency', 'default': '00', 'value_map': {'00': 'OFF'}},
    '044': {'name': 'AM HCUT SLOPE', 'description': 'AM high-cut slope', 'default': '0', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '045': {'name': 'AM MIC SELECT', 'description': 'AM mic input', 'default': '0', 'value_map': {'0': 'MIC', '1': 'REAR'}},
    '046': {'name': 'AM OUT LEVEL', 'description': 'AM audio output level', 'default': '050', 'value_map': None},
    '047': {'name': 'AM PTT SELECT', 'description': 'AM PTT control', 'default': '2', 'value_map': {'0': 'DAKY', '1': 'RTS', '2': 'DTR'}},
    '048': {'name': 'AM PORT SELECT', 'description': 'AM input port', 'default': '1', 'value_map': {'0': 'DATA', '1': 'USB'}},
    '049': {'name': 'AM DATA GAIN', 'description': 'AM input gain (REAR)', 'default': '050', 'value_map': None},
    '050': {'name': 'CW LCUT FREQ', 'description': 'CW low-cut filter frequency', 'default': '04', 'value_map': {'00': 'OFF'}},  # 04 = 250 Hz (50 Hz/step)
    '051': {'name': 'CW LCUT SLOPE', 'description': 'CW low-cut slope', 'default': '1', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '052': {'name': 'CW HCUT FREQ', 'description': 'CW high-cut filter frequency', 'default': '11', 'value_map': {'67': 'OFF'}},
    '053': {'name': 'CW HCUT SLOPE', 'description': 'CW high-cut slope', 'default': '1', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '054': {'name': 'CW OUT LEVEL', 'description': 'CW output level', 'default': '050', 'value_map': None},
    '055': {'name': 'CW AUTO MODE', 'description': 'CW keying on SSB', 'default': '0', 'value_map': {'0': 'OFF', '1': '50M', '2': 'ON'}},
    '056': {'name': 'CW BK-IN TYPE', 'description': 'CW break-in mode', 'default': '1', 'value_map': {'0': 'SEMI', '1': 'FULL'}},
    '057': {'name': 'CW BK-IN DELAY', 'description': 'CW delay time', 'default': '0140', 'unit': 'msec', 'value_map': None},
    '058': {'name': 'CW WAVE SHAPE', 'description': 'CW rise/fall time', 'default': '2', 'value_map': {'0': '2 msec', '1': '4 msec'}},
    '059': {'name': 'CW FREQ DISPLAY', 'description': 'Pitch display mode', 'default': '0', 'value_map': {'0': 'DIRECT FREQ', '1': 'PITCH OFFSET'}},
    '060': {'name': 'PC KEYING', 'description': 'Keying method for RTTY/DATA', 'default': '0', 'value_map': {'0': 'OFF', '1': 'DAKY', '2': 'RTS', '3': 'DTR'}},
    '061': {'name': 'QSK DELAY TIME', 'description': 'Key-up delay for QSK', 'default': '0', 'value_map': {'0': '15 msec', '1': '20 msec', '2': '25 msec', '3': '30 msec'}},
    '062': {'name': 'DATA MODE', 'description': 'Data scheme', 'default': '1', 'value_map': {'0': 'PSK', '1': 'OTHERS'}},
    '063': {'name': 'PSK TONE', 'description': 'PSK tone frequency', 'default': '1', 'value_map': {'0': '1000 Hz', '1': '1500 Hz', '2': '2000 Hz'}},
    '064': {'name': 'OTHER DISP (SSB)', 'description': 'Displayed offset in DATA mode', 'default': '+1500', 'unit': 'Hz'},
    '065': {'name': 'OTHER SHIFT (SSB)', 'description': 'Carrier point in DATA mode', 'default': '+1500', 'unit': 'Hz'},
    '066': {'name': 'DATA LCUT FREQ', 'description': 'DATA low-cut filter frequency', 'default': '00', 'value_map': {'00': 'OFF'}},
    '067': {'name': 'DATA LCUT SLOPE', 'description': 'DATA low-cut slope', 'default': '0', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '068': {'name': 'DATA HCUT FREQ', 'description': 'DATA high-cut filter frequency', 'default': '00', 'value_map': {'00': 'OFF'}},
    '069': {'name': 'DATA HCUT SLOPE', 'description': 'DATA high-cut slope', 'default': '0', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '070': {'name': 'DATA IN SELECT', 'description': 'DATA input jack', 'default': '1', 'value_map': {'0': 'REAR', '1': 'MIC'}},
    '071': {'name': 'DATA PTT SELECT', 'description': 'PTT control for data', 'default': '2', 'value_map': {'0': 'DAKY', '1': 'RTS', '2': 'DTR'}},
    '072': {'name': 'DATA PORT SELECT', 'description': 'Data input port', 'default': '1', 'value_map': {'0': 'DATA', '1': 'USB'}},
    '073': {'name': 'DATA OUT LEVEL', 'description': 'Data I/O level', 'default': '050', 'value_map': None},
    '074': {'name': 'FM MIC SELECT', 'description': 'FM mic input', 'default': '0', 'value_map': {'0': 'MIC', '1': 'REAR'}},
    '075': {'name': 'FM OUT LEVEL', 'description': 'FM audio output level', 'default': '050', 'value_map': None},
    '076': {'name': 'FM PKT PTT SELECT', 'description': 'PTT for FM packet', 'default': '1', 'value_map': {'0': 'DAKY', '1': 'RTS', '2': 'DTR'}},
    '077': {'name': 'FM PKT PORT SELECT', 'description': 'FM packet input port', 'default': '1', 'value_map': {'0': 'DATA', '1': 'USB'}},
    '078': {'name': 'FM PKT TX GAIN', 'description': 'FM packet TX gain', 'default': '050', 'value_map': None},
    '079': {'name': 'FM PKT MODE', 'description': 'FM packet baud rate', 'default': '0', 'value_map': {'0': '1200', '1': '9600'}},
    '080': {'name': 'RPT SHIFT 28MHz', 'description': 'Repeater offset for 28 MHz', 'default': '0100', 'unit': 'kHz', 'value_map': None},
    '081': {'name': 'RPT SHIFT 50MHz', 'description': 'Repeater offset for 50 MHz', 'default': '0500', 'unit': 'kHz', 'value_map': None},
    '082': {'name': 'RPT SHIFT 144MHz', 'description': 'Repeater offset for 144 MHz', 'default': '4000', 'unit': 'kHz', 'value_map': None},
    '083': {'name': 'RPT SHIFT 430MHz', 'description': 'Repeater offset for 430 MHz', 'default': '05000', 'unit': 'kHz', 'value_map': None},
    '084': {'name': 'ARS 144MHz', 'description': 'Automatic Repeater Shift on 144 MHz', 'default': '0', 'value_map': {'0': 'OFF', '1': 'ON'}},
    '085': {'name': 'ARS 430MHz', 'description': 'Automatic Repeater Shift on 430 MHz', 'default': '1', 'value_map': {'0': 'OFF', '1': 'ON'}},
    '086': {'name': 'DCS POLARITY', 'description': 'DCS code polarity', 'default': '0', 'value_map': {'0': 'Tn-Rn', '1': 'Tn-Riv', '2': 'Tiv-Rn', '3': 'Tiv-Riv'}},
    '087': {'name': 'RADIO ID', 'description': 'Unique transceiver ID (read-only)', 'default': 'G0lYb', 'value_map': None},
    '088': {'name': 'GM DISPLAY', 'description': 'Group members sort mode', 'default': '0', 'value_map': {'0': 'DISTANCE', '1': 'STRENGTH'}},
    '089': {'name': 'DISTANCE', 'description': 'Distance unit for GM', 'default': '1', 'value_map': {'0': 'km', '1': 'mile'}},
    '090': {'name': 'AMS TX MODE', 'description': 'AMS transmission behavior', 'default': '0', 'value_map': {'0': 'AUTO', '1': 'MANUAL', '2': 'DN', '3': 'VW', '4': 'ANALOG'}},
    '091': {'name': 'STANDBY BEEP', 'description': 'Beep when digital contact ends', 'default': '1', 'value_map': {'0': 'OFF', '1': 'ON'}},
    '092': {'name': 'RTTY LCUT FREQ', 'description': 'RTTY low-cut filter frequency', 'default': '05', 'value_map': {'00': 'OFF'}},
    '093': {'name': 'RTTY LCUT SLOPE', 'description': 'RTTY low-cut slope', 'default': '1', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '094': {'name': 'RTTY HCUT FREQ', 'description': 'RTTY high-cut filter frequency', 'default': '47', 'value_map': {'67': 'OFF'}},
    '095': {'name': 'RTTY HCUT SLOPE', 'description': 'RTTY high-cut slope', 'default': '1', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '096': {'name': 'RTTY SHIFT PORT', 'description': 'Input jack for RTTY shift', 'default': '1', 'value_map': {'0': 'SHIFT', '1': 'DTR', '2': 'RTS'}},
    '097': {'name': 'RTTY POLARITY-RX', 'description': 'RTTY receive shift direction', 'default': '0', 'value_map': {'0': 'NOR', '1': 'REV'}},
    '098': {'name': 'RTTY POLARITY-TX', 'description': 'RTTY transmit shift direction', 'default': '0', 'value_map': {'0': 'NOR', '1': 'REV'}},
    '099': {'name': 'RTTY OUT LEVEL', 'description': 'Data output level in RTTY', 'default': '050', 'value_map': None},
    '100': {'name': 'RTTY SHIFT FREQ', 'description': 'RTTY frequency shift width', 'default': '0', 'value_map': {'0': '170 Hz', '1': '200 Hz', '2': '425 Hz', '3': '850 Hz'}},
    '101': {'name': 'RTTY MARK FREQ', 'description': 'RTTY mark frequency', 'default': '1', 'value_map': {'0': '1275 Hz', '1': '2125 Hz'}},
    '102': {'name': 'SSB LCUT FREQ', 'description': 'SSB low-cut filter frequency', 'default': '01', 'value_map': {'00': 'OFF'}},
    '103': {'name': 'SSB LCUT SLOPE', 'description': 'SSB low-cut slope', 'default': '0', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '104': {'name': 'SSB HCUT FREQ', 'description': 'SSB high-cut filter frequency', 'default': '67', 'value_map': {'67': 'OFF'}},
    '105': {'name': 'SSB HCUT SLOPE', 'description': 'SSB high-cut slope', 'default': '0', 'value_map': {'0': '6 dB/oct', '1': '18 dB/oct'}},
    '106': {'name': 'SSB MIC SELECT', 'description': 'SSB mic input', 'default': '0', 'value_map': {'0': 'MIC', '1': 'REAR'}},
    '107': {'name': 'SSB OUT LEVEL', 'description': 'SSB audio output level', 'default': '050', 'value_map': None},
    '108': {'name': 'SSB PTT SELECT', 'description': 'SSB PTT control', 'default': '2', 'value_map': {'0': 'DAKY', '1': 'RTS', '2': 'DTR'}},
    '109': {'name': 'SSB PORT SELECT', 'description': 'SSB input port', 'default': '1', 'value_map': {'0': 'DATA', '1': 'USB'}},
    '110': {'name': 'SSB TX BPF', 'description': 'DSP bandpass filter for SSB TX', 'default': '3', 'value_map': {'0': '100-3000 Hz', '1': '100-2900 Hz', '2': '200-2800 Hz', '3': '300-2700 Hz', '4': '400-2600 Hz'}},
    '111': {'name': 'APF WIDTH', 'description': 'Audio peak filter bandwidth in CW', 'default': '1', 'value_map': {'0': 'NARROW', '1': 'MEDIUM', '2': 'WIDE'}},
    '112': {'name': 'CONTOUR LEVEL', 'description': 'Gain of audio contour circuit', 'default': '-15', 'value_map': None},
    '113': {'name': 'CONTOUR WIDTH', 'description': 'Bandwidth (Q) of contour circuit', 'default': '10', 'value_map': None},
    '114': {'name': 'IF NOTCH WIDTH', 'description': 'Bandwidth of DSP IF notch filter', 'default': '1', 'value_map': {'0': 'NARROW', '1': 'WIDE'}},
    '115': {'name': 'SCP DISPLAY MODE', 'description': 'Scope display type', 'default': '1', 'value_map': {'0': 'SPECTRUM', '1': 'WATER FALL'}},
    '116': {'name': 'SCP SPAN FREQ', 'description': 'Frequency span of spectrum scope', 'default': '03', 'value_map': {'0': '50 kHz', '1': '100 kHz', '2': '200 kHz', '3': '500 kHz', '4': '1000 kHz'}},
    '117': {'name': 'SPECTRUM COLOR', 'description': 'Color for spectrum scope', 'default': '2', 'value_map': {'0': 'BLUE', '1': 'GRAY', '2': 'GREEN', '3': 'ORANGE', '4': 'PURPLE', '5': 'RED', '6': 'SKY BLUE'}},
    '118': {'name': 'WATER FALL COLOR', 'description': 'Color for waterfall display', 'default': '7', 'value_map': {'0': 'BLUE', '1': 'GRAY', '2': 'GREEN', '3': 'ORANGE', '4': 'PURPLE', '5': 'RED', '6': 'SKY BLUE', '7': 'MULTI'}},
    '119': {'name': 'PRMTRC EQ1 FREQ', 'description': 'Center frequency for low band parametric mic EQ', 'default': '05', 'value_map': {'00': 'OFF'}},
    '120': {'name': 'PRMTRC EQ1 LEVEL', 'description': 'Gain for low band parametric mic EQ', 'default': '+09', 'unit': 'dB', 'value_map': None},
    '121': {'name': 'PRMTRC EQ1 BWTH', 'description': 'Bandwidth for low band parametric mic EQ', 'default': '08', 'value_map': None},
    '122': {'name': 'PRMTRC EQ2 FREQ', 'description': 'Center frequency for mid band parametric mic EQ', 'default': '04', 'value_map': {'00': 'OFF'}},
    '123': {'name': 'PRMTRC EQ2 LEVEL', 'description': 'Gain for mid band parametric mic EQ', 'default': '+07', 'unit': 'dB', 'value_map': None},
    '124': {'name': 'PRMTRC EQ2 BWTH', 'description': 'Bandwidth for mid band parametric mic EQ', 'default': '08', 'value_map': None},  # From patterns in manual
    '125': {'name': 'PRMTRC EQ3 FREQ', 'description': 'Center frequency for high band parametric mic EQ', 'default': '08', 'value_map': {'00': 'OFF'}},
    '126': {'name': 'PRMTRC EQ3 LEVEL', 'description': 'Gain for high band parametric mic EQ', 'default': '+05', 'unit': 'dB', 'value_map': None},
    '127': {'name': 'PRMTRC EQ3 BWTH', 'description': 'Bandwidth for high band parametric mic EQ', 'default': '04', 'value_map': None},
    '128': {'name': 'P PRMTRC EQ1 FREQ', 'description': 'Center frequency for low band parametric EQ (processor)', 'default': '01', 'value_map': {'00': 'OFF'}},
    '129': {'name': 'P PRMTRC EQ1 LEVEL', 'description': 'Gain for low band parametric EQ (processor)', 'default': '+10', 'unit': 'dB', 'value_map': None},
    '130': {'name': 'P PRMTRC EQ1 BWTH', 'description': 'Bandwidth for low band parametric EQ (processor)', 'default': '03', 'value_map': None},
    '131': {'name': 'P PRMTRC EQ2 FREQ', 'description': 'Center frequency for mid band parametric EQ (processor)', 'default': '01', 'value_map': {'00': 'OFF'}},
    '132': {'name': 'P PRMTRC EQ2 LEVEL', 'description': 'Gain for mid band parametric EQ (processor)', 'default': '-03', 'unit': 'dB', 'value_map': None},
    '133': {'name': 'P PRMTRC EQ2 BWTH', 'description': 'Bandwidth for mid band parametric EQ (processor)', 'default': '07', 'value_map': None},
    '134': {'name': 'P PRMTRC EQ3 FREQ', 'description': 'Center frequency for high band parametric EQ (processor)', 'default': '18', 'value_map': {'00': 'OFF'}},
    '135': {'name': 'P PRMTRC EQ3 LEVEL', 'description': 'Gain for high band parametric EQ (processor)', 'default': '+10', 'unit': 'dB', 'value_map': None},
    '136': {'name': 'P PRMTRC EQ3 BWTH', 'description': 'Bandwidth for high band parametric EQ (processor)', 'default': '02', 'value_map': None},
    '137': {'name': 'SPEECH PROC LEVEL', 'description': 'Speech processor level', 'default': '100', 'value_map': None},
    '138': {'name': 'RF POWER SETTING', 'description': 'RF power output', 'default': '100', 'unit': 'W', 'value_map': None},
    '139': {'name': 'MIC GAIN', 'description': 'Microphone gain', 'default': '050', 'value_map': None},
    '140': {'name': 'PROC LEVEL', 'description': 'Processor level', 'default': '050', 'value_map': None},
    '141': {'name': 'MONITOR SELECT', 'description': 'Monitor select', 'default': '2', 'value_map': None},
    '142': {'name': 'MONITOR LEVEL', 'description': 'Monitor level', 'default': '1', 'value_map': None},
    '143': {'name': 'VOX SELECT', 'description': 'VOX input select', 'default': '010', 'value_map': {'0': 'MIC', '1': 'DATA'}},
    '144': {'name': 'VOX GAIN', 'description': 'VOX gain', 'default': '0500', 'value_map': None},
    '145': {'name': 'VOX DELAY', 'description': 'VOX delay time', 'default': '050', 'unit': 'msec', 'value_map': None},
    '146': {'name': 'ANTI VOX GAIN', 'description': 'Anti-VOX gain', 'default': '050', 'value_map': None},
    '147': {'name': 'DATA VOX GAIN', 'description': 'Data VOX gain', 'default': '0100', 'value_map': None},
    '148': {'name': 'DATA VOX DELAY', 'description': 'Data VOX delay', 'default': '000', 'unit': 'msec', 'value_map': None},
    '149': {'name': 'ANTI DVOX GAIN', 'description': 'Anti-Data VOX gain', 'default': '0', 'value_map': None},
    '150': {'name': 'PRT/WIRES FREQ', 'description': 'WIRES-X frequency setting method', 'default': '0', 'value_map': {'0': 'MANUAL', '1': 'PRESET'}},
    '151': {'name': 'PRESET FREQUENCY', 'description': 'Preset WIRES-X frequency', 'default': '14655000', 'unit': 'Hz', 'value_map': None},
    '152': {'name': 'SEARCH SETUP', 'description': 'WIRES-X search mode', 'default': '0', 'value_map': {'0': 'HISTORY', '1': 'ACTIVITY'}},
    '153': {'name': 'WIRES DG-ID', 'description': 'WIRES-X Digital Group ID', 'default': '000', 'value_map': {'000': 'AUTO'}},
    '154': {'name': 'Unknown/Reserved', 'description': 'Possibly reserved or undocumented', 'default': '???', 'value_map': None},  # No info; shows raw
    # Add remaining from your dump/manual as needed
}

# ================= Helper Functions =================
def get_mapped_value(menu_str, raw_value):
    info = MENU_MAP.get(menu_str, {'name': 'Unknown', 'description': '', 'default': 'N/A', 'unit': '', 'value_map': None})
    name = info['name']
    desc = info['description']
    default = info['default']
    unit = info.get('unit', '')
    value_map = info.get('value_map')
    
    translated = value_map.get(raw_value, raw_value) if value_map else raw_value
    if unit:
        translated += f' {unit}'
    
    return {
        'menu': menu_str,
        'raw': raw_value,
        'name': name,
        'translated': translated,
        'default': default,
        'desc': desc
    }

def send_command(ser, cmd):
    full_cmd = cmd.upper() + ";" if not cmd.endswith(";") else cmd.upper()
    ser.write(full_cmd.encode('ascii'))
    ser.flush()
    time.sleep(0.4)
    
    response = ""
    start = time.time()
    while time.time() - start < 2.5:
        if ser.in_waiting > 0:
            response += ser.read(ser.in_waiting).decode('ascii', errors='ignore')
        time.sleep(0.05)
    return response.strip()
# Add these helper functions near the top (after send_command)
def get_frequency(ser):
    resp = send_command(ser, "FA")
    if resp.startswith("FA") and resp.endswith(";"):
        freq_str = resp[2:-1]  # e.g. "014250000"
        try:
            freq_mhz = int(freq_str) / 1_000_000.0
            return f"{freq_mhz:.6f}"
        except:
            return "Error"
    return "???"

# Update the get_mode function
def get_mode(ser):
    resp = send_command(ser, "MD")
    if resp.startswith("MD") and resp.endswith(";"):
        code = resp[2:-1]  # e.g., "0A"
        modes = {
            "01": "LSB",
            "02": "USB",
            "03": "CW",
            "04": "FM",
            "05": "AM",
            "06": "FSK (RTTY-LSB)",
            "07": "CW-R (CW-LSB)",
            "08": "DATA-L (Data LSB)",
            "09": "FSK-R (RTTY-USB)",
            "0A": "DATA-U (Data USB)",
            "0B": "FM-N (Narrow FM)",
            "0C": "DATA-FM",
            "0D": "AM-N (Narrow AM)",
            "0E": "DATA-FM-N (Narrow Data FM)"
        }
        return modes.get(code.upper(), f"Unknown ({code})")
    return "???"

        
# ================= GUI Class =================
class FT991AGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("FT-991A CAT Control & Menu Dumper")
        self.root.geometry("700x600")
        
        self.ser = None
        
        # Frames
        top_frame = tk.Frame(root, padx=10, pady=10)
        top_frame.pack(fill=tk.X)
        
        control_frame = tk.Frame(root, padx=10, pady=5)
        control_frame.pack(fill=tk.X)
        
        status_frame = tk.Frame(root, padx=10, pady=5)
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        # Status labels (freq/mode from earlier)
        tk.Label(top_frame, text="Frequency (MHz):", font=("Arial", 12)).grid(row=0, column=0, sticky="w", pady=5)
        self.freq_label = tk.Label(top_frame, text="---", font=("Arial", 14, "bold"), fg="blue")
        self.freq_label.grid(row=0, column=1, sticky="w", pady=5)
        
        tk.Label(top_frame, text="Mode:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", pady=5)
        self.mode_label = tk.Label(top_frame, text="---", font=("Arial", 14, "bold"), fg="green")
        self.mode_label.grid(row=1, column=1, sticky="w", pady=5)
        
        tk.Button(top_frame, text="Refresh Status", command=self.refresh_status).grid(row=2, column=0, columnspan=2, pady=10)
        
        # Dump buttons
        tk.Button(control_frame, text="Dump Menus to TXT", command=self.start_dump_txt).grid(row=0, column=0, padx=5, pady=5)
        tk.Button(control_frame, text="Export Menus to CSV", command=self.start_export_csv).grid(row=0, column=1, padx=5, pady=5)
        
        # Status log
        self.status_text = scrolledtext.ScrolledText(status_frame, height=20, width=80, font=("Consolas", 10))
        self.status_text.pack(fill=tk.BOTH, expand=True)
        self.log("GUI initialized. Attempting radio connection....")

#        self.connect_radio()
        # Delay connection slightly or call after mainloop starts
        self.root.after(100, self.connect_radio)  # runs after 100 ms
        
#    def log(self, msg):
#        self.status_text.insert(tk.END, msg + "\n")
#        self.status_text.see(tk.END)
#        self.root.after(30000, self.periodic_refresh)
        
    def periodic_refresh(self):
        self.refresh_status()
        self.root.after(30000, self.periodic_refresh)

    def log(self, msg):
        if hasattr(self, 'status_text') and self.status_text:
            self.status_text.insert(tk.END, msg + "\n")
            self.status_text.see(tk.END)
        else:
            print(msg)  # fallback to console until GUI is ready
            
    def connect_radio(self):
        try:
            self.ser = serial.Serial(
                port=PORT, baudrate=BAUDRATE, timeout=TIMEOUT,
                bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                rtscts=True
            )
            self.ser.setRTS(True)
            self.log(f"Connected on {PORT} @ {BAUDRATE} baud (RTS/CTS enabled)")
            self.refresh_status()
        except Exception as e:
            self.log(f"Connection failed: {e}")
            messagebox.showerror("Error", f"Cannot connect: {e}")
    
#    def refresh_status(self):
#        if not self.ser or not self.ser.is_open:
#            self.log("Not connected.")
#            return
        # Simplified; add get_frequency/get_mode from earlier script if needed
#        self.log("Status refresh (freq/mode placeholder)")
# Replace the existing refresh_status() with this
    def refresh_status(self):
        if not self.ser or not self.ser.is_open:
            self.log("Not connected.")
            self.freq_label.config(text="---")
            self.mode_label.config(text="---")
            return
    
        try:
            freq = get_frequency(self.ser)
            mode = get_mode(self.ser)
            self.freq_label.config(text=freq if freq else "Error")
            self.mode_label.config(text=mode if mode else "Error")
            self.log(f"Updated: Freq = {freq} MHz, Mode = {mode}")
        except Exception as e:
            self.log(f"Refresh error: {e}")
            self.freq_label.config(text="Error")
            self.mode_label.config(text="Error")
            
    def perform_dump(self, export_csv=False):
        if not self.ser or not self.ser.is_open:
            self.log("Not connected to radio.")
            return
        
        results = []
        self.log("Starting menu dump (001-153)...")
        
        for menu_num in range(1, 154):  # Up to 153
            menu_str = f"{menu_num:03d}"
            self.log(f"Querying Menu {menu_str}...")
            cmd = f"EX{menu_str}"
            resp = send_command(self.ser, cmd)
            
            if resp.startswith(f"EX{menu_str}") and resp.endswith(";"):
                raw_value = resp[len(f"EX{menu_str}"):].strip().rstrip(';')
                mapped = get_mapped_value(menu_str, raw_value)
                line = f"Menu {menu_str}: {raw_value} --> {mapped['name']}: {mapped['translated']} ({mapped['desc']}. Default: {mapped['default']})"
                self.log(line)
                results.append(mapped)
            else:
                self.log(f"Menu {menu_str}: Error or no response ({resp})")
            time.sleep(0.2)
        
        self.log("Dump complete.")
        
        # Save TXT
        with open(TXT_OUTPUT, 'w', encoding='utf-8') as f:
            f.write(f"FT-991A Menu Dump - {time.strftime('%Y-%m-%d %H:%M')}\n")
            f.write("=================\n\n")
            for res in results:
                f.write(f"Menu {res['menu']}: {res['raw']} --> {res['name']}: {res['translated']} ({res['desc']}. Default: {res['default']})\n")
        self.log(f"TXT saved: {TXT_OUTPUT}")
        
        # Save CSV if requested
        if export_csv:
            with open(CSV_OUTPUT, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["Menu", "Raw Value", "Name", "Translated Value", "Default", "Description"])
                for res in results:
                    writer.writerow([res['menu'], res['raw'], res['name'], res['translated'], res['default'], res['desc']])
            self.log(f"CSV exported: {CSV_OUTPUT}")
    
    def start_dump_txt(self):
        threading.Thread(target=self.perform_dump, args=(False,), daemon=True).start()
    
    def start_export_csv(self):
        threading.Thread(target=self.perform_dump, args=(True,), daemon=True).start()

# ================= Main =================
if __name__ == "__main__":
    root = tk.Tk()
    app = FT991AGUI(root)
    root.mainloop()