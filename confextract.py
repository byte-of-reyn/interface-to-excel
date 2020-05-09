"""
To Do:
- Capability for multiple configuration lines when allowing VLANs via a trunk
- Extraction of L2 VLANs
- Extraction of L3 port information
- Extraction of static routes
- 'show int status' parsing
"""

import sys
import os.path
import csv
import string
import re
import datetime

from xlsxwriter.workbook import Workbook
from datetime import datetime

VERSION = '0.02'

#excel header list
excel_head = [
    'switch_desc',
    'switch_speed',
    'switch_duplex',
    'switch_status',
    'switch_mode',
    'switch_access_vlan',
    'switch_trunk_native_vlan',
    'switch_trunk_allow_vlan',
    'switch_span_type',
    'switch_ch_group'
]

#init excel strings, ready for entry
def_port_vals = {   
    excel_head[0]   :   'n/a',
    excel_head[1]   :   'n/a',  
    excel_head[2]   :   'n/a', 
    excel_head[3]   :   'n/a',  
    excel_head[4]   :   'n/a',
    excel_head[5]   :   'n/a', 
    excel_head[6]   :   'n/a',
    excel_head[7]   :   'n/a', 
    excel_head[8]   :   'n/a',
    excel_head[9]   :   'n/a'
}

statusPatt = re.compile(r'.*[#]{1}[\s]*sh[o]{0,1}[w]{0,1} int[e]{0,1}[r]{0,1}[f]{0,1}[a]{0,1}[c]{0,1}[e]{0,1} statu[s]{0,1}[\s]*')
descPatt = re.compile(r'.*[#]{1}[\s]*sh[o]{0,1}[w]{0,1} int[e]{0,1}[r]{0,1}[f]{0,1}[a]{0,1}[c]{0,1}[e]{0,1} statu[s]{0,1}[\s]*')

file_buff = []          #buffer for configuration file
file_name = ""          #config filename
db_dev = {}             #primary database for device interfaces, vlans, VRFs, etc.
db_err = []             #contains a listing of runtime errors
expected_args = 1

"""
Checks for valid file on OS
"""
def IsValidFile(file):
    valid = None
    
    if (os.path.isfile(file)):
        valid = True
    
    return valid    
    
"""
Dumps config file into buffer line by line
"""
def BufferFile(file):
    buffer = []
    
    #file check
    if(IsValidFile(file)!=None):
        print("opening file.")
        config_file = open(file,'r')    
    else:
        print("Unable to find config file on system. Exiting program.")
        sys.exit(1)
    
    #extract config file content to buffer
    while True:
        line = config_file.readline()
        if len(line) == 0: #EOF
            break
        else:
            buffer.append(line)
            
    config_file.close()
    return buffer
    
"""
Searches array elements for string match
Return Type: Tuple
"""
def ArrStrSearch(str_var,arr_regex):
        loc = 0 #holds arr location of matched regex 
        res = False
        
        if (str_var):
            for regex in arr_regex:
                res = re.search((re.compile(regex)),str_var)
                if res:
                    try:
                        res = regex,res.group(2),loc
                    except:
                        res = regex,res.group(0),loc
                    break
                loc = loc + 1
        return res

"""
Interface extraction in preparation for parsing
"""
def IntExtr(file):
    #initial config file parse
    flag_int = False #denotes if currently in interface

    #temp holder for each interface's details
    interfaces = {}
    sub_cmd = []
    curr_int = ''
    
    #start/end int regex for search
    arr_regex = [
        r'(^[Ii]nterface) (.*\d+)',
        r'(^[Ii]nterface) (.*\d+/*\d*/*\d*)',
        r'^[\s]+.*',
        r'^([Rr]outer) ((eigrp)*(bgp)*(ospf)*) \d+',
    ]
    
    #begin reading buffer
    for line in file:
        
        #search line for commands of interest
        str_match = ArrStrSearch(line,arr_regex)
        
        #interface sub-command parse
        if(flag_int == True):
        
            #regex found
            if(str_match):
                if((str_match[2] == 0) or (str_match[2] == 1)): #line beginning with 'interface' found
                    sub_cmd = None
                    curr_int = str_match[1].strip()
                    interfaces[curr_int] = []
                #sub cmd found under interface
                elif(str_match[2] == 2): #sub-cmd found
                    if sub_cmd:
                        #print(str_match[1])
                        interfaces[curr_int].append(str_match[1])
                    else:
                        interfaces[curr_int].append(str_match[1])
                        #print("Sub-CMD Match: ",str_match[1])
                #router config found, drop out of loop
                elif(str_match[2] == 4): #start of router section located. Reset flag.
                    flag_int = False
                else: #end of int found
                    if sub_cmd:
                        interfaces[curr_int] = sub_cmd
                    flag_int = False
                    #print("Int Dropout Match: ",str_match[1])
                    
            #no regex match - end of int
            if(not str_match):
                if sub_cmd:
                    interfaces[curr_int].append(sub_cmd)
                    #print("\nInterface: ",curr_int,"\n\n", interfaces[curr_int])
                sub_cmd = None
                flag_int = False
                #print("No regex match :(  ", line)
    
        #search for next interface
        else:
            #print("No Reg Match: ", line)          
            if(str_match):
                if(str_match[0] == arr_regex[0]): #new interface found
                    #print("New interface regex:",str_match[1])  
                    curr_int = str_match[1].strip()
                    interfaces[curr_int] = []
                    flag_int = True
                    
                #router config found, drop out of loop
                elif(str_match[2] == 4): #line beginning with 'interface' found
                    #print("Log: Interface configuration parse complete")
                    break
                
    #return parsed interface buffer
    #print(interfaces)
    return interfaces

"""
Interface extraction in preparation for parsing.
Searches through defined dictionary of regexp for excel column allocation
Uses dic key as excel headers
"""
def IntParse(port_arr):
    
    #init temp holders
    interfaces = {}
    sub_cmds = {}
    curr_port = None

    
    #init regex patterns
    arr_regex = [
        '(^\s+description )(.*)',
        '(^\s+speed )(.*)',
        '(^\s+duplex )(.*)',
        '(^\s+)([no]*\s+shutdown)',
        '(^\s+switchport mode )(.*)',
        '(^\s+switchport access vlan )(.*)',
        '(^\s+switchport trunk native vlan )(.*)',
        '(^\s+switchport trunk vlan allowed )(.*)',
        '(^\s+spanning-tree port type )(.*)',
        '(^\s+channel-group )(.*)'
    ]
    for port in port_arr:
        #instantiate dictionary for this port
        interfaces[port] = {}
        for val in def_port_vals:
            interfaces[port][val] = "n/a"
        #check for valid commands and insert into port dictionary
        for item in port_arr[port]:
            cmd_match = ArrStrSearch(item,arr_regex) #pass regex_arr to search function
            if cmd_match:
                interfaces[port][excel_head[cmd_match[2]]] = cmd_match[1]

    return interfaces

"""
Check for minimum init args passed from cmd line
"""
def ArgCheck():
    if(len(sys.argv)==expected_args):
        return False
    else:
        return True

"""
Output sorted port information to .xlsx file
"""
def XLSXOutput(interfaces,out_file_name):
    #init vars for temp location in WS
    xls_col = 0
    xls_row = 0
    
    #create the new file/worksheet
    workbook = Workbook(out_file_name, {'constant_memory':True})
    worksheet = workbook.add_worksheet("device")
    
    #create headers in current worksheet
    worksheet.write_string(xls_row,xls_col,("Interface"))
    for header in excel_head:
        xls_col = xls_col + 1
        worksheet.write_string(xls_row,xls_col,(header.strip()))
    xls_row = 1
    
    #fill columns with port data
    for port,children in interfaces.items():
        xls_col = 0
        worksheet.write_string(xls_row,xls_col,(port.strip()))
        
        for header in excel_head:
            xls_col = xls_col + 1
            worksheet.write_string(xls_row,xls_col,children[header])

        #move to next row
        xls_row = xls_row + 1
    
    workbook.close()    
    
"""
Script Main
""" 
out_file_name = "CCE_01.xlsx"
 
if(ArgCheck()):
    file_name = sys.argv[1]
    out_file_name = sys.argv[2]
else:
    print("python confextract.py [input] [output]")
    sys.exit(1)

#buffer file to temp var
file_buff = BufferFile(file_name)

#interface extraction and parse for excel output
db_dev["interfaces"] = IntExtr(file_buff)

#sort interfaces for excel input
db_dev["interfaces_xls"] = IntParse(db_dev["interfaces"])

#output file to .xlsx
XLSXOutput(db_dev["interfaces_xls"],out_file_name)
