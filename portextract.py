"""
Script used to extract L2 port configuration from provided Cisco configuration backup files
"""
#!/usr/bin/python

import sys
import os.path
import re

from xlsxwriter.workbook import Workbook

VERSION = '0.02'

#excel header list
excel_head = [
    'port_desc',
    'port_speed',
    'port_duplex',
    'port_status',
    'port_mode',
    'port_access_vlan',
    'port_trunk_native_vlan',
    'port_trunk_allow_vlan',
    'port_stp_type',
    'port_channel_group',
    'port_mtu',
    'port_bpduguard',
    'port_portfast'
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
    excel_head[9]   :   'n/a',
    excel_head[10]   :   '1500',
    excel_head[11]   :   'n/a',
    excel_head[12]   :   'n/a'
}

db_dev = {}             #database for device interfaces, vlans, VRFs, etc.
db_err = []             #contains a listing of runtime errors

"""
Checks for file validity on native OS
"""
def is_valid_file(file):
    valid = None
    if os.path.isfile(file):
        valid = True
    return valid

"""
Dumps config file into buffer line by line
"""
def buffer_file(files):
    buffer = {}

    for file in files["in_files"]:
        #file check
        if is_valid_file(file) is not None:
            print("Performing configuration extraction from file", file, "...")
            line_buffer = ""
            try:
                config_file = open(file, 'r')
                #extract config file content to buffer
                while True:
                    line = config_file.readline()
                    if len(line) == 0:#EOF
                        break
                    else:
                        line_buffer += line

                buffer[file] = line_buffer
                config_file.close()
            except:
                print("File access error:", file)
                break
        else:
            print("Skipping file read for", file, ". Access error...")
    return buffer

"""
Searches array elements for string match
Return Type: Tuple
"""
def array_string_search(str_var, arr_regex):
    #holds arr location of matched regex
    loc = 0#
    res = False

    if str_var:
        try:
            for regex in arr_regex:
                res = re.search((re.compile(regex)), str_var)
                if res:
                    try:
                        res = regex, res.group(2), loc
                    except:
                        res = regex, res.group(0), loc
                    break
                loc = loc + 1
        except:
            pass### FIX UP ERROR RETURN ###
    return res

"""
Interface extraction in preparation for parsing
"""
def interface_extract(file_buffer):
    #initial config file parse
    flag_int = False #denotes if currently in interface

    #holder for device interface details
    devices = {}

    #start/end int regex for search
    arr_regex = [
        r'(^[Ii]nterface) (.*\d+)',
        r'(^[Ii]nterface) (.*\d+/*\d*/*\d*)',
        r'^[\s]+.*',
        r'^([Rr]outer) ((eigrp)*(bgp)*(ospf)*) \d+'
    ]

    #console status update
    print("Extracting configuration elements from provided configuration files...")

    #outer loop for each provided buffer/config file
    for filename, buffer in file_buffer.items():
        #reset temporary variables
        interfaces = {}
        sub_cmd = []
        curr_int = ''

        #Split solid config file based on newline returns
        lines = buffer.split('\n')

        #inner loop to parse buffer lines
        for line in lines:
            #search line for commands of interest
            str_match = array_string_search(line, arr_regex)

            #interface sub-command parse
            if flag_int:
                #current config line matches regex
                if str_match:
                    #line beginning with 'interface' found
                    if (str_match[2] == 0) or (str_match[2] == 1):
                        sub_cmd = None
                        curr_int = str_match[1].strip()
                        interfaces[curr_int] = []

                    #sub cmd found under interface
                    elif str_match[2] == 2: #sub-cmd found
                        if sub_cmd:
                            interfaces[curr_int].append(str_match[1])
                        else:
                            interfaces[curr_int].append(str_match[1])

                    #router config found, drop out of loop
                    elif str_match[2] == 4:
                        flag_int = False

                    else:#end of int found
                        if sub_cmd:
                            interfaces[curr_int] = sub_cmd
                        flag_int = False

                #no regex match - end of interface configuration
                if not str_match:
                    if sub_cmd:
                        interfaces[curr_int].append(sub_cmd)
                    sub_cmd = None
                    flag_int = False

            #search for next interface
            else:
                if str_match:
                    if str_match[0] == arr_regex[0]:#new interface found
                        curr_int = str_match[1].strip()
                        interfaces[curr_int] = []
                        flag_int = True

                    #router config found, drop out of loop
                    elif str_match[2] == 4:#line beginning with 'interface' found
                        break
        #place interfaces into device dictionary
        devices[filename] = interfaces
    #return parsed interface buffer
    return devices

"""
Interface extraction in preparation for parsing.
Searches through defined dictionary of regexp for excel column allocation
Uses dic key as excel headers
"""
def interface_parse(devices):
    #init temp holders
    parsed_device = {}

    #initialize regex patterns
    arr_regex = [
        r'(^\s+description )(.*)',
        r'(^\s+speed )(.*)',
        r'(^\s+duplex )(.*)',
        r'(^\s+)([no]*\s+shutdown)',
        r'(^\s+switchport mode )(.*)',
        r'(^\s+switchport access vlan )(.*)',
        r'(^\s+switchport trunk native vlan )(.*)',
        r'(^\s+switchport trunk allowed vlan )(\d+.*)',
        r'(^\s+spanning-tree port type )(.*)',
        r'(^\s+channel-group )(.*)',
        r'(^\s+mtu )(\d+)',
        r'(^\s+spanning-tree bpduguard)(enable)\s+',
        r'(^\s+spanning-tree )(portfast)\s+',
    ]

    #console status update
    print("Parsing extracted interface items...")

    #outside regex to match additional VLANs permitted over trunk
    trunk_allowed_add = r'(^\s+switchport trunk allowed vlan add )(\d+.*)'

    #outer loop to iterate through each device
    for key, port_arr in devices.items():
        #temporary holder for stripped configuration elements
        interfaces = {}

        #inner loop to iterate through the device interface array
        for port in port_arr:
            #instantiate dictionary for this port
            interfaces[port] = {}
            for val in def_port_vals:
                interfaces[port][val] = "n/a"
            #check for valid commands and insert into port dictionary
            for item in port_arr[port]:
                cmd_match = array_string_search(item, arr_regex) #pass regex_arr to search function
                #standard switchport configuration item identified
                if cmd_match:
                    interfaces[port][excel_head[cmd_match[2]]] = cmd_match[1]

                #check for 'switchport trunk allowed add' line item
                res = re.search((re.compile(trunk_allowed_add)), item)
                #append allowed VLANs onto previous cell string
                if res:
                    vlan = interfaces[port][excel_head[7]]
                    vlan = (vlan + "," + res.group(2))
                    interfaces[port][excel_head[7]] = vlan
        #add cleaned output to dictionary
        parsed_device[key] = interfaces

    #return parsed dictionary
    return parsed_device

"""
Check for minimum init args passed from cmd line
"""
def argument_check():
    in_flag = False
    out_flag = False
    files = {
        "in_files" : [],
        "out_file" : []
    }

    if len(sys.argv) > 2:
        for arg in sys.argv:
            #check for input start
            if arg == '-i':
                in_flag = True
                out_flag = False
                continue
            #set for output start
            if arg == '-o':
                out_flag = True
                in_flag = False
                continue
           #check input file validity
            if in_flag and arg != '-i':
                if os.path.exists(arg):
                    files["in_files"].append(arg)
                else:
                    print("Input file", arg, "does not exist. Skipping file...")
                    continue
            if out_flag and arg != '-o':
                files["out_file"].append(arg)
                continue
    return files

"""
Output sorted port information to .xlsx file
"""
def xlsx_output(devices, out_file_name):
    try:
        #create the new file/worksheet
        workbook = Workbook(out_file_name, {'constant_memory':True})

        #console status update
        print("Writing output to", out_file_name, "...")

        #outer loop to iterate through devices
        for device, interfaces in devices.items():
            #init vars for temp location in worksheet
            xls_col = 0
            xls_row = 0

            #new worksheet per device config file
            worksheet = workbook.add_worksheet(device)
            #print("\n\n\nDEVICE:",device)
            #create headers in current worksheet
            worksheet.write_string(xls_row, xls_col, ("Interfaces"))
            for header in excel_head:
                xls_col = xls_col + 1
                worksheet.write_string(xls_row, xls_col, (header.strip()))
            xls_row = 1

            #inner loop to iterate through interfaces and their configuration elements
            #fill columns with port data
            for port, children in interfaces.items():
                xls_col = 0
                worksheet.write_string(xls_row, xls_col, (port.strip()))
                #print("PORT:",port.strip())
                #insert string into current cell
                for header in excel_head:
                    xls_col = xls_col + 1
                    worksheet.write_string(xls_row, xls_col, children[header])

                #move to next row
                xls_row = xls_row + 1

        #close the workbook
        workbook.close()

        #console status update
        print("Process complete!")

    except:
        print("Error: Workbook updates failed.")
        print("Check if your destination excel file is open at the moment...")
        sys.exit(1)

#main function
def main():
    #verify provided arguments
    files = argument_check()
    if files:
        if len(files["in_files"]) < 0:
            print("Error: Input file/s not provided.")
            print("python portextract.py -i input -o output")
            sys.exit(1)
        if len(files["out_file"]) > 0:
            out_file_name = files["out_file"][0]
        else:
            print("Error: Output file not provided.")
            print("python portextract.py -i input -o output")
            sys.exit(1)
    else:
        print("python portextract.py -i input -o output")
        sys.exit(1)

    #buffer file to temp var
    file_buff = buffer_file(files)
    if file_buff:

        #interface extraction and parse for excel output
        db_dev["interfaces"] = interface_extract(file_buff)

        #sort interfaces for excel input
        db_dev["interfaces_xls"] = interface_parse(db_dev["interfaces"])

        #output file to .xlsx
        xlsx_output(db_dev["interfaces_xls"], out_file_name)

    else:
        #empty file buffer
        print("Empty file buffer. Please check your provided configuration files...")
        sys.exit(1)

#call main method
if __name__ == '__main__':
    main()
