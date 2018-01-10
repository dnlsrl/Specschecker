#!/usr/bin/env python

'''
SPECSCHECKER 2

THIS IS A COMPLETE REWRITE OF ORIGINAL SCRIPT.
ITS AIM IS TO BE BETTER STRUCTURED THAN THE ORIGINAL
WHICH ALTHOUGH WORKED GREAT, IT WAS STUPIDLY WRITTEN, SORRY FOR THAT

IT QUERIES AND DISPLAYS TECHNICAL SPECS FOR THE PC IT'S RUNNING ON
IF ANTIVIRUS FLAGS THE SCRIPT, IT'S A FALSE POSITIVE

CONTACT FOR TROUBLESHOOTING
dnlsrl.kaiser@gmail.com
'''

import csv
import regex                    # $pip install regex
import platform
import win32com.client          # https://sourceforge.net/projects/pywin32/files/pywin32/
from collections import OrderedDict as od

def valueRe(value):
    '''CHECKES WHETHER A VALUE HAS ANY TRAILING SPACES (TEST FURTHER)'''
    # I don't know what I'm doing

    # 1. Define regex expression
    r = '^[-\w/#]*|[-\w/#]*$'

    # 2. Find the match in the value string
    for x in regex.finditer(r, value):
        if x.group(0) != '':
            return x.group(0)
    else:
        # Returns the value itself if for any reason the regex search does not work
        # Reasons being, for example, that the idiot programmer didn't test against corner cases
        return value

def get_specs():
    '''QUERIES THE VALUES AND APPENDS TO A DICTIONARY. RETURNS A DICTIONARY'''

    # 1. Dictionary of values
    values = od({
        'vendor': '',           #
        'model': '',            #
        'sn': '',               #
        'pn': '',               #
        'os': '',               #
        'os_ver': '',           #
        'os_build': '',         #
        'os_arch': '',          #
        'os_id': '',            #
        'is_licensed': False,   #
        'cpu': '',              #
        'cpu_cores': 0,         #
        'cpu_arch': 0,          #
        'ram': 0,               #
        'ram_sizes': [],        #
        'ram_vendors': [],      #
        'ram_sn': [],           #
        'ram_pn': [],           #
        'storage': 0,           #
        'hdd_sizes': [],        #
        'hdd_interfaces': [],   #
        'hdd_models': [],       #
        'hdd_vendors': [],      #
        'hdd_sn': [],           #
        'hdd_statuses': [],     #
        'net_adapters': [],     #
        'net_vendors': [],      #
        'mac_addresses': [],    #
        'connected_to': [],     #
        'has_admin': False,     #
        })

    # 2. Query the computer specification values

    strComputer = '.'
    objWMIService = win32com.client.Dispatch('WbemScripting.SWbemLocator')
    objSWbemServices = objWMIService.ConnectServer(strComputer, 'root\cimv2')

    # QUERIES FOR
    # OS NAME, VERSION, BUILD, ARCHITECTURE, PRODUCT ID
    operatingSystem = objSWbemServices.ExecQuery('SELECT Caption, Version, BuildNumber, OSArchitecture, SerialNumber FROM Win32_OperatingSystem')
    # VENDOR, MODEL, SERIAL NUMBER
    oemInfo = objSWbemServices.ExecQuery('SELECT Vendor, Name, IdentifyingNumber FROM Win32_ComputerSystemProduct')
    # PART NUMBER, TOTAL PHYSICAL MEMORY
    physicalPC = objSWbemServices.ExecQuery('SELECT SystemSKUNumber, TotalPhysicalMemory FROM Win32_ComputerSystem')
    # INDIVIDUAL SIZE, MANUFACTURER, PN, SN, LOCATION
    ramSpec = objSWbemServices.ExecQuery('SELECT Capacity, DeviceLocator, Manufacturer, PartNumber, SerialNumber FROM Win32_PhysicalMemory')
    # DISK DRIVES INTERFACE, MANUFACTURER, SIZE, MODEL, SN
    hddSpec = objSWbemServices.ExecQuery('SELECT InterfaceType, Manufacturer, Model, SerialNumber, Size, Status FROM Win32_DiskDrive')
    # CPU NAME, NUMBER OF CORES, SN, PN
    processorSpec = objSWbemServices.ExecQuery('SELECT Name, NumberOfCores, Architecture FROM Win32_Processor')
    # USERS
    winUsers = objSWbemServices.ExecQuery('SELECT Name FROM Win32_UserAccount')
    # NETWORK ADAPTERS
    networkSpec = objSWbemServices.ExecQuery('SELECT NetConnectionID, Manufacturer, MACAddress, NetEnabled FROM Win32_NetworkAdapter')
    # LICENSING INFO
    licenseInfo = objSWbemServices.ExecQuery('SELECT ApplicationID, LicenseStatus FROM SoftwareLicensingProduct')

    # LIST FOR POSSIBLE VALUES INPUTTED BY THE MANUFACTURER, WHICH CAN BE ANNOYINGLY RANDOM
    customOEM = [
        'To Be Filled By O.E.M.',
        '(Unidades de disco estándar)',
        'Fill By OEM',
        'SerNum1',
        ]

    # ASSIGN QUERY RESULTS TO VALUES DICTIONARY
    for objItem in operatingSystem:
        # OPERATING SYSTEM
        if objItem.Caption != None:
            values['os'] = objItem.Caption
        # VERSION
        if objItem.Version != None:
            count = 0
            for x in objItem.Version:
                if x == '.':
                    count += 1
                if count < 2:
                    values['os_ver'] += x
        # BUILD
        if objItem.BuildNumber != None:
            values['os_build'] = objItem.BuildNumber
        # ARCHITECTURE
        if objItem.OSArchitecture != None:
            values['os_arch'] = objItem.OSArchitecture
        # PRODUCT ID
        if objItem.SerialNumber != None:
            values['os_id'] = objItem.SerialNumber
        
    for objItem in oemInfo:
        # VENDOR
        if objItem.Vendor != None:
            values['vendor'] = objItem.Vendor
        # MODEL
        # Lenovo computers somehow don't store the complete model info in Win32_ComputerSystemProduct.Name
        if objItem.Name != None and objItem.Vendor != 'LENOVO':
            values['model'] = objItem.Name
        else:
            values['model'] = objItem.Version + ' Machine Type ' + objItem.Name
        # SERIAL NUMBER
        if objItem.IdentifyingNumber != None:
            values['sn'] = objItem.IdentifyingNumber
        
    for objItem in physicalPC:
        # PART NUMBER
        if float(values['os_ver']) >= 10:
            if objItem.SystemSKUNumber != None and objItem.SystemSKUNumber not in customOEM:
                values['pn'] = objItem.SystemSKUNumber
            else:
                values['pn'] = 'N/A'
        else:
            # Property not supported before Windows 10 and Windows Server 2016 Technical Preview
            # https://msdn.microsoft.com/en-us/library/aa394102(v=vs.85).aspx
            values['pn'] = 'N/A'
        # TOTAL PHYSICAL MEMORY
        if objItem.TotalPhysicalMemory != None:
            values['ram'] = round(int(objItem.TotalPhysicalMemory) / (1024 ** 3))

    # Dictionary of vendor codes, some of which don't directly specify a name
    ramVendors = {
        'CE80000000000000': 'Samsung',
        '4F01000000000000': 'Trascend',
        'FF04000000000000': 'Ramos',
        'AD80000000000000': 'Hynix',
        'Kinston': 'Kingston',
        }
            
    for objItem in ramSpec:
        # RAM SIZE (LIST)
        if objItem.Capacity != None:
            values['ram_sizes'].append(round(int(objItem.Capacity) / (1024 ** 3)))
        # RAM MANUFACTURER
        if objItem.Manufacturer != None and objItem.Manufacturer not in ramVendors:
            values['ram_vendors'].append(objItem.Manufacturer)
        else:
            values['ram_vendors'].append(ramVendors[objItem.Manufacturer])
        # RAM PART NUMBER
        if objItem.PartNumber != None:
            values['ram_pn'].append(valueRe(objItem.PartNumber))
        # RAM SERIAL NUMBER
        if objItem.SerialNumber != None and objItem.SerialNumber not in customOEM:
            values['ram_sn'].append(objItem.SerialNumber)
        else:
            values['ram_sn'].append('N/A')
        
    for objItem in hddSpec:
        # HDD INTERFACE
        # If the storage type is USB, the drive is skipped. This script is not interested in external drives.
        if objItem.InterfaceType != None and objItem.InterfaceType != 'USB':
            values['hdd_interfaces'].append(objItem.InterfaceType)
        else:
            continue
        # HDD MANUFACTURER
        if objItem.Manufacturer != None and objItem.Manufacturer not in customOEM:
            values['hdd_vendors'].append(objItem.Manufacturer)
        else:
            values['hdd_vendors'].append('N/A')
        # HDD MODEL
        if objItem.Model != None:
            values['hdd_models'].append(objItem.Model)
        # HDD SERIAL NUMBER
        if objItem.SerialNumber != None:
            values['hdd_sn'].append(valueRe(objItem.SerialNumber))
        # HDD SIZE
        if objItem.Size != None:
            values['hdd_sizes'].append(round(int(objItem.Size) / (1024 ** 3)))
        # HDD STATUS
        if objItem.Status != None:
            values['hdd_statuses'].append(objItem.Status)

    # STORAGE SIZE
    values['storage'] = sum(values['hdd_sizes'])

    # KNOWN ARCHITECTURES DICTIONARY
    architectures = {
        0: 'x86',
        1: 'MIPS',
        2: 'Alpha',
        3: 'PowerPC',
        6: 'ia64',
        9: 'x64',
        }
        
    for objItem in processorSpec:
        # CPU NAME
        if objItem.Name != None:
            values['cpu'] = objItem.Name
        # NUMBER OF CORES
        if objItem.NumberOfCores != None:
            values['cpu_cores'] = int(objItem.NumberOfCores)
        # CPU ARCHITECTURE
        if objItem.Architecture != None:
            values['cpu_arch'] = architectures[objItem.Architecture]
        
    for objItem in winUsers:
        # CHECKS IF COMPUTER HAS AN ADMIN USER
        # This is optional.
        # Added this because I use to add an "Admin" user to every computer I repair so as to get access to it in case the user loses their password.
        if objItem.Name != None and objItem.Name == 'Admin':
            values['has_admin'] = True
            break

    netDictionary = {
        'Conexión de área local': 'Ethernet',
        'Conexión de red inalámbrica': 'Wi-Fi',
        'Conexión de red Bluetooth': 'Bluetooth',
        }

    for objItem in networkSpec:
        # IDENTIFIES MAIN NETWORK ADAPTERS
        # Virtual adapters are skipped. For the moment, it only ignores VirtualBox.
        if objItem.NetConnectionID != None and objItem.NetConnectionID != '' and 'VirtualBox' not in objItem.NetConnectionID:
            if objItem.NetConnectionID in netDictionary:
                values['net_adapters'].append(netDictionary[objItem.NetConnectionID])
            else:
                values['net_adapters'].append(objItem.NetConnectionID)
        else:
            continue
        # NETWORK ADAPTER VENDOR
        if objItem.Manufacturer != None:
            values['net_vendors'].append(objItem.Manufacturer)
        # MAC ADDRESS
        if objItem.MACAddress != None:
            values['mac_addresses'].append(objItem.MACAddress)
        # DETECT ADAPTERS CONNECTED TO THE NETWORK
        if objItem.NetEnabled != None and objItem.NetEnabled == True:
            values['connected_to'].append(objItem.NetConnectionID)

    for objItem in licenseInfo:
        # CHECKS WHETHER THE OPERATING SYSTEM IS LICENSED
        if objItem.ApplicationID != None and objItem.ApplicationID == '55c92734-d682-4d71-983e-d6ec3f16059f':
            if objItem.LicenseStatus != None and objItem.LicenseStatus == 1:
                values['is_licensed'] = True
                break

    return values

def toScreen(values):
    '''TAKES A DICTIONARY OF VALUES AND DISPLAYS THEM TO THE SCREEN'''

    for x in values:
        if type(values[x]) == list:
            print('{}: '.format(x.upper()), ', '.join('{}'.format(str(e)) for e in values[x]))
        elif type(values[x]) == int or type(values[x]) == float or type(values[x]) == bool:
            print('{}: {}'.format(x.upper(), str(values[x])))
        else:
            print('{}: {}'.format(x.upper(), values[x]))

def toCSV(values, filename):
    '''TAKES A DICTIONARY OF VALUES AND EXPORTS TO A CSV FILE'''

    headers = [x for x in values]
    unnecessary_dictInList = [values]
    with open('{}.csv'.format(filename), 'w', newline = '') as csvfile:
        specswriter = csv.DictWriter(csvfile, delimiter = ',', fieldnames = headers)
        specswriter.writeheader()
        for x in unnecessary_dictInList:
            specswriter.writerow(x)

    print('File created and saved as {}.csv\n'.format(filename))

def nameRe(filename):
    '''USES REGEX TO CHECK WHETHER THE FILENAME HAS ANY INVALID CHARACTERS (TEST FURTHER)'''

    r = '^[-\w]+$'

    for x in regex.finditer(r, filename):
        if x.group(0) != '':
            return True
    else:
        return False

def createFile(specs):
    '''FUNCTION TO EXPORT THE DATA TO A CSV FILE'''

    filename = input('Please name your new file: ')
    if filename != '' and nameRe(filename) == True:
        toCSV(specs, filename)
    else:
        print('The file name you specified is invalid, try with another name.')
        print('Hint: special characters except for dash(-) and underscore(_) are not valid. Do not separate words with SPACE either.\n')
        createFile(specs)

def main():

    print('WELCOME TO SPECSCHECKER 2')
    print('****\n')

    print('I\'m querying the specifications for this computer: {}'.format(platform.node()))
    print('Please wait...\n')
    specs = get_specs()

    toScreen(specs)

    choice = input('\nWould you like me to create a file using this data? (Y/n) ')
    if choice.lower() == 'y' or choice == '':
        createFile(specs)
    
    print('Thanks for choosing this software.')
    exit = input('Press ENTER to exit the program...' )
    return 0

main()
