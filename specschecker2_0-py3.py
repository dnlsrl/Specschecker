#!/usr/bin/env python
# -*- coding: ISO-8859-1 -*-

'''
SPECSCHECKER 2.0

THIS IS A COMPLETE REWRITE OF ORIGINAL SCRIPT.
ITS AIM IS TO BE BETTER STRUCTURED THAN THE ORIGINAL
WHICH ALTHOUGH WORKED GREAT, IT WAS STUPIDLY WRITTEN, SORRY FOR THAT

IT QUERIES AND DISPLAYS TECHNICAL SPECS FOR THE PC IT'S RUNNING ON
IF ANTIVIRUS FLAGS THE SCRIPT, IT'S A FALSE POSITIVE

CONTACT FOR TROUBLESHOOTING
dnlsrl.kaiser@gmail.com
'''

import csv
import regex			        # $pip install regex
import win32com.client	        # https://sourceforge.net/projects/pywin32/files/pywin32/

def get_specs():
    '''QUERIES THE VALUES AND APPENDS TO A DICTIONARY. RETURNS A DICTIONARY'''

    # 1. Dictionary of values
    values = {
        'vendor': '',           #
        'model': '',            #
        'sn': '',               #
        'pn': '',               #
        'os': '',               #
        'ver': '',              #
        'build': '',            #
        'arch': '',             #
        'cpu': '',              #
        'cpu_cores': 0,         #
        'cpu_sn': '',           #
        'cpu_pn': '',           #
        'ram': 0,               #
        'ram_sizes': [],        #
        'ram_vendors': [],      #
        'ram_sn': [],           #
        'ram_pn': [],           #
        'ram_locations': [],    #
        'storage': 0,           #
        'hdd_sizes': [],        #
        'hdd_interfaces': [],   #
        'hdd_models': [],       #
        'hdd_vendors': [],      #
        'hdd_sn': [],           #
        'has_admin': None,      #
        }

    # 2. Query the computer specification values

    strComputer = '.'
    objWMIService = win32com.client.Dispatch('WbemScripting.SWbemLocator')
    objSWbemServices = objWMIService.ConnectServer(strComputer, 'root\cimv2')

    # QUERIES FOR
    # OS NAME, VERSION, BUILD, ARCHITECTURE
    colItems0 = objSWbemServices.ExecQuery('SELECT * FROM Win32_OperatingSystem')
    # VENDOR, MODEL, SERIAL NUMBER
    colItems1 = objSWbemServices.ExecQuery('SELECT * FROM Win32_ComputerSystemProduct')
    # PART NUMBER, TOTAL PHYSICAL MEMORY
    colItems2 = objSWbemServices.ExecQuery('SELECT * FROM Win32_ComputerSystem')
    # INDIVIDUAL SIZE, MANUFACTURER, PN, SN, LOCATION
    colItems3 = objSWbemServices.ExecQuery('SELECT * FROM Win32_PhysicalMemory')
    # DISK DRIVES INTERFACE, MANUFACTURER, SIZE, MODEL, SN
    colItems4 = objSWbemServices.ExecQuery('SELECT * FROM Win32_DiskDrive')
    # CPU NAME, NUMBER OF CORES, SN, PN
    colItems5 = objSWbemServices.ExecQuery('SELECT * FROM Win32_Processor')
    # USERS
    colItems6 = objSWbemServices.ExecQuery('SELECT * FROM Win32_UserAccount')
        
        
    # ASSIGN QUERY RESULTS TO VALUES DICTIONARY
    for objItem in colItems0:
        # OPERATING SYSTEM
        if objItem.Caption != None:
            values['os'] = objItem.Caption
        # VERSION
        if objItem.Version != None:
            values['ver'] = objItem.Version[0:2]
        # BUILD
        if objItem.BuildNumber != None:
            values['build'] = objItem.BuildNumber
        # ARCHITECTURE
        if objItem.OSArchitecture != None:
            values['arch'] = objItem.OSArchitecture
        
    for objItem in colItems1:
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
        
    for objItem in colItems2:
        # PART NUMBER
        if values['ver'] == '10':
            if objItem.SystemSKUNumber != None and objItem.SystemSKUNumber != 'To be filled by O.E.M.':
                values['pn'] = objItem.SystemSKUNumber
            else:
                values['pn'] = 'N/A'
        else:
            # Apparently Windows versions below 10 don't store the computer's part number in Win32_ComputerSystem.SystemSKUNumber
            # I need to look into this
            values['pn'] = 'N/A'
        # TOTAL PHYSICALMEMORY
        if objItem.TotalPhysicalMemory != None:
            values['ram'] = round(int(objItem.TotalPhysicalMemory) / (1024 ** 3))
        
    for objItem in colItems3:
        # RAM SIZE (LIST)
        if objItem.Capacity != None:
            values['ram_sizes'].append(round(int(objItem.Capacity) / (1024 ** 3)))
        # RAM PHYSICAL LOCATION
        if objItem.DeviceLocator != None:
            values['ram_locations'].append(objItem.DeviceLocator)
        # RAM MANUFACTURER
        if objItem.Manufacturer != None:
            values['ram_vendors'].append(objItem.Manufacturer)
        # RAM PART NUMBER
        if objItem.PartNumber != None:
            values['ram_pn'].append(objItem.PartNumber)
        # RAM SERIAL NUMBER
        if objItem.SerialNumber != None:
            values['ram_sn'].append(objItem.SerialNumber)
        
    for objItem in colItems4:
        # HDD INTERFACE
        if objItem.InterfaceType != None and objItem.InterfaceType != 'USB':
            values['hdd_interfaces'].append(objItem.InterfaceType)
        else:
            continue
        # HDD MANUFACTURER
        if objItem.Manufacturer != None and objItem.Manufacturer != '(Unidades de disco estándar)':
            values['hdd_vendors'].append(objItem.Manufacturer)
        else:
            values['hdd_vendors'].append('N/A')
        # HDD MODEL
        if objItem.Model != None:
            values['hdd_models'].append(objItem.Model)
        # HDD SERIAL NUMBER
        if objItem.SerialNumber != None:
            values['hdd_sn'].append(objItem.SerialNumber)
        # HDD SIZE
        if objItem.Size != None:
            values['hdd_sizes'].append(round(int(objItem.Size) / (1024 ** 3)))

    # STORAGE SIZE
    values['storage'] = sum(values['hdd_sizes'])
        
    for objItem in colItems5:
        # CPU NAME
        if objItem.Name != None:
            values['cpu'] = objItem.Name
        # NUMBER OF CORES
        if objItem.NumberOfCores != None:
            values['cpu_cores'] = int(objItem.NumberOfCores)
        # CPU PART NUMBER
        if objItem.PartNumber != None and objItem.PartNumber != 'To Be Filled By O.E.M.':
            values['cpu_pn'] = objItem.PartNumber
        else:
            values['cpu_pn'] = 'N/A'
        # CPU SERIAL NUMBER
        if objItem.SerialNumber != None and objItem.SerialNumber != 'To Be Filled By O.E.M.':
            values['cpu_sn'] = objItem.SerialNumber
        else:
            values['cpu_sn'] = 'N/A'
        
    for objItem in colItems6:
        # CHECKS IF COMPUTER HAS AN ADMIN USER
        if objItem.Name != None and objItem.Name == 'Admin':
            values['has_admin'] = True
            break

    return values

def toScreen(values):
    '''TAKES A DICTIONARY OF VALUES AND DISPLAYS THEM TO THE SCREEN'''

    for x in values:
        if type(values[x]) == list:
            print(x.upper() + ': ' + ', '.join(str(e) for e in values[x]))
        elif type(values[x]) == int or type(values[x]) == float or type(values[x]) == bool:
            print(x.upper() + ': ' + str(values[x]))
        else:
            print(x.upper() + ': ' + values[x])

def toCSV(values, filename):
    '''TAKES A DICTIONARY OF VALUES AND EXPORTS TO A CSV FILE'''

    headers = [x for x in values]
    # Apparently to print the values of a dictionary I need to place the dictionary inside a list because
    # csvwriter.writerow(row) only takes lists for parameters. It kind of make sense if you're writing several dictionaries.
    unnecessary_dictInList = [values]
    with open('%s.csv' % (filename), 'w', newline = '') as csvfile:
        # The fieldnames attribute is needed as well, sigh
        specswriter = csv.DictWriter(csvfile, delimiter = ',', fieldnames = headers)

        for x in unnecessary_dictInList:
            specswriter.writerow(x)

    print('File created and saved as %s' % (filename))

def nameRe(filename):
    '''USES REGEX TO CHECK WHETHER THE FILENAME HAS ANY INVALID CHARACTERS'''

    r = regex.match('^[\w-]+', filename)
    if len(r[0]) == len(filename):
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
        print('Hint: special characters except for dash(-) and underscore(_) are not valid. Do not separate with SPACE either.\n')
        createFile(specs)

def main():

    print('#******************* WELCOME TO SPECSCHECKER 2.0 *******************#')
    print('Thanks for choosing this software.')
    print('This is a script to automate the retrieving of a computer\'s technical specifications\n')
    
    specs = get_specs()
    print('These are the main specifications for this workstation:\n')
    toScreen(specs)

    choice = input('\nWould you like me to create a file with this data? (Y/n) ')
    if choice.lower() == 'y' or choice == '':
        createFile(specs)
    
    exit = input('Press ENTER to exit the program...' )
    return 0

main()
