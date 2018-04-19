#!/usr/bin/env python

import regex
import platform
import collections
import win32com.client


class computer:

    release = float(platform.win32_ver()[0])
    values = collections.OrderedDict()

    def __init__(self):
        pass

    def connection(self):
        strComputer = '.'
        objWMIService = win32com.client.Dispatch('WbemScripting.SWbemLocator')
        objSWbemServices = objWMIService.ConnectServer(strComputer, 'root\cimv2')
        return objSWbemServices

    def operating_system(self, connection):
        query = connection.ExecQuery('SELECT Caption, Version, BuildNumber, OSArchitecture, SerialNumber FROM Win32_OperatingSystem')
        for item in query:
            if item.Caption is not None:
                self.values['os'] = item.Caption
            if item.Version is not None:
                count = 0
                for x in item.Version:
                    if x == '.':
                        count += 1
                    if count < 2:
                        self.values['os_ver'] += x
            if item.BuildNumber is not None:
                self.values['os_build'] = item.BuildNumber
            if item.OSArchitecture is not None:
                self.values['os_arch'] = item.OSArchitecture
            if item.SerialNumber is not None:
                self.values['os_id'] = item.SerialNumber

    def oem_info(self, connection):
        query = connection.ExecQuery('SELECT Vendor, Name, IdentifyingNumber, Version FROM Win32_ComputerSystemProduct')
        for item in query:
            if item.Vendor is not None:
                self.values['vendor'] = item.Vendor
            if item.Name is not None and item.Vendor is not 'LENOVO':
                self.values['model'] = item.Name
            else:
                self.values['model'] = item.Version + ' Machine Type ' + item.Name
            if item.IdentifyingNumber is not None:
                self.values['sn'] = item.IdentifyingNumber

    def physical_pc(self, connection):
        messages = [
            'To Be Filled By O.E.M.',
            'Fill By OEM',
        ]
        query = connection.ExecQuery('SELECT SystemSKUNumber, TotalPhysicalMemory FROM Win32_ComputerSystem')
        for item in query:
            if self.release >= 10:
                if item.SystemSKUNumber is not None and item.SystemSKUNumber not in messages:
                    self.values['pn'] = item.SystemSKUNumber
                else:
                    self.values['pn'] = 'N/A'
