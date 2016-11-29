# -*- coding: utf-8 -*-

'''
Created on 2016. m�j. 20.

Introduction:

This Function Set Designed for editing values in Excel Table for BMW_ACSM5 project Validation and Variation file editing.
However this application is really task specific it contains many function that can be used for general purpose.

Usage:

Edit the signalargs dictionary (seen below) for to choose which value will be changed.
Edit the fileargs and set the PATH of the editable files.
Edit the cells with using: Excelapp.setsignals() function


@author: SzuroveczD
'''

import win32com.client as win32
from datetime import datetime

# Editable Signal Arguments

excel = win32.gencache.EnsureDispatch('Excel.Application')

signalargs = {
    0: 'QUAL_DEQUAL.QU_ACLNX_RAW.DiffTime',  # CV
    1: 'QUAL_DEQUAL.QU_ACLNY_RAW.DiffTime',  # DB
    2: 'QUAL_DEQUAL.QU_ACLNZ_RAW.DiffTime',  # EL
    3: 'QUAL_DEQUAL.QU_ACLNY_RAW_2.DiffTime'  # DN
}


class ExcelApp:

    def __init__(self, filename, worksheet):
        '''Init filename and worksheet'''
        self.filename = filename
        self.worksheet = worksheet
        self.logdata = ''

    def rownum(self):
        '''Return the number of rows in the sheet'''
        xldata = self.worksheet.UsedRange.Value
        return len(xldata)

    def addvalue(self, range, data):
        '''Set data on a specific range
        :param int range: the range of the data in excel file.
        :param str data: the data what you insert.
        '''
        self.worksheet.Range(range).Value = data

    def getdata(self, rangeval):
        '''Grab the data from a specific range
        :param str rangeval: The range where the data. Ex.: A15:B20
        :return: (list) What contains the cell values get from the rangeval.
        '''
        data = excel.Range(rangeval)
        return data.Value

    def setlogdata(self, elements, i, rangeval, counter):
        '''
        If something inconsistent this function returns the details where is the problem
        :param elements: The elements on the cell
        :param rangeval : The value of the range
        :param pos: The position where the data can be found in the cell.
        '''
        self.logdata += ' \n ' + str(len(elements[0])) + ' But found ' + str(
            len(i)) + ' element in the range ' + str(rangeval[:2]) + str(counter + 3 - 1) + ' ' + str(i)

    def getlogdata(self):
        return self.logdata

    def consistency_checker(self, data, signalargs, valoneortwo, *pos):

        self.set_consistency(True)
        for position in pos:
            for index in signalargs:
                self.colval = self.getsginalrange(signalargs[index], valoneortwo)
                newlist, counter, newelement = [], 0, str(data)
                elements = [str(element[0]).split(",") for element in self.getdata(self.colval)]
                for i in elements:
                    # Go Through in each cell on the column
                    counter += 1
                    if len(i) != len(elements[0]) and len(i) != 1:
                        # If a cell contains more elements than expected set the new position of
                        # the data
                        self.set_consistency(False)
                        self.setnotexpectedvals(len(i))
                        self.setlogdata(elements, i, self.colval, counter)

    def setnotexpectedvals(self, nexpectedvals):
        '''This method has been called if number of measurement values are not the same as expected'''
        self.nexpectedvals = nexpectedvals

    def setnewpos(self, newpos):
        self.newpos = newpos

    def getnewpos(self):
        return self.newpos

    def getnotexpectvals(self):
        return self.nexpectedvals

    def setdatas(self, rangeval, newdata, pos):
        '''
        Set the data on a specific range with the requested position
        :param str rangeval: The range where the data could be found. Ex.: A15:B20
        :param str newdata : The data what you want to insert.
        :param str pos : The position where the data can be found in the cell.
        '''

        newlist, counter, newelement = [], 0, str(newdata)
        elements = [str(element[0]).split(",") for element in self.getdata(rangeval)]
        for i in elements:
            # Go Through in each cell on the column
            counter += 1
            if self.get_consistency() == False and len(i) == self.getnotexpectvals():
                # Here we check if we find not expected values
                print self.setlogdata(elements, i, rangeval, counter)
                newpos = int(self.getnewpos())
                i[newpos] = newelement
            # If not expected values
            if len(i) == len(elements[0]):
                i[pos] = newelement
            # If the current cell contains the same number of elements as the first element.
            # Replace the corresponded data with the new one
            elif len(i) == 1:
                newlist.append(",".join(i))
                continue
            # If the cell contains only one element don't touch it.
            elif len(i) != len(elements[0]):
                self.setlogdata(elements, i, rangeval, counter)
            newlist.append(",".join(i))
        self.worksheet.Range(rangeval).Value = zip([i for i in newlist])

    def setsignals(self, data, signalargs, valoneortwo, *pos):
        '''Change Signals Arguments in the beginning of the file'''
        for position in pos:
            for index in signalargs:
                self.colval = self.getsginalrange(signalargs[index], valoneortwo)
                self.setdatas(self.colval, data, position)

    def get_consistency(self):
        return self.verif

    def set_consistency(self, verif):
        self.verif = verif

    def getnumofvalsrow(self, rangeval):
        '''Returns the number of rows'''
        data = [element[0].split(",") for element in self.getdata(rangeval)]

    def getcolumnames(self):
        '''Returns a list with the name of the column'''
        signallist = []
        data = [element for element in self.getdata("A1:HY1")]
        for i in data:
            for j in range(len(i)):
                signallist.append(str(i[j]))
        return signallist

    def getsignalcolval(self, signalname, pos):
        '''Returns the column number of the singal in excel table
           pos = 0,    Returns the column number of the signal
           pos = 1,    Returns the column number of the val1 of the signal
           pos = 2,    Returns the column number of the val2 of the signal'''
        counter = 0
        for i in self.getcolumnames():
            counter += 1
            if i == signalname:
                return counter + pos

    def getsignalcolletter(self, signalname, pos):
        '''Returns the name of the column where the signal exist if
           pos = 0,    Returns the column name of the signal
           pos = 1,    Returns the column name of the val1 of the signal
           pos = 2,    Returns the column name of the val2 of the signal'''
        column_name = self.getsignalcolval(signalname, pos)
        first_val = column_name // 26
        sec_val = column_name % 26
        if first_val >= 1:
            first_val = chr(first_val + 64)
        sec_val = chr(sec_val + 64)
        return str(first_val) + str(sec_val)

    def getsginalrange(self, signalname, pos):

        return str(self.getsignalcolletter(str(signalname), pos)) + str(3) + ":" + \
            str(self.getsignalcolletter(str(signalname), pos)) + str(self.rownum())
