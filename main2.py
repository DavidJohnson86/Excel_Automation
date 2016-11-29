#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
==============================================================================
        GUI for ISC Variation Validation File Editing via win32 API
==============================================================================
                            OBJECT SPECIFICATION
==============================================================================
$ProjectName: $
$Source: main.py
$Revision: 1.1 $
$Author: David Szurovecz $
$Date: 2016/08/03 16:41:32CEST $
$Name:  $
Improvements :

Needed to fix : Some variables has been called and set both Class
                Needed more set get methods
                Needed to optimize consistency checker in the main file


(Bug)fixes:


History:

============================================================================
"""
import Tkinter as tk
import threading
import tkMessageBox
import ISC_Validation_0_8
import win32com.client as win32
import datetime
import ntpath
ntpath.basename("a/b/c")
from easygui import fileopenbox, diropenbox

signalargs = {
    0: 'QUAL_DEQUAL.QU_ACLNX_RAW.DiffTime',  # CV
    1: 'QUAL_DEQUAL.QU_ACLNY_RAW.DiffTime',  # DB
    2: 'QUAL_DEQUAL.QU_ACLNZ_RAW.DiffTime',  # EL
    3: 'QUAL_DEQUAL.QU_ACLNY_RAW_2.DiffTime'  # DN
}


class WarningGUI(tk.Frame):

    def __init__(self, parent, Appobj, *args, **kwargs):
        '''Init'''
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = tk.Frame(parent)
        self.parent.pack()
        self.Appobj = Appobj
        self.initUI()

    def initUI(self):
        "Draw elements"

        self.userInteractWindow = tk.Text(self.parent)
        self.userInteractWindow.insert(tk.END, "Warning\n" +
                                       '\n' +
                                       str(self.Appobj.filename) +
                                       str(self.Appobj.getlogdata()))
        self.userInteractWindow.pack()
        self.userInteractMessage = tk.Label(
            self.parent,
            text="Please verify your position of data: ")
        self.userInteractMessage.pack()
        self.userInteractEntry = tk.Entry(self.parent, text="OK")
        self.userInteractEntry.pack()
        self.userInteractButton = tk.Button(self.parent, text="OK", command=self.ok)
        self.userInteractButton.pack()

    def ok(self):

        self.Appobj.setnewpos(self.userInteractEntry.get())
        self.quit()


class MainApplication(tk.Frame):

    def __init__(self, parent, *args, **kwargs):
        '''Init'''
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = tk.Frame(parent)
        self.parent.pack()
        root.title('ISC TOOL')
        self.file_list = u''
        self.save_list = u''
        self.running = True
        self.initUI()

    def initUI(self):
        '''Draw the UI elements'''

        # Enter File Details Frame
        stepOne = tk.LabelFrame(self.parent, text=" 1. Enter File Details: ")
        stepOne.grid(row=0, columnspan=7, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        # Enter Data Details Frame
        stepTwo = tk.LabelFrame(self.parent, text=" 2. Enter Data Details: ")
        stepTwo.grid(row=2, columnspan=7, sticky='W', padx=5, pady=5, ipadx=5, ipady=5)

        # OKAY Button

        stepThree = tk.Button(self.parent, text="OK", command=self.okbutton)
        stepThree.grid(row=3, column=0, sticky='W' + 'E', padx=5, pady=5, ipadx=20, ipady=5)

        # Cancel Button

        stepFour = tk.Button(self.parent, text="Cancel", command=self.cancelbutton)
        stepFour.grid(row=3, column=1, sticky='W', padx=5, pady=5, ipadx=23, ipady=5)

        # File Selection Text

        inFileLbl = tk.Label(stepOne, text="Select the File:")
        inFileLbl.grid(row=0, column=0, sticky='E', padx=5, pady=2)

        # File Selection Entry

        self.inFileTxt = tk.Entry(stepOne)
        self.inFileTxt.grid(row=0, column=1, columnspan=7, sticky="W", pady=3)

        # File Browse Button

        inFileBtn = tk.Button(stepOne, text="Browse ...", command=self.browsebutton)
        inFileBtn.grid(row=0, column=8, sticky='W', padx=5, pady=2)

        # File Save Label

        outFileLbl = tk.Label(stepOne, text="Save File to:")
        outFileLbl.grid(row=1, column=0, sticky='E', padx=5, pady=2)

        # File Save Entry

        self.outFileTxt = tk.Entry(stepOne)
        self.outFileTxt.grid(row=1, column=1, columnspan=7, sticky="WE", pady=2)

        # File Save Button

        outFileBtn = tk.Button(stepOne, text="Browse ...", command=self.savebutton)
        outFileBtn.grid(row=1, column=8, sticky='W', padx=5, pady=2)

        # Value Label

        valTblLbl = tk.Label(stepTwo, text="Enter the value what you want to apply:")
        valTblLbl.grid(row=3, column=0, sticky='W', padx=5, pady=2)

        # Value Entry

        self.val = tk.IntVar()

        self.valTblEnt = tk.Entry(stepTwo, textvariable=self.val)
        self.valTblEnt.grid(row=3, column=1, columnspan=3, pady=2, sticky='WE')

        # Position value

        posLbl = tk.Label(stepTwo, text="Enter the position of the new value:")
        posLbl.grid(row=4, column=0, padx=5, pady=2, sticky='W')

        # Position Entry

        self.p1 = tk.IntVar()
        self.posEnt = tk.Entry(stepTwo, textvariable=self.p1)
        self.posEnt.grid(row=4, column=1, columnspan=3, pady=2, sticky='WE')

        # Val1 & Val2 Label

        fldLbl = tk.Label(stepTwo, text="Choose which value want to modify:")
        fldLbl.grid(row=6, column=0, padx=5, pady=2, sticky='W')

        # Val1 & Val2 Check button

        self.v1 = tk.BooleanVar()
        self.v2 = tk.BooleanVar()
        self.val1Chk = tk.Checkbutton(stepTwo, text="Val1", variable=self.v1)
        self.val1Chk.grid(row=7, column=0, sticky='W', padx=5, pady=2)
        self.val2Chk = tk.Checkbutton(stepTwo, text="Val2", variable=self.v2)
        self.val2Chk.grid(row=8, column=0, sticky='W', padx=5, pady=2)

    def browsebutton(self):
        '''Select the files to Edit'''
        self.file_list = fileopenbox(
            default=r'd:\\08 ----------------Support Jobs----------------\\02_Variation_Files\\Delta_TEST_2016_05_19\\2016.07.18_004.018.003\\',
            multiple=True)
        if self.file_list:
            for item in self.file_list:
                self.inFileTxt.insert(0, str(item.encode('utf-8')))

    def savebutton(self):
        '''Select the files save path'''
        self.save_list = diropenbox(default=r'd:\\')
        if self.save_list:
            self.outFileTxt.insert(0, self.save_list.encode('utf-8'))

    def okbutton(self):
        '''Input Verification and start the process'''
        if self.v1.get() == True:
            self.colsel = 1
        if self.v2.get() == True:
            self.colsel = 2
        if self.v1.get() == False and self.v2.get() == False:
            self.errormessage()
        if self.file_list and self.save_list and self.val and self.v1:
            self.process()

    def cancelbutton(self):
        root.destroy()

    def errormessage(self):
        '''Show an Error window'''
        tkMessageBox.showinfo("Error", "Missing Data")

    def path_leaf(self, path):
        '''Extract file name from path '''
        head, tail = ntpath.split(path)
        return tail or ntpath.basename(head)

    def intInitGUI(self):
        "Draw elements"
        self.top = tk.Toplevel(self)
        self.userInteractWindow = tk.Label(self.top, text="Warning\n" + str(self.message))
        self.userInteractWindow.pack()
        self.userInteractEntry = tk.Entry(self.top, text="OK")
        self.userInteractEntry.pack()
        self.userInteractButton = tk.Button(self.top, text="OK", command=self.sendnewpos)
        self.userInteractButton.pack()

    def sendnewpos(self):
        self.Accepted = True

    def process(self):

        for files in range(0, len(self.file_list)):
            # Modify Values in the file
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            self.file_name = self.file_list[files]
            wb = excel.Workbooks.Open(self.file_name)
            worksheet = [i for i in wb.Worksheets][0]
            with open(''.join(str(datetime.datetime.now()).split(' ')[0] + '.txt'), 'a') as f:
                f.write('\n' + self.file_name.encode('utf-8'))
            Appobj = ISC_Validation_0_8.ExcelApp(self.file_name, worksheet)
            self.Appobj = Appobj
            self.Appobj.consistency_checker(self.val.get(), signalargs, self.colsel, self.p1.get())

            ''''A flag indicates that the the all csv elements are the same
                amount as expected in a columns .
            '''
            if self.Appobj.get_consistency() == True:
                # print 'Succesful: ' + str(self.file_name)
                pass

            else:
                # If not expected csv elemenets
                self.message = self.Appobj.getlogdata()
                # print '\nWarning: ' + str(self.file_name)
                # print self.message
                # Ask user to enter new position
                IntGUI = tk.Tk()
                WarningGUI(IntGUI, Appobj)
                IntGUI.mainloop()
                IntGUI.destroy()

                '''self.newpos = raw_input('Please verify the position of the corresponding value: ')
                self.Appobj.setnewpos(self.newpos)'''
            # Edit Excel file
            self.Appobj.setsignals(self.val.get(), signalargs, self.colsel, self.p1.get())
            # Save file to specific folder
            wb.SaveAs(self.save_list + '\\' + self.path_leaf(self.file_name) + '.xls')
            excel.Visible = True
            # excel.Application.Quit()
        root.quit()


if __name__ == "__main__":

    root = tk.Tk()
    run = MainApplication(root)
    root.mainloop()
