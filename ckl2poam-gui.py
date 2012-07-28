#!/usr/bin/env python
#ckl2poam.py
#converts stig_viewer checklist file to a POA&M Excel spreadsheet
#Allyn Stott (allyn.stott@navy.mil)
#If you update with changes (please do!), email them out and I'll keep track of changes

VERSION = 0.1

import sys
import subprocess
import Tkinter
import tkFileDialog
import FileDialog
from xml.dom.minidom import parseString, Node
from openpyxl import Workbook

class ckl2poamGUI:

    def __init__(self):
        
        app = Tkinter.Tk()
        app.title('ckl2poam GUI')
        fr1 = Tkinter.Frame(app, width = 300, height = 100)
        fr1.pack(side="top", pady=10)
        fr2 = Tkinter.Frame(app, width = 300, height = 100)
        fr2.pack(side="top", pady=10)
        fr3 = Tkinter.Frame(app, width = 300, height = 100)
        fr3.pack(side="bottom", pady=10)
        
        getFileButton = Tkinter.Button(fr1, text='Open CKL file ...')
        getFileButton.bind('<Button>', self.GetFile)
        getFileButton.pack(side='left')
        self.filein = Tkinter.Entry(fr1)
        self.filein.pack(side='right')
        
        saveFileButton = Tkinter.Button(fr2, text='Save XLSX file ...')
        saveFileButton.bind('<Button>', self.SaveFile)
        saveFileButton.pack(side='left')
        self.fileout = Tkinter.Entry(fr2)
        self.fileout.pack(side='right')

        okaybutton = Tkinter.Button(fr3, text=' OK ')
        okaybutton.bind("<Button>", self.CreateXLSX)
        okaybutton.pack(side="left")

        cancelbutton = Tkinter.Button(fr3, text='Cancel')
        cancelbutton.bind("<Button>", self.KillApp)
        cancelbutton.pack(side="right")
        
        ws = app.winfo_screenwidth()
        hs = app.winfo_screenheight()
        x = (ws/2) - (400/2)
        y = (hs/2) - (250/2)
        app.geometry('%dx%d+%d+%d' % (400, 160, x, y))

        app.mainloop()
        
    def KillApp(self, event):
        sys.exit(1)
    
    #initialize and create excel workbook/worksheet    
    def CreateXLSX(self, event):
        wb = Workbook()
        ws1 = wb.get_active_sheet()
        ws1.title = 'POA&M'
        ws2 = wb.create_sheet()
        ws2.title = 'Remediation'
        c1Values = ['Weakness','CAT','Severity','IA Control and Impact Code','POC','Resources Required','Scheduled Completion Date','Milestones with Completion Dates','Milestone Changes','Source Identifying Weakness','Status','Comments']
        c2Values = ['Weakness','CAT','Severity','IA Control and Impact Code','Source Identifying Weakness','Vulnerability Discussion','Check Content','Fix Text','Comments']
        for i in range(len(c1Values)):
            ws1.cell(row = 0, column = i).value = c1Values[i]
        for i in range(len(c2Values)):
            ws2.cell(row = 0, column = i).value = c2Values[i]
                    
        #fill excel workbook with values from ckl file
        rowCounter = 1
        vulnTag = self.dom.getElementsByTagName('VULN')
        for vuln in vulnTag:
            status = vuln.getElementsByTagName('STATUS')[0].firstChild.data
            if status == 'Not_Reviewed' or status == 'Open':
                severity = vuln.getElementsByTagName('ATTRIBUTE_DATA')[1].firstChild.data
                if severity == 'high':
                    cat = 'CAT I'
                elif severity == 'medium': 
                    cat = 'CAT II'
                elif severity == 'low':
                    cat = 'CAT III'
                else:
                    cat = ''

                if vuln.getElementsByTagName('ATTRIBUTE_DATA')[2].hasChildNodes():
                    groupTitle = vuln.getElementsByTagName('ATTRIBUTE_DATA')[2].firstChild.data
                else:
                    comments = ''

                if vuln.getElementsByTagName('ATTRIBUTE_DATA')[3].hasChildNodes():
                    ruleid = vuln.getElementsByTagName('ATTRIBUTE_DATA')[3].firstChild.data
                else:
                    ruleid = ''
                if vuln.getElementsByTagName('ATTRIBUTE_DATA')[5].hasChildNodes():
                    ruleTitle = vuln.getElementsByTagName('ATTRIBUTE_DATA')[5].firstChild.data
                else:
                    ruleTitle = ''
                if vuln.getElementsByTagName('ATTRIBUTE_DATA')[6].hasChildNodes():
                    vulnDiscuss = vuln.getElementsByTagName('ATTRIBUTE_DATA')[6].firstChild.data
                else:
                    vulnDiscuss = ''
                if vuln.getElementsByTagName('ATTRIBUTE_DATA')[7].hasChildNodes():
                    iaControls = vuln.getElementsByTagName('ATTRIBUTE_DATA')[7].firstChild.data
                else:
                    iaControls = ''
                if vuln.getElementsByTagName('ATTRIBUTE_DATA')[8].hasChildNodes():
                    chkContent = vuln.getElementsByTagName('ATTRIBUTE_DATA')[8].firstChild.data
                else:
                    chkContent = ''
                if vuln.getElementsByTagName('ATTRIBUTE_DATA')[9].hasChildNodes():
                    fixText = vuln.getElementsByTagName('ATTRIBUTE_DATA')[9].firstChild.data
                else:
                    fixText = ''
                if vuln.getElementsByTagName('COMMENTS')[0].hasChildNodes():
                    comments = vuln.getElementsByTagName('COMMENTS')[0].firstChild.data
                else:
                    comments = ''
                    
                poamValues = [ruleTitle,cat,severity,iaControls,'','','','','',groupTitle + '\n' + ruleid,status,' ']
                remeValues = [ruleTitle,cat,severity,iaControls,groupTitle + '\n' + ruleid,vulnDiscuss,chkContent,fixText,comments]
                for i in range(len(poamValues)):
                    ws1.cell(row = rowCounter, column = i).value = poamValues[i]
                for i in range(len(remeValues)):
                    ws2.cell(row = rowCounter, column = i).value = remeValues[i]
                rowCounter = rowCounter + 1
        wb.save(self.fout)
        sys.exit(0)
    
    #open ckl file and parse with xml.dom.minidom.parseString()    
    def GetFile(self, event):
        self.fin = tkFileDialog.askopenfilename()
        self.filein.insert(0, self.fin)
        file = open(self.fin, 'r')
        data = file.read()
        file.close()
        self.dom = parseString(data)
        
    #write all values to xlsx file
    def SaveFile(self, event):
        self.fout = tkFileDialog.asksaveasfilename(defaultextension=".xlsx")
        self.fileout.insert(0, self.fout)
            
if __name__ == '__main__':
    try:
        ckl2poamGUI()
    except KeyboardInterrupt:
        raise SystemExit('Aborted by user request.')