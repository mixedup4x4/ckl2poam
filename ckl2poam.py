#!/usr/bin/env python
#ckl2poam.py
#converts stig_viewer checklist file to a POA&M Excel spreadsheet
#Last updated by Allyn Stott (allyn.stott@navy.mil)
#If you update with changes (please do!), email them out and I'll keep track of changes

VERSION = 0.1

import sys
from xml.dom.minidom import parseString, Node
from openpyxl import Workbook

# show help
if len(sys.argv) < 3:
    print 'ckl2poam version', VERSION
    print 'Usage:', sys.argv[0], 'input.ckl output.xlsx'
    sys.exit(1)

#open ckl file from stigviewer
file = open(sys.argv[1], 'r')
data = file.read()
file.close()
dom = parseString(data)

#initialize excel workbook/worksheet for poa&m
wb = Workbook()
ws1 = wb.get_active_sheet()
ws1.title = 'POA&M'
ws2 = wb.create_sheet()
ws2.title = 'Remediation'
c1Values = ['Weakness','CAT','Severity','IA Control and Impact Code','POC','Resources Required','Scheduled Completion Date','Milestones with Completion Dates','Milestone Changes','Source Identifying Weakness','Status','Finding Details','Comments']
c2Values = ['Weakness','CAT','Severity','IA Control and Impact Code','Source Identifying Weakness','Vulnerability Discussion','Check Content','Fix Text','Finding Details','Comments']
for i in range(len(c1Values)):
    ws1.cell(row = 0, column = i).value = c1Values[i]
for i in range(len(c2Values)):
    ws2.cell(row = 0, column = i).value = c2Values[i]

#fill excel workbook with values from ckl file
rowCounter = 1
vulnTag = dom.getElementsByTagName('VULN')
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
            groupTitle = ''

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

        if vuln.getElementsByTagName('FINDING_DETAILS')[0].hasChildNodes():
            findDetails = vuln.getElementsByTagName('FINDING_DETAILS')[0].firstChild.data
        else:
            findDetails = ''

        if vuln.getElementsByTagName('COMMENTS')[0].hasChildNodes():
            comments = vuln.getElementsByTagName('COMMENTS')[0].firstChild.data
        else:
            comments = ''
            
        poamValues = [ruleTitle,cat,severity,iaControls,'','','','','',groupTitle + '\n' + ruleid,status,findDetails,comments]
        remeValues = [ruleTitle,cat,severity,iaControls,groupTitle + '\n' + ruleid,vulnDiscuss,chkContent,fixText,findDetails,comments]
        for i in range(len(poamValues)):
            ws1.cell(row = rowCounter, column = i).value = poamValues[i]
        for i in range(len(remeValues)):
            ws2.cell(row = rowCounter, column = i).value = remeValues[i]
        rowCounter = rowCounter + 1

#write all values to xlsx file
wb.save(sys.argv[2])
