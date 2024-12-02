# Created by Allyn Stott (allyn.stott@navy.mil)
# If you update with changes (please do!), email them out and I'll keep track of changes

# Updated by mixedup4x4 - I have not updated with the latest STIG Viewer or any possible changes in new .ckl files.  Simply modernized what was already here.  I also have no updated the items under dist  (.dmg / .exe)

Converts DISA's stig_viewer checklist file (ckl) to a POA&M Excel spreadsheet (xlsx). When using the DISA stig_viewer you can save your checklist where the output is a CKL file (really just an XML file). ckl2poam takes the XML output (the CKL file) and puts it into POA&M format, preserving the status and comments fields. The Excel spreadsheet (xlsx) features two tabs: POA&M and Remediation. This allows mitigators to use DISA's stig_viewer as their checklist tool as they manually walk a STIG and then put the output into something that both management (POA&M tab) and sysadmin's (Remediation tab) will be able to use. Both the "Finding Details" and "Comments" text fields are preserved in both tabs. 

To run the python tools, you will need to install openpyxl (pip install openpyxl).

A CLI and a GUI version of this tool is available. 

In the "dist" folder I have included an Apple Disk Image (.dmg) and a win32 binary (.exe) for Microsoft Windows. I created these using pyinstaller so users not familiar with Python can still use the tool.
