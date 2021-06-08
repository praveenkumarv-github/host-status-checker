# host-status-checker
ping all you remote and get results in google sheet using gspread api



**GCP project API setup :**
https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html

**Python version : Python2.7**

Dependencies(local) : 

  * Pyping(actual worker)

  * Gspread (Python API for Google Sheets )

  * oauth2client (client library for OAuth 2.0)
 
  * urllib3==1.25.11 ()



Reference :(inspired from burnash)

https://github.com/burnash/gspread



User Scope :
############USER INPUT#############################

dir_path = r""  #path for the oauth2client key

ip_source_sheet_link = ""           #for example any" link

ip_source_sheet_link_name = ""      #for example "for example any" Name

destination_sheet_link = ""         #for example "for example any" link

destination_sheet_link_name = ""    #for example "for example any" Name

############USER INPUT End's##########################

Windows Task scheduler 

>Create a bat file 

_"C:\Python27\python.exe" "\path\to\master_ping.py"
pause_

Schedule a auto trigger using Windows Task scheduler 

Refer : https://www.windowscentral.com/how-create-and-run-batch-file-windows-10



**Google script (Automation for alignment,text conversion,catastrophe replacement)**

Schedule a auto trigger :

* arrange.gs 

* convert_into_text_from_formula.gs

* findandReplace.gs





