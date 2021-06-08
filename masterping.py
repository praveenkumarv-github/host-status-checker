import pyping
import platform 
from time import sleep
import time
import random
import os
import datetime
import socket
import re

import gspread
from oauth2client.service_account import ServiceAccountCredentials


############USER INPUT#############################
dir_path = r" "  #path for the oauth2client key
ip_source_sheet_link = " "           #for example WFH-System-details" link
ip_source_sheet_link_name = " "      #for example "WFH-System-details" Name
destination_sheet_link = " "         #for example "power_on_status_daily_report_generated" link
destination_sheet_link_name = ""    #for example "power_on_status_daily_report_generated" Name
############USER INPUT End's##########################



start_time = time.time()

#########################gsheet###########################################################
      #SCOPE CREATED TO GET ALL VALID IP FROM WFH SHEET#
# decalring the regex pattern for IP addresses
pattern = re.compile("^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$")

pattern2 = re.compile("^([^.]+)")


# initializing the list objects
ip_db =[]
##############################################

validIP_count = 0
validIP = []

#######gsheet#####
# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
#scope = ['https://spreadsheets.google.com/feeds']
os.chdir(dir_path)
creds = ServiceAccountCredentials.from_json_keyfile_name('project01-8ef1ff3bd3dd.json', scope)
client = gspread.authorize(creds)

wfh_sheet = ip_source_sheet_link
sheet = client.open_by_url(wfh_sheet)
worksheet = sheet.worksheet(ip_source_sheet_link_name)

ip_list = worksheet.col_values(11)
for ip in ip_list :
    ip_db.append(ip)
    # print(ip)
# print(ip_db)

for i in ip_db:
    result = pattern.search(i)
    if result :
        validIP_count = validIP_count + 1
        validIP.append(result.group(0))
        print(result.group(0))
# print(validIP)

#########################gsheet###############################
# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
#scope = ['https://spreadsheets.google.com/feeds']
os.chdir(dir_path)
creds = ServiceAccountCredentials.from_json_keyfile_name('project01-8ef1ff3bd3dd.json', scope)
client = gspread.authorize(creds)

power_on_status_daily_report_generated = destination_sheet_link
sheet = client.open_by_url(power_on_status_daily_report_generated)
#########################gsheet###############################
count = 0 
datetime = datetime.datetime.now()
datetime = time.strftime("%d/%m-%H:%M")
worksheetname =  ("{}".format(datetime))
worksheet = sheet.add_worksheet(title=worksheetname, rows="500", cols="13")


##creating coloumn name 
ip_gdata = worksheet.update("A1"," IP Address")
time.sleep(2)
username_gdata = worksheet.update("B1","Online UserName")
time.sleep(2)
department_gdata = worksheet.update("C1","Department")
time.sleep(2)
location_gdata = worksheet.update("D1"," Location")
time.sleep(2)
botAssetId_gdata = worksheet.update("E1","Assest ID")
time.sleep(2)

worksheet.format("A2:A400", {
    "backgroundColor": {
      "green": 50
    }
})

ip_gdata = worksheet.update("G1"," IP Address")
time.sleep(2)
username_gdata = worksheet.update("H1","Offline UserName")
time.sleep(2)
department_gdata = worksheet.update("I1","Department")
time.sleep(2)
location_gdata = worksheet.update("J1"," Location")
time.sleep(2)
botAssetId_gdata = worksheet.update("K1"," Assest ID")
time.sleep(2)

worksheet.format("G2:G400", {
    "backgroundColor": {
      "red": 50
    }
})


count = 0
activeIP = 0
inactiveIP = 0
activeIP_count = 0
inactiveIP_count = 0

cellNumberCount  = 1  #for cell starting point
cellNumberCountB = 1



for host in validIP :
    count = count + 1
    r = pyping.ping(host)
    if r.ret_code == 0:
        activeIP = activeIP + 1
        try:
          addr1 = socket.gethostbyaddr(host)
        except socket.herror, e:
          addr1 = ["hostname not found"," "]
        data = (("{} - {} - ACTIVE \n").format(host,addr1[0]))
        #print(host)
        print(("{} -> {} - Active").format(host,addr1[0]))

        # with open("activelog.txt",'a') as fopen :
        #     fopen.write(data)

        ###g_data  
        #########This Block is inject Active IP's in A coulmn#######
        cellNumberCount = cellNumberCount + 1    
        gdata1 = (("{}").format(host))
        cellName1 = ("A{}".format(cellNumberCount))
        ip_gsheet1 = worksheet.update(cellName1, gdata1)   #----------------------->  ###g_data###

      
      

        ####OnlineFormula
        var_vlookup = "="
        online_username_formula = ('''=VLOOKUP(A{},IMPORTRANGE("https://sheet-link_that need to be VlookupedG"),6,FALSE)''').format(cellNumberCount)
        online_username_cellname = ("B{}").format(cellNumberCount)

        online_department_formula = ('''=VLOOKUP(A{},IMPORTRANGE("https://sheet-link_that need to be Vlookuped"),4,FALSE)''').format(cellNumberCount)
        online_department_cellname = ("C{}").format(cellNumberCount)

        online_location_formula = ('''=VLOOKUP(A{},IMPORTRANGE("https://sheet-link_that need to be VlookupedD"),2,FALSE)''').format(cellNumberCount)
        online_location_cellname = ("D{}").format(cellNumberCount)

        online_botassid_formula = ('''=VLOOKUP(A{},IMPORTRANGE("https://sheet-link_that need to be VlookupedD"),3,FALSE)''').format(cellNumberCount)
        online_botassid_cellname = ("E{}").format(cellNumberCount)

        online_username_gdata = worksheet.update(online_username_cellname, online_username_formula)
        time.sleep(2)
        online_department_gdata = worksheet.update(online_department_cellname, online_department_formula)
        time.sleep(2)
        online_location_gdata = worksheet.update(online_location_cellname, online_location_formula)
        time.sleep(2)
        online_botassid_gdata = worksheet.update(online_botassid_cellname, online_botassid_formula)
        time.sleep(2)

    else:
        inactiveIP = inactiveIP + 1
        data = (("{}\n").format(host))
        print("Failed with {} / {}  Inactive".format(r.ret_code,host))
        
        ###g_data
         #########This Block is inject InActive IP's in G coulmn#######
        gdata = (("{}").format(host))
        inactiveIP_count = inactiveIP_count +1
        cellNumberCountB = cellNumberCountB + 1
        cellName = ("G{}".format(cellNumberCountB))
        ip_gsheet = worksheet.update(cellName, gdata)   #----------------------->  ###g_data###
      
        # with open("inactivelog.txt",'a') as fopen :
        #     fopen.write(data)

        ####ofllineFormula
        oflline_username_formula = ('''=VLOOKUP(G{},IMPORTRANGE("https://sheet-link_that need to be VlookupedG"),6,FALSE)''').format(cellNumberCountB)
        oflline_username_cellname = ("H{}").format(cellNumberCountB)

        oflline_department_formula = ('''=VLOOKUP(G{},IMPORTRANGE("https://sheet-link_that need to be Vlookuped","resource!B:E"),4,FALSE)''').format(cellNumberCountB)
        oflline_department_cellname = ("I{}").format(cellNumberCountB)

        oflline_location_formula = ('''=VLOOKUP(G{},IMPORTRANGE("https://sheet-link_that need to be VlookupedD"),2,FALSE)''').format(cellNumberCountB)
        oflline_location_cellname = ("J{}").format(cellNumberCountB)

        oflline_botassid_formula = ('''=VLOOKUP(G{},IMPORTRANGE("https://sheet-link_that need to be VlookupedD"),3,FALSE)''').format(cellNumberCountB)
        oflline_botassid_cellname = ("K{}").format(cellNumberCountB)

        oflline_username_gdata = worksheet.update(oflline_username_cellname, oflline_username_formula)
        time.sleep(2)
        oflline_department_gdata = worksheet.update(oflline_department_cellname, oflline_department_formula)
        time.sleep(2)
        oflline_location_gdata = worksheet.update(oflline_location_cellname, oflline_location_formula)
        time.sleep(2)
        oflline_botassid_gdata = worksheet.update(oflline_botassid_cellname, oflline_botassid_formula)
        time.sleep(2)

total_time = time.time() - start_time
time_in_minites = total_time / 60
time_in_minites = str(time_in_minites)
time_in_minites = time_in_minites[0:4]
print("total IP {}".format(count))
print("total activeIP  {}".format(activeIP))
print("total inactiveIP {}".format(inactiveIP))

total_timeGdata = worksheet.update("F10", "Total-time-Taken")
total_timeGdata2 = worksheet.update("F11",time_in_minites)

activeIP_gdata = (("   Online = {}").format(activeIP))
activeIP = worksheet.update("F12",activeIP_gdata )
time.sleep(2)
inactiveIP_gdata = (("   Offline = {}").format(inactiveIP))
activeIP = worksheet.update("F13",inactiveIP_gdata )

print(("total time taken  {}").format(time_in_minites))






##########hostname injection###########################
sheet_Hostname_injection = ip_source_sheet_link
sheet = client.open_by_url(sheet_Hostname_injection)
worksheet = sheet.worksheet("sheetName")


cellNum = 1 ###initiator for cell D1 to injest hostname & E1 to ipaddress


for ip in validIP :
    cellNum = cellNum + 1
    hostnameCell = ("A{}".format(cellNum))
    ipaddrCell =   ("B{}".format(cellNum))
    try:
        hostname = socket.gethostbyaddr(ip)
        ipaddr = hostname[2]
        hostname = hostname[0]
        hostname = hostname.replace("#text that need to delimated#","")
        # print(hostname)
        # print(ipaddr[0])
        print(ipaddr[0], hostname)
        worksheet.update(hostnameCell, ipaddr[0])  #ipaddress injection to gsheet  
        time.sleep(1)
        worksheet.update(ipaddrCell,hostname )  #hostname injection to gsheet
        time.sleep(1)
        
    except socket.herror, e:
        errormsg = [ipaddr[0],"hostname not found"," "]
        print(errormsg[0])
        worksheet.update(hostnameCell, errormsg[0])  #errormsg injection to gsheet  
        time.sleep(1)
        worksheet.update(ipaddrCell,hostname )  #hostname injection to gsheet
        time.sleep(1)
        


















