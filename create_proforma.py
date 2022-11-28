import subprocess
import keyring
import sys
import win32com.client
import time
import csv
import os
import shutil
import re
import operator
import keyring
import datetime
import psutil

from win32com.client import Dispatch
from win32com import client


user = os.getlogin()
sapuser = "SAP USERNAME" # CHANGE TO YOUR SAP USERNAME
todays_date = datetime.datetime.today().strftime('%Y-%m-%d')

password = keyring.get_password('sap_password', f'{sapuser}')

print("Veryfying if SAP is running")

if "saplogon.exe" in (i.name() for i in psutil.process_iter()):
    print("SAP is already running")
else:
    print("SAP is not running. Starting...")
    subprocess.check_call(['C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe', '-system=PE1', '-client=500', f'-user={sapuser}', f'-pw={password}', '-language=EN'])
    time.sleep(20)

SapGuiAuto = win32com.client.GetObject("SAPGUI")
if not type(SapGuiAuto) == win32com.client.CDispatch:
    exit


application = SapGuiAuto.GetScriptingEngine
if not type(application) == win32com.client.CDispatch:
    SapGuiAuto = None
    exit


connection = application.Children(0)
if not type(connection) == win32com.client.CDispatch:
    application = None
    SapGuiAuto = None
    exit


session = connection.Children(0)
if not type(session) == win32com.client.CDispatch:
    connection = None
    application = None
    SapGuiAuto = None
    exit


print("Searching for order")

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text = "VL06O"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/btnBUTTON1").press()
session.findById("wnd[0]/usr/ctxtIF_VSTEL-LOW").text = "ika1"
session.findById("wnd[0]/usr/ctxtIT_KODAT-HIGH").text = f"{todays_date}"
session.findById("wnd[0]").sendVKey(8)
x = 5
external_list = []
intercompany_list = []
for i in range(10):
    try:
        ship_to_party = session.findById(f"wnd[0]/usr/lbl[119,{x}]").text
        proforma_number = session.findById(f"wnd[0]/usr/lbl[6,{x}]").text
        if "Icomera" not in ship_to_party:
            external_list.append(proforma_number)
        else:
            intercompany_list.append(proforma_number)
        x = x+1
    except Exception:
        pass
print("External list: ", external_list)
print()
print("intercompany list", intercompany_list)

session.findById("wnd[0]/tbar[0]/btn[12]").press()
session.findById("wnd[0]/tbar[0]/btn[12]").press()
session.findById("wnd[0]/tbar[0]/btn[12]").press()

for i in external_list:
    session.findById("wnd[0]/tbar[0]/okcd").text = "VF01"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/cmbRV60A-FKART").key = "ZF11"
    session.findById("wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,0]").text = f"{i}"
    session.findById("wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,0]").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nVF02"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[0]/btn[12]").press()

for i in intercompany_list:
    session.findById("wnd[0]/tbar[0]/okcd").text = "VF01"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/cmbRV60A-FKART").key = "ZF12"
    session.findById("wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,0]").text = f"{i}"
    session.findById("wnd[0]/usr/tblSAPMV60ATCTRL_ERF_FAKT/ctxtKOMFK-VBELN[0,0]").caretPosition = 10
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nVF02"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[0]/btn[12]").press()