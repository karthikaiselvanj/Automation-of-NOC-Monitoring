import pandas as pd
import sys
import numpy as np
import pymysql
from xlrd import open_workbook
import xlrd

# To store the old data temporarily
#df2 =df.copy(deep=True)

# Data frame to store the old data
File_path = r"C:/Users/karthikais/Documents/Application_monitoring/App_monitoring.xlsx"
old=pd.read_excel(File_path, sheetname='Last_Update_status')
df2=pd.DataFrame(old)



#Connection Establishment to Database
conn=pymysql.connect(host='10.1.1.67',user='user',password='user',port=3306,db='monitor')

# To get present count
df=pd.read_sql_query('''SELECT projectname,DATE_FORMAT(FROM_UNIXTIME(TIMESTAMP),"%d-%m-%y"),
                        DATE_FORMAT(FROM_UNIXTIME(TIMESTAMP),"%H"),updating,notupdating,total 
                        FROM projectlastrecords ORDER BY total DESC''',conn)
df = df.rename(index=str, columns={"projectname": "Project", 
                              '''DATE_FORMAT(FROM_UNIXTIME(TIMESTAMP),"%d-%m-%y")''': "Date",
                             '''DATE_FORMAT(FROM_UNIXTIME(TIMESTAMP),"%H")''': "Data_for_Hour",
                             'updating' : 'Updating',
                              "notupdating" : "Not Updating",
                              "total" : "Total Installed Sites"                              
                             })

# Updating/ Not Updating Percent
df['Updating Percent'] = df['Updating'] / df['Total Installed Sites']
df['Not Updating Percent'] = df['Not Updating'] / df['Total Installed Sites']
# round off to decimals
df= df.round({'Updating Percent' : 1,
         'Not Updating Percent' : 1})

#Merging DF and DF2 
df = pd.merge(df, df2.iloc[:,0:4], on='Project', how='inner').drop(["Date_y","Data_for_Hour_y"],axis=1)
df = df.rename(index=str, columns={"Date_x": "Date",
                                  "Data_for_Hour_x": "Data_for_Hour",
                                  "Updating_x" : "Updating",
                                  "Updating_y" : "Last_updated_count",
                                  })
## Projects to Exclude
loc = (File_path) 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(1) 
sheet.cell_value(0, 0) 
Excluded_projects = []
for i in range(sheet.nrows): 
    Excluded_projects.append(sheet.cell_value(i, 0))
    
df = df[~df.Project.isin(Excluded_projects)]

config=pd.DataFrame()
config = pd.read_excel(File_path, sheetname='Drop_config')

df3 = pd.merge(df, config.iloc[:,0:4], on='Project', how='left')
df3['Drop_count'] = df3['Drop_count'].fillna(0)
df3['Drop_count']=np.where(df3['Drop_count']==0,df3['Last_updated_count']*0.2,df3['Drop_count'])
df3= df3.round({'Drop_count' : 0})


# email configuration
import smtplib
from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
recipients = ['sudhakar.v@invendis.com','babu.bj@invendis.com','support@invendis.com']

def send_emails (project,Updating,Not_Updating):
    fromaddr = 'app.alerts@itocsense.com'
    msg = MIMEMultipart()
    msg['Subject'] = project + "\t Updating count is getting low"
    body = "Updating count is {}. \nNot Updating Count is {}. \n".format(Updating,Not_Updating)
    msg.attach(MIMEText(body, 'plain'))
    server = SMTP('smtp.office365.com:25')
    server.ehlo()
    server.starttls()
    server.login('app.alerts@itocsense.com', 'Invendis@123')
    text = msg.as_string()
    server.sendmail(fromaddr, recipients, text)
    server.quit()


# Module for Desktop Notifications
from win32api import *
from win32gui import *
import win32con
import sys, os
import struct
import time

# Class
class WindowsBalloonTip:
    def __init__(self, title, msg):
        message_map = { win32con.WM_DESTROY: self.OnDestroy,}

        # Register the window class.
        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = 'PythonTaskbar'
        wc.lpfnWndProc = message_map # could also specify a wndproc.
        classAtom = RegisterClass(wc)

        # Create the window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow(classAtom, "Taskbar", style, 0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, 0, 0, hinst, None)
        UpdateWindow(self.hwnd)

        # Icons managment
        iconPathName = os.path.abspath(os.path.join( sys.path[0], 'balloontip.ico' ))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
            hicon = LoadImage(hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
            hicon = LoadIcon(0, win32con.IDI_APPLICATION)
        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, 'Tooltip')

        # Notify
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(NIM_MODIFY, (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20, hicon, 'Balloon Tooltip',"Drop count is {}".format(msg), 200, title))
        # self.show_balloon(title, msg)
        time.sleep(5)

        # Destroy
        DestroyWindow(self.hwnd)
        classAtom = UnregisterClass(classAtom, hinst)
    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0) # Terminate the app.

# Function
def balloon_tip(title, msg):
    w=WindowsBalloonTip(title, msg)
#Comparision condition
for index,rows in df3.iterrows():
    if rows['Last_updated_count']-rows['Updating'] > rows['Drop_count']:
        send_emails(rows[0],rows[3],rows[4])
        if __name__ == '__main__':
            balloon_tip(str(rows[0]),str(rows['Last_updated_count']-rows['Updating']))
        print(rows[0])
    else:
        continue
        


from openpyxl import load_workbook
#file = r'C:/Users/karthikais/Documents/Application_monitoring/App_monitoring.xlsx'
book=load_workbook(File_path)
writer=pd.ExcelWriter(File_path, engine="openpyxl")
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
df3.to_excel(writer, sheet_name="Last_Update_status", index=False)
writer.save()


