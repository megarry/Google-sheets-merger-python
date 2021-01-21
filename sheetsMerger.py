import sys
import gspread
from oauth2client.client import GoogleCredentials
import pandas as pd
gc = gspread.oauth()

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.client import GoogleCredentials


# Authenticate and create the PyDrive client.
gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

import time 
import datetime 


def checkupdation():
    flag = 0
    print("Checking...")
    #title = 'Master' <-- put the name of sheet in which you want combined data. ex. Master
    listed = drive.ListFile({'q': "mimeType = 'application/vnd.google-apps.spreadsheet' and title = 'Master'"}).GetList()    
    for file in listed:
      masterdate = file['modifiedDate'] #masterdate is the datetime when the Master sheet was last updated 
    print("Master Date: " ,masterdate)
    
    #query format: 'q': "'unique id of the folder where sheets are stored' in parents and other filters"
    listed = drive.ListFile({'q': "'1js1nSOPfeXNIggsheodvb1r3PhEQokT5' in parents and mimeType = 'application/vnd.google-apps.spreadsheet'"}).GetList()
    for file in listed:
        print("Checking: ",file['title'])
        sheetdate = file['modifiedDate']   #sheetdate is the datetime when sheet which is to be merged is updated
        print("Sheet date: ", sheetdate)
        if(time.mktime(datetime.datetime.strptime(sheetdate,"%Y-%m-%dT%H:%M:%S.%fZ").timetuple()) >  #converting datetime in Unix format
        time.mktime(datetime.datetime.strptime(masterdate,"%Y-%m-%dT%H:%M:%S.%fZ").timetuple())):
            flag = 1
            break
    print(flag)
    return flag


def mergeSheets():
    print("Merging Sheets...")
    frames = []
    #loop for getting the files to be merged
    listed = drive.ListFile({'q': "'1js1nSOPfeXNIggsheodvb1r3PhEQokT5' in parents and mimeType = 'application/vnd.google-apps.spreadsheet'"}).GetList()
    for file in listed:
      worksheet = gc.open_by_key(file['id']).sheet1
      rows = worksheet.get_all_values()
      df = pd.DataFrame.from_records(rows[1:],columns = rows[0])
      df['title'] = file['title']
      df['modifieddate'] = file['modifiedDate']
      frames.append(df)

    #concatinating all spreadsheets
    merged = pd.concat(frames)
    #Updating the master
    print("Updating Master file")
    mergedData = [merged.columns.to_list()] + merged.to_numpy().tolist()
    wsMaster = gc.open_by_key("1ljg8DztCx41D6Avhk2g8tgJ9CU7KRyoR74sBdqM0bMY").worksheet("Sheet1")
    wsMaster.update("A1",mergedData,value_input_option="USER_ENTERED")
    print("Done Updating")

def main():
    masterdate = 0
    sheetdate = 0
    
    print("In main function...")
    flag = checkupdation()
    if(flag == 1):
      mergeSheets()
    else:
      print("Exiting without Updation")
      sys.exit()

main()