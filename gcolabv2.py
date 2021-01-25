import gspread
from oauth2client.client import GoogleCredentials
import pandas as pd
gc = gspread.oauth()

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.client import GoogleCredentials
import re
import time 
import datetime


# Authenticate and create the PyDrive client.
gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

masterSheet = gc.open('Master')
wksMaster = gc.open('Master').sheet1

#query format: 'q': "'unique id of the folder where sheets are stored' in parents and other filters"
listed = drive.ListFile({'q': "'1js1nSOPfeXNIggsheodvb1r3PhEQokT5' in parents and mimeType = 'application/vnd.google-apps.spreadsheet'"}).GetList()
for file in listed:
  sheetdate = file['modifiedDate']
  print(file['title']+" modified date : "+sheetdate)                                                 # Printing Modified date of the sheet
  masterCell = wksMaster.find(file['title'])
  masterDate = wksMaster.cell(masterCell.row, 4).value
  print("date in master file: ",masterDate)                                                          # Printing Modified date of the Master Sheet
  if(time.mktime(datetime.datetime.strptime(sheetdate,"%Y-%m-%dT%H:%M:%S.%fZ").timetuple()) >        # converting datetime in Unix format
     time.mktime(datetime.datetime.strptime(masterDate,"%Y-%m-%dT%H:%M:%S.%fZ").timetuple())):
     print("Updating Master")
     cells = wksMaster.findall(file['title'])                                                        # Deleting existing rows if old data present
     wksMaster.delete_rows(cells[0].row,cells[-1].row)                                               # Deleting existing rows if old data present
     wksSheet = gc.open_by_key(file['id']).sheet1                                                    # Deleting existing rows if old data present
     rows = wksSheet.get_all_values()                                                                # Storing new data into dataframe 
     df = pd.DataFrame.from_records(rows[1:],columns = rows[0])                                      # Storing new data into dataframe
     df['title'] = file['title']                                                                     # Adding Filename to the DF
     df['modifieddate'] = file['modifiedDate']                                                       # Adding Modified date to the DF
     data = df.values.tolist()                                                                       # Storing new data into dataframe
     masterSheet.values_append('sheet1!A1:D1', {'valueInputOption' : 'USER_ENTERED', 
                                          'insertDataOption' : 'INSERT_ROWS'}, {'values' : data})    #Appending data 
     print("Master Updated")
  else:
    print("All is well")
	
