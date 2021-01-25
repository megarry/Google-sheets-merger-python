'''
Libraries used: gspread                   --> For operations on sheets. (ex. Open Spreadsheet, Open Worksheet, Delete rows, etc.)
                pydrive                   --> For operations on files in Drive (ex. Iterate through available files, fetch properties of files, etc.)
                oauth2client              --> For Oauth2 
                pandas                    --> For manipulating data (ex. add custom columns in dataframe -> to add in sheets )
                time                      --> For converting datetime format into Unix time format

Reference for these libraries: 
                gspread                   --> https://gspread.readthedocs.io/en/latest/
                pydrive                   --> https://pythonhosted.org/PyDrive/
                oauth2client              --> https://pypi.org/project/oauth2client/
                pandas                    --> https://pandas.pydata.org/docs/
                time                      --> https://docs.python.org/3/library/time.html

--> 'Try' part of the code looks if any file is newly updated. If it founds any updated file, it deleted old data from Master file and appends new data only.
--> If new sheet is created in the folder from where we merge sheets to create the Master sheet, It is handled by the 'except' part of the code. It Merges all sheets irrespective of their prsence in Master sheet and new Master sheet replaces old Master sheet so that it contains all sheets merged without duplications.
'''
import sys
import gspread
from oauth2client.client import GoogleCredentials
import pandas as pd                                                                  
gc = gspread.oauth()                                                                 # Place client_secrets.json file in C:\Users\abc\AppData\Roaming\gspread
                                                                                     # Please create the gspread directory in Roaming directory if absent
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.client import GoogleCredentials
import time 
import datetime


# Authenticate and create the clients.
gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)                                                          # Place client_secrets.json file in the working directory

masterSheet = gc.open('Master')
wksMaster = gc.open('Master').sheet1

try:
    #query format: 'q': "'unique id of the folder where sheets are stored' in parents and 'other filters'"
    listed = drive.ListFile({'q': "'1js1nSOPfeXNIggsheodvb1r3PhEQokT5' in parents and mimeType = 'application/vnd.google-apps.spreadsheet'"}).GetList()
    for file in listed:
      sheetdate = file['modifiedDate']
      print(file['title']+" modified date : "+sheetdate)                                                 # Printing Last Modified date of the sheet
      masterCell = wksMaster.find(file['title'])
      masterDate = wksMaster.cell(masterCell.row, 4).value                                               # in place of 4,enter col.no of Modifieddate (row,col)
      print("date in master file: ",masterDate)                                                          # Printing Last Modified date of the Master Sheet
      if(time.mktime(datetime.datetime.strptime(sheetdate,"%Y-%m-%dT%H:%M:%S.%fZ").timetuple()) >        # converting datetime in Unix format
         time.mktime(datetime.datetime.strptime(masterDate,"%Y-%m-%dT%H:%M:%S.%fZ").timetuple())):
         print("Updating Master")
         cells = wksMaster.findall(file['title'])                                                        # Get list of rows where old data is present
         wksMaster.delete_rows(cells[0].row,cells[-1].row)                                               # Delete all rows of old data
         wksSheet = gc.open_by_key(file['id']).sheet1                                                    # Open newly modified sheet 
         rows = wksSheet.get_all_values()                                                                # Fetch all values (List of List) 
         df = pd.DataFrame.from_records(rows[1:],columns = rows[0])                                      # Storing new data into dataframe df (Dictionary)
         df['title'] = file['title']                                                                     # Adding Filename to the DF
         df['modifieddate'] = file['modifiedDate']                                                       # Adding Modified date to the DF
         data = df.values.tolist()                                                                       # Storing new data into data as List of List
         masterSheet.values_append('sheet1!A1:D1', {'valueInputOption' : 'USER_ENTERED', 
                                              'insertDataOption' : 'INSERT_ROWS'}, {'values' : data})    #Appending data 
         print("Master Updated")
      else:
        print("It is up-to-date")
     
except(gspread.exceptions.CellNotFound):
    def checkupdation():
        flag = 0
        print("Checking...")
        #title = 'Master' <-- put the name of sheet in which you want combined data. ex. Master (This should be pre-existing in the Drive with schema)
        listed = drive.ListFile({'q': "mimeType = 'application/vnd.google-apps.spreadsheet' and title = 'Master'"}).GetList()    
        for file in listed:
          masterdate = file['modifiedDate']                                           #masterdate is the datetime when the Master sheet was last updated 
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