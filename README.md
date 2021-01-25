
<h1> Code for ease of work </h1>
<h2> sheetsMerger.py </h2>
<b>How to merge google sheets in a drive folder?</b><br>
- This code takes files from a specific folder in your Google Drive, merge them and stores in another sheet in Google Drive. When you Re-run this script, it checks the last modified dates on the separate sheets and the Master sheet and if any sheet is modified after last modification of Master sheet, It re merges all sheets and stores it in the Master.<br>
- I have used gspread and pydrive libraries for this code. <br>
- You will need to:  <br> </t> 1) Create a project in GCP <br> </t> 2) Enable Google Drive and Google Sheets API and download the credentials.  <a href= "https://developers.google.com/sheets/api/quickstart/python" > How to </a> <br>
- Required documentation is available in the Library docs.

<h2> sheetsMergerV2.py </h2>
- Separate files have normal data <br>
- Master file has merged data + filename (where the data came from) and Last modified date-time of that sheet.<br>
- This code, when run, compares the Last modified date time of the sheet with the Last modified date-time recorded in the Master sheet.<br>
- If the sheet has been updated after the last Merge operation, Last modified date-time for that file changes.<br>
- so the code compares these timestamps and updates the Merged sheet by inserting only the updated data. (Does not merge all files unlike the V1)<br>
- It does not support if the separate sheets folder is added with new sheets. Throws error.

<h2> sheetsMergerV3.py </h2>
- Supports all the functionality of V2 
- Supports if new file is added to separate sheets folder.
