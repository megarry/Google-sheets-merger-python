
<h1> Code for ease of work </h1>
<h2> sheetsMerger.py </h2>
<b>How to merge google sheets in a drive folder?</b><br>
- This code takes files from a specific folder in your Google Drive, merge them and stores in another sheet in Google Drive. When you Re-run this script, it checks the last modified dates on the separate sheets and the Master sheet and if any sheet is modified after last modification of Master sheet, It re merges all sheets and stores it in the Master.<br>
- I have used gspread and pydrive libraries for this code. <br>
- You will need to:  <br> </t> 1) Create a project in GCP <br> </t> 2) Enable Google Drive and Google Sheets API and download the credentials. <br>
- Required documentation is available in the Library docs.
