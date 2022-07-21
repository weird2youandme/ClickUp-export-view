# ClickUp-export-view
This python application calls ClickUp api using a view-id and builds an Excel file. This is helpful if you do not have a business version and are limited in ability to export. The application will ask for the users API key and 

# Overview
If you have a free or Unlimited account with ClickUp, you are limited in exporting excel files. This application pulls the details of a viw and creates a XLSX file. 

The application requires the API key which you can get by going to settings->My APPS->Apps
The application also requires the workspace id which is the number following the forward slash in the url. Example https://app.clickup.com/########/v
The last item needed is the view-id  this is the last segment of the URL  example https://app.clickup.com/########/v/l/6-350936207-1

# How it works
ClickUp stores the current selection including filters and columns into a view. This view is what the API calls to retrieve the data so the result is the same filters and columns as you are viewing in the applicaion. This works on lists and table views.

The api pulls all of the details from custom fields for naming the columns correctly. It then pulls the elements of the view-id that is provided and begins to pull the data. Dates, numbers and text items are edited for excel.

# Options
Base items? The application will ask if you would like Base columns (URL, Space, folder, list, id) as these are not fields you can add to the view but can be very helpful when pulling from the everything list. 

Rebuild Custom fields? The application will also ask if you would like to pull the custom fields. This process can be slow depending on how many and how detailed the pull down options are. The custom field data is stored in a JSON file when done which is used if you select No. There is no need to rebuild this unless you are running for the first time or after creating new custom fields.  

# To use (python installed) 
Create a click up folder and place the program there and run it from command prompt. The result and custom json  will be placed in the same folder. I did not add the file logic so it assums everything in the same folder. 
I included an executable version for those who want the application without the code. 

# To use no python (windows exe)
Download the exe file and place in a file location. Run the exe and the custom fields json and resulting report will be placed in the same location. 

** This is a simple tool, I am not a developer so please do not critiscize my code. If anyone desires to fix it, that would be great. 
*** I am working on transitioning to openpyxl which will clean up a lot of the edits for commas and types.

