# Sharepoint-Uploader
Tool to upload data to Sharepoint 

A simple Tool to upload bulk data to a sharepoint list (100000+ records); Tested with Sharepoint online. 

Copy paste on browser to upload 10000+ results in hanging the browser and does not ensure a succesful commit, with this app Bulk upload can happen easily

Simple Steps for Upload


### Step 1: Enter the Sharepoint URL
ex; https://www.abc22.sharepoint.com/teams/websiteName
![STEP 1](https://github.com/brvinodh/Sharepoint-Uploader/blob/master/Images/1.jpg)

### Step 2: Popup appears where you will need to enter your sharepoint credentails
![STEP 1](https://github.com/brvinodh/Sharepoint-Uploader/blob/master/Images/2.jpg)

#### Step 3: Select the List/Table on Sharepoint which you need to upload to.
This list should be auto-populated after user authentication
![STEP 1](https://github.com/brvinodh/Sharepoint-Uploader/blob/master/Images/3.jpg)

### Step 4: Select the Primary Key for the List.
Ex: Email ID if you are uploading to an Employee Table. 
Based on this key the program would compare if the record exists already on Sharepoint if yes then update else insert
![STEP 1](https://github.com/brvinodh/Sharepoint-Uploader/blob/master/Images/4.jpg)

### Step 5: Insert Tab Seperated Data in the "Update Data" Tab 
![STEP 1](https://github.com/brvinodh/Sharepoint-Uploader/blob/master/Images/5.jpg)

### Step 6: The tab seperated data would now be parsed and split into a table format. 
Map each column data against the corresponding column name
![STEP 1](https://github.com/brvinodh/Sharepoint-Uploader/blob/master/Images/6.jpg)

### Step 7: Continue mapping for all the columns
![STEP 1](https://github.com/brvinodh/Sharepoint-Uploader/blob/master/Images/7.jpg)

### Steps 8: Once Done with Mapping hit "Start Update" button at the bottom Right
![STEP 1](https://github.com/brvinodh/Sharepoint-Uploader/blob/master/Images/9.jpg)

### Step 9: The Update program would now run and you should have the status complete
![STEP 1](https://github.com/brvinodh/Sharepoint-Uploader/blob/master/Images/10.jpg)



