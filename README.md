# AppsScript
Some work with Google Apps Scripts

## sheetsToDocs.gs

This takes a Google Spreadsheet of information, and will fill a Google Doc Template that you made. 

### Steps to Create
1) Create a New Google Doc
2) All fields that you want to import, need to be enclosed in double pound signs
    -  Example: "##TITLE##"
    -  Note: these tags need to be in the _BODY_ of the doc, not header or footer
    -  Example:
  
![Screen Shot 2021-06-10 at 10 44 54 AM](https://user-images.githubusercontent.com/47643209/121547618-4a356980-c9da-11eb-8986-5f700900ba28.png)


3) Create a New Google Sheet
4) Add Required Columns:
    - Create Document (Check Box)
    - Document Created (Empty - for now)
    - Document Link (Empty - for now)
    - Doc Name (this will be the name of the doc you will be creating)
    - Example:
   
![Google Sheet Setup](https://user-images.githubusercontent.com/47643209/121547438-1fe3ac00-c9da-11eb-8eb0-02467b01e331.png)


6) Ensure that the column headers (row 1) match your template fields
7) Copy this script into your own script editor

![Script Editor](https://user-images.githubusercontent.com/47643209/121548834-58d05080-c9db-11eb-8eb6-912c4313e0d9.png)

9) Copy over the custom fields
  - ID of your template into line __21__ of the script 
10) Set up a trigger on edit
    10a) Navigate Here
    
![Trigger](https://user-images.githubusercontent.com/47643209/121549196-a3ea6380-c9db-11eb-9b28-006ecb27121d.png)

   10b) Configure Your Trigger
   
   Note: You can set how often you want to get an update on this script, daily is default
    
![Configure Trigger](https://user-images.githubusercontent.com/47643209/121549271-b795ca00-c9db-11eb-891e-08024d2030c3.png)

11) New Document will be created in the same folder as your template. I reccomend giving your template a view only access so people don't mess it up


### Completed Sheet and Doc

You now can dynamicly add columns and template tags. As long as they match, they will be imported to your document.

If you aren't seeing the template be filled properly, check that the column name and the tag are exactly the same.

Formating is controlled by the Google Doc, not the sheet, so make sure your format is correct in the template.


You now will see your created document with the forms filled in:

![Created Document](https://user-images.githubusercontent.com/47643209/121547261-f9be0c00-c9d9-11eb-9329-633d99040754.png)

How the sheet looks upon completion:

![Example Sheet After Script Runs](https://user-images.githubusercontent.com/47643209/121547068-d09d7b80-c9d9-11eb-9a29-bb03807aa1ff.png)


### How it works
The script will trigger when you check the box in the "Create Document" column. It checks for a few things prior to running:
1) Did you just check this cell ("Create Document")
2) Is the cell named "Document Created" empty
    - if not, it will not run again
    - this is a great way to create a new template, delete the data in the "Dcoument Created" column, and check the box again


