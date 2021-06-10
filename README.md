# AppsScript
Some work with Google Apps Scripts

## sheetsToDocs.gs

This takes a Google Spreadsheet of information, and will fill a Google Doc Template that you made. 

Steps
1) Create a New Google Doc
2) All fields that you want to import, need to be enclosed in double pound signs
    -  Example: "##TITLE##"
    -  Note: these tags need to be in the _BODY_ of the doc, not header or footer
3) Create a New Google Sheet
4) Add Required Columns:
    - Create Document (Check Box)
    - Document Created (Empty - for now)
    - Document Link (Empty - for now)
    - Doc Name (this will be the name of the doc you will be creating)
6) Ensure that the column headers (row 1) match your template fields
7) Copy this script into your own script editor
8) Copy over the custom fields
  - ID of your template into line __21__ of the script 
9) Set up a trigger on edit

You now can dynamicly add columns and template tags. As long as they match, they will be imported in. 

If you aren't seeing the template be filled properly, check that the column name and the tag are exactly the same.
