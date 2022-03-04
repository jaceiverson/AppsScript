// Declare some constants
const HEADERROW = 1
const HEADERS = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('A'+HEADERROW+':'+HEADERROW).getValues();
const DOC_TITLE_COLUMN = "TITLE"
const TEMPLATE_TAB_NAME = "Templates"
const CONFIG_TAB_NAME = "Configuration"


function onOpen(e) {
  // Each time the sheet is opened we will create the custom menu bar
  updateMenu()
}

function onEdit(e){
  // Each time an edit is made, if our Click Based Creation is True (see toggle) 
  // each time a checkbox is checked, an outline will be created
  // This only happens if it hasn't been done before
  if (getClickBasedCreationValue()){
    outlineCreationCheckboxOnEdit()
  }
}

function updateMenu(){
  // Will update the menu bar with our custom outline creation menu
  SpreadsheetApp.getUi()
      .createMenu('Outline Creation Menu')
      .addSubMenu(
        SpreadsheetApp.getUi().createMenu('Bulk Create Outlines')
              .addItem("All Rows", 'bulkMenu')
              .addItem("Current Selection","selectionMenu")
        )
      .addSeparator()
      .addItem("Toggle Click Based Creation (Currently: " + getClickBasedCreationValue().toString().toUpperCase() + ")","toggleClickCreation")
      .addSeparator()
      .addItem("Outline Creation Help","helpMenu")
      .addToUi();
}

function bulkMenu(){
  // function called when menu item selected
  var startRow = HEADERROW + 1
  var numRows = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow()
  bulkOutlineCreation(startRow,numRows)
}

function selectionMenu(){
  // function called when menu item selected
  var startRow = SpreadsheetApp.getActiveRange().getRow()
  var numRows = SpreadsheetApp.getActiveRange().getNumRows()
  bulkOutlineCreation(startRow,numRows)
}

function helpMenu(){
  // function called when menu item selected
  var documentationURL = "https://docs.google.com/document/d/1TPFwAZiTP1eed5QA1Jsv3vtQL2UtQRG_y9CZBLGv3NU/edit#heading=h.65i2rbcp33pi"
  openUrl(documentationURL)
}

function openUrl(url){
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed to documentation.</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening Documentation..." );
}

function getDocOwner(){
  // Gets the emails to assign the owner
  // Comes from the config tab
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_TAB_NAME);
  return mySheet.getRange("B1").getValue();
}

function getClickBasedCreationValue(){
  // Returns True or False based on what value is on the config Tab
  // This determines if we are able to create outlines on click
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_TAB_NAME);
  return mySheet.getRange("B2").getValue();
}

function findCheckBoxes(rowStart,rowCount,columnName) {
  // Returns the number of checkboxes (true) in the passed in column name
  var checkbox = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(rowStart,HEADERS[0].indexOf(columnName)+1,rowCount).getValues();
  var counter = 0
  for (var i =0; i < checkbox.length; i++){
    if (checkbox[i][0] == true){
      counter++;
    }
  }
  return counter;
}

function findEmptyCells(rowStart,rowCount,columnName) {
  // Returns the number of empty cells in the passed in column name
  var checkbox = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(rowStart,HEADERS[0].indexOf(columnName)+1,rowCount).getValues();
  var counter = 0
  for (var i =0; i < checkbox.length; i++){
    if (checkbox[i][0] == null | checkbox[i][0] == ""){
      counter++;
    }
  }
  return rowCount - counter
}

function confirmationBox(boxTitle,boxMessage){
  // Asks for confirmation to create an outline
  var ui = SpreadsheetApp.getUi();
  var createOutline = ui.alert(boxTitle, boxMessage, ui.ButtonSet.YES_NO);
  if (createOutline == ui.Button.YES) {
    return true
    } 
  else {
    return false
    }
}

function completionBox(title,urlLink){
  // will display the dialog box when the script is done running
  var rawHTML = '<a onclick="google.script.host.close()" target="_blank" href="' + String(urlLink) + '">' + title + '</a>'
  var htmlSerivce = HtmlService.createHtmlOutput(rawHTML).setWidth(300).setHeight(100)
  SpreadsheetApp.getUi()
      .showModelessDialog(htmlSerivce, "Your outline has been created")
}

function getTemplateID(templateName){
  // given a template name, returns the template id
  Logger.log("TEMPLATE: " + templateName)
  //if templateName is Null, we will default to Base
  if (templateName == null | templateName == ""){
    templateName = "Base Template"
  }
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATE_TAB_NAME);
  var templateNameTable = mySheet.getRange(1,1,mySheet.getLastRow()).getValues();
  var templateIdTable = mySheet.getRange(1,2,mySheet.getLastRow()).getValues();
  for(nn=0;nn<templateNameTable.length;++nn){
    if (templateNameTable[nn][0]==templateName){break} ;// if a match in column B is found, break the loop
      }
  return templateIdTable[nn][0]
}

function bulkOutlineCreation(startRow,numRows){
  // Will run the outline creation on each checkbox that doesn't already have an outline
  var checkedBoxes = findCheckBoxes(startRow,numRows,"Create Outline")
  var docsWithCreateDate = findEmptyCells(startRow,numRows,"Outline Created Date")
  var promptQuestion = "You are attempting to create " + (checkedBoxes-docsWithCreateDate) + " outlines (You have "+ checkedBoxes + " checked boxes and "+ docsWithCreateDate+" already created outlines). Would you like to continue?"
  // Asks user for confirmation
  if (confirmationBox("Bulk Outline Creation",promptQuestion)){
    // Loop through each row passed in
    for (var row = startRow; row<(startRow+numRows); row++){
      outlineCreationBulk(row)
    }
    Logger.log("Bulk creation complete.\nstart row: " + startRow + " row count: " + numRows)
  }
  else{
    Logger.log("Bulk creation button selection NO.\nstart row: " + startRow + " row count: " + numRows)
  }
}

function toggleClickCreation(){
  // Toggles if you want to enable the ability to create outlines on click
  var currentSelection = getClickBasedCreationValue()
  var promptQuestion = "Currently the ability to create outlines with the checkboxes is set to ---" + currentSelection.toString().toUpperCase() + "---.\n\n" +
                        "Would you like to change that to ---" + (!currentSelection).toString().toUpperCase() +  "--- ?"
  if (confirmationBox("Enable Click Based Outline Creation",promptQuestion)){
    var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_TAB_NAME);
    mySheet.getRange("B2").setValue(!currentSelection);
    // Notify user of the change
    SpreadsheetApp.getActiveSpreadsheet().toast("Click Based Outline Creation set to " + getClickBasedCreationValue().toString().toUpperCase())
    Logger.log("Click Based Outline Creation set to " + getClickBasedCreationValue())
    // reset the Menu to reflect updated value
    updateMenu()
  }
  else{
    Logger.log("Click Based Outline Creation not changed " + currentSelection)
  }
}

function outlineCreationCheckboxOnEdit() {
  // Declare current sheet for easy access
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // We will use these values a lot
  var HEADERS = sheet.getRange('A'+HEADERROW+':'+HEADERROW).getValues();

  if (HEADERS[0].includes("Create Outline")){
    // Declaring row/column values for ease of use
    var row = sheet.getActiveCell().getRow();
    var column = sheet.getActiveCell().getColumn();

    // Hard Coded Columns 
    // Needed for validation of document creation
    var checkBox = sheet.getRange(row,HEADERS[0].indexOf("Create Outline")+1);
    var docCreated = sheet.getRange(row,HEADERS[0].indexOf("Outline Created Date")+1);
    var newDocLink = sheet.getRange(row,HEADERS[0].indexOf("Outline Link")+1);

    if (checkBox.getValue() === true && column === HEADERS[0].indexOf("Create Outline") + 1 && docCreated.isBlank() === true) {
      var docTitle = sheet.getRange(row,HEADERS[0].indexOf(DOC_TITLE_COLUMN)+1).getValue();
      var documentName = '(internal) | ' + docTitle + ' Outline'

      // Give dialog box to confirm (YES OR NO Buttons)
      if (confirmationBox('Outline Creation','You are going to create an outline for\n\n"' + docTitle + '"\n\nWould you like to continue?')){
        // give a popup
        SpreadsheetApp.getActiveSpreadsheet().toast("Your outline:\n\n" + documentName + "\n\nis being created");

        // Get the template name from the column (if it exists)
        // If it doesn't (the column) we will use the base template
        try{
          var templateName = sheet.getRange(row,HEADERS[0].indexOf("Outline Template")+1).getValue();
        }
        catch{
          var templateName = null
        }
        var templateId = getTemplateID(templateName);
        var documentId = DriveApp.getFileById(templateId).makeCopy().getId();
        DriveApp.getFileById(documentId).setName(documentName + ' | '+ Utilities.formatDate(new Date(), "GMT+1",'MMM yyyy'));

        // Save the URL for future use, we will put it back in the sheet
        var url = 'https://docs.google.com/document/d/'+ documentId
        Logger.log("DOCUMENT CREATED at " + url)

        // share the file with the domain in edit mode
        var newFile = DriveApp.getFileById(documentId)
        // newFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT)
        // make the owner from the sheet
        try{
          newFile.setOwner(getDocOwner())
        }
        catch(e){
          Logger.log("CANNOT SET OWNER: " + e)
        }
        
        var body = DocumentApp.openById(documentId).getBody();
        //var docHeader = DocumentApp.openById(documentId).getHeader();
        for(var i = 0; i < HEADERS[0].length -1; i++){
          body.replaceText("##"+HEADERS[0][i]+"##",sheet.getRange(row,i+1).getValue())
        }

        // Update sheet with Date of creation and the url to access
        docCreated.setValue(Utilities.formatDate( new Date(), "GMT+1",'MM/dd/yyyy'));
        var hyperlink = '=HYPERLINK("' + url+ '","' + documentName + '")';
        newDocLink.setFormula(hyperlink);
        
        completionBox(documentName,url);
      }
      else{
        SpreadsheetApp.getActiveSpreadsheet().toast('Your outline was not created.');
        Logger.log("USER selected NO to create outline.")
      }
    }
    else{
      Logger.log("Outline not created.")
    }
  }
  else{
    Logger.log("Header row (" + HEADERROW + ") did not contain 'Create Outline'")
  }
  
}

function outlineCreationBulk(row) {
  // Declare current sheet for easy access
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (HEADERS[0].includes("Create Outline")){
    // Hard Coded Columns 
    // Needed for validation of document creation
    var checkBox = sheet.getRange(row,HEADERS[0].indexOf("Create Outline")+1)
    var docCreated = sheet.getRange(row,HEADERS[0].indexOf("Outline Created Date")+1);
    var newDocLink = sheet.getRange(row,HEADERS[0].indexOf("Outline Link")+1);

    if (checkBox.getValue() === true && docCreated.isBlank() === true) {
      var docTitle = sheet.getRange(row,HEADERS[0].indexOf(DOC_TITLE_COLUMN)+1).getValue();
      var documentName = '(internal) | ' + docTitle + ' Outline'
      // give a popup
      SpreadsheetApp.getActiveSpreadsheet().toast("Your outline:\n\n" + documentName + "\n\nis being created");

      // Get the template name from the column (if it exists)
      // If it doesn't (the column) we will use the base template
      try{
        var templateName = sheet.getRange(row,HEADERS[0].indexOf("Outline Template")+1).getValue();
      }
      catch{
        var templateName = null
      }
      var templateId = getTemplateID(templateName);
      var documentId = DriveApp.getFileById(templateId).makeCopy().getId();
      DriveApp.getFileById(documentId).setName(documentName + ' | '+ Utilities.formatDate(new Date(), "GMT+1",'MMM yyyy'));

      // Save the URL for future use, we will put it back in the sheet
      var url = 'https://docs.google.com/document/d/'+ documentId
      Logger.log("DOCUMENT CREATED at " + url)

      // share the file with the domain in edit mode
      var newFile = DriveApp.getFileById(documentId)
      // newFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT)
      // make the owner from the sheet
      try{
        newFile.setOwner(getDocOwner())
      }
      catch(e){
        Logger.log("CANNOT SET OWNER: " + e)
      }
      
      var body = DocumentApp.openById(documentId).getBody();
      //var docHeader = DocumentApp.openById(documentId).getHeader();
      for(var i = 0; i < HEADERS[0].length -1; i++){
        body.replaceText("##"+HEADERS[0][i]+"##",sheet.getRange(row,i+1).getValue())
      }

      // Update sheet with Date of creation and the url to access
      docCreated.setValue(Utilities.formatDate( new Date(), "GMT+1",'MM/dd/yyyy'));
      var hyperlink = '=HYPERLINK("' + url+ '","' + documentName + '")';
      newDocLink.setFormula(hyperlink);
      
      completionBox(documentName,url);
    }
    else{
      Logger.log("Outline not created.")
    }
  }
  else{
    Logger.log("Header row (" + HEADERROW + ") did not contain 'Create Outline'")
  }

}
