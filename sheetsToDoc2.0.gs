function getDocOwner(){
  // Gets the list of emails to send notifications to, 
  // returns comma separated string of all emails
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Outline Document Owner");
  var email = mySheet.getRange(1,1,mySheet.getLastRow()).getValue();
  return email;
}

function completionBox(title,urlLink){
  // will display the dialog box when the script is done running
  var rawHTML = '<a onclick="google.script.host.close()" target="_blank" href="' + String(urlLink) + '">' + title + '</a>'
  var htmlSerivce = HtmlService.createHtmlOutput(rawHTML)
  SpreadsheetApp.getUi()
      .showModelessDialog(htmlSerivce, "Your outline has been created")
}

function outlineCreation() {
  // Declare current sheet for easy access
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // We will use these values a lot
  var headerRow = 2
  var headers = sheet.getRange('A'+headerRow+':'+headerRow).getValues();

  if (headers[0].includes("Create Outline")){
    // Declaring row/column values for ease of use
    var row = sheet.getActiveCell().getRow();
    var column = sheet.getActiveCell().getColumn();

    // Hard Coded Columns 
    // Needed for validation of document creation
    var checkBox = sheet.getRange(row,headers[0].indexOf("Create Outline")+1)
    var docCreated = sheet.getRange(row,headers[0].indexOf("Outline Created Date")+1);
    var newDocLink = sheet.getRange(row,headers[0].indexOf("Outline Link")+1);

    if (checkBox.getValue() === true && column === headers[0].indexOf("Create Outline") + 1 && docCreated.isBlank() === true) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Your outline is being created');

      var docTitle = sheet.getRange(row,headers[0].indexOf("Recommended <H1>")+1).getValue();
      
      //Make a copy of the template file
      var templateId = '1g2PMUcwT4zXYLzQf_SW3lczYKle-NmQY3wqbz8hAw10';
      var documentId = DriveApp.getFileById(templateId).makeCopy().getId();
      var documentName = '(internal) Rocket Auto | ' + docTitle + ' Outline'
      DriveApp.getFileById(documentId).setName(documentName + ' | '+ Utilities.formatDate(new Date(), "GMT+1",'MMM yyyy'));

      // Save the URL for future use, we will put it back in the sheet
      var url = 'https://docs.google.com/document/d/'+ documentId
      Logger.log("DOCUMENT CREATED at " + url)

      var body = DocumentApp.openById(documentId).getBody();
      // share the file with the domain in edit mode
      var newFile = DriveApp.getFileById(documentId)
      newFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT)
      // make the owner from the sheet
      try{
        newFile.setOwner(getDocOwner())
      }
      catch(e){
        Logger.log("CANNOT SET OWNER: " + e)
      }

      
      //var docHeader = DocumentApp.openById(documentId).getHeader();
      for(var i = 0; i < headers[0].length -1; i++){
        body.replaceText("##"+headers[0][i]+"##",sheet.getRange(row,i+1).getValue())
      }

      // Update sheet with Date of creation and the url to access
      docCreated.setValue(Utilities.formatDate( new Date(), "GMT+1",'MM/dd/yyyy'));
      newDocLink.setValue(String(url));
      
      completionBox(documentName,url);
    }
    else{
      Logger.log("Outline not created.")
    }
  }
  else{
    Logger.log("Header row (" + headerRow + ") did not contain 'Create Outline'")
  }
  
}
