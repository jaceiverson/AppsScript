function fillTemplate() {
  // Declare current sheet for easy access
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // We will use these values a lot
  var headers = sheet.getRange('A1:1').getValues();

  // Check to make sure we are on the right sheet
  if (headers[0].includes("Create Document")){
    // Declaring row value for ease of use
    var row = sheet.getActiveCell().getRow();

    // Hard Coded Columns 
    // Needed for validation of document creation
    var checkBox = sheet.getRange(row,headers[0].indexOf("Create Document")+1)
    var docCreated = sheet.getRange(row,headers[0].indexOf("Document Created")+1);
    var newDocLink = sheet.getRange(row,headers[0].indexOf("Document Link")+1);
    
    // Check to make sure that we havn't created a doc already
    if (checkBox.getValue() === true && sheet.getActiveCell().getColumn() === headers[0].indexOf("Create Document") + 1 && docCreated.isBlank() === true) {

      //Make a copy of the template file
      var templateId = 'YOUR TEMPLATE ID';
      var documentId = DriveApp.getFileById(templateId).makeCopy().getId();
      var documentName = sheet.getRange(row,headers[0].indexOf("Doc Name")+1).getValue();
      DriveApp.getFileById(documentId).setName(documentName + ' | '+ Utilities.formatDate(new Date(), "GMT+1",'MMM yyyy'));

      // Save the URL for future use, we will put it back in the sheet
      var url = 'https://docs.google.com/document/d/'+ documentId
      Logger.log("DOCUMENT CREATED at " + url)

      // Replace all the template tags with sheets values
      var body = DocumentApp.openById(documentId).getBody();
      for(var i = 0; i < headers[0].length -1; i++){
        body.replaceText("##"+headers[0][i]+"##",sheet.getRange(row,i+1).getValue())
      }

      // Update sheet with Date of creation and the url to access
      docCreated.setValue(Utilities.formatDate( new Date(), "GMT+1",'MM/dd/yyyy'));
      newDocLink.setValue(String(url));
}
}
}
