function main() {
  var form = FormApp.getActiveForm();
  var formResponses = form.getResponses();
  var formResponse = formResponses[0];
  var itemResponses = formResponse.getItemResponses();
  var dois = itemResponses[0].getResponse();
  var scheduleDate = itemResponses[1].getResponse();
  dois = dois.split("\n").filter(Boolean);
  
  /*
  //Test content
  var scheduleDate = "19/05/2016";
  var dois = ["srep25042","srep25046"];
  */
    
  //Open Go Live, get sheets
  var golive = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1F9dr6Jcy-V8epCeDN8uIdvno3emcWG9jGN6sXlQyt3o/edit");
  SpreadsheetApp.setActiveSpreadsheet(golive);
  var sheetInProduction = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("In Production");
  var sheetScheduled = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scheduled");
  var sheetLookup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lookup");
  var movedRows = [];
  var skippedRows = [];
  
  //For each doi from form, find in Go Live; set pub date & move row if matched
  for (i = 0; i < dois.length; i++) {
    var lastRowInProduction = sheetInProduction.getLastRow();
    var doisInProduction = sheetInProduction.getRange(2,2,lastRowInProduction,1).getValues();
    var found = false;
    
    for (j = 0; j < doisInProduction.length; j++) { 
      var doiMatch = doisInProduction[j].toString().trim().toLowerCase().search(dois[i]);
      if (doiMatch >= 0) {
        var cellPubDate = sheetInProduction.getRange(j+2, 14); //This is j+2 because the array 'dataInProduction' goes from (0,0), whereas the sheet InProduction starts at (1,1), plus sheet also has column header row
        cellPubDate.setValue(scheduleDate);
        var row = sheetInProduction.getRange(j+2,1,1,15); //Get contents of whole row
        var targetRowNo = sheetScheduled.getLastRow() + 1;
        var targetRow = sheetScheduled.getRange(targetRowNo,1);
        row.copyTo(targetRow);
        found = true;
        
        
        // Get email addresses
        var dataEmailAddress = "";
        var lastRowLookup = sheetLookup.getLastRow();
        var doisLookup = sheetLookup.getRange(2,6,lastRowLookup,1).getValues();
        for (k = 0; k < doisLookup.length; k++) {
          var doiMatch2 = doisLookup[k].toString().search(dois[i]);
          if (doiMatch2 >= 0) {
            dataEmailAddress = sheetLookup.getRange(k+2,5).getValue().toString();        
            break;
          };
        };
        
        
        // Compile paper details
        var paperDetails = {
          doi: dois[i],
          name: sheetInProduction.getRange(j+2, 4).getValue().toString(),
          emailAddress: dataEmailAddress,
          detailsTableRow: function(){return "<tr><td>" + this.doi + "</td><td>" + this.name + "</td><td>" + this.emailAddress + "</td></tr>"}
        };
        
        
        sheetInProduction.deleteRow(j+2); //This should be very last operation as it messes up cell refs by removing a row
        break;
      }
    }
    
    if (found == true) {
      movedRows.push(paperDetails);
    } else {
      skippedRows.push(dois[i]);
    } 
  }
  

 //Email results, delete form responses to reset for next input
 var robotBlob = DriveApp.getFileById("0B6GrjeVCGUCfN3lYdFZubERLcDQ").getBlob(); //0B6GrjeVCGUCfN3lYdFZubERLcDQ
 var emailBody = ""
 var emailBodyHTML = "<style> \
              table, th, td { \
                font-family: arial; \
                font-size: 13px; \
              } \
              th, td { \
                padding: 5px; \
              } \
              p { \
                font-family: arial; \
                font-size: 13px; \
              } \
            </style>"

            
            
 // Output moved paper details into table
 if (movedRows.length > 0) {
   emailBody += "The following papers have been scheduled for " + scheduleDate +
         ":\n\n" + paperDetails.detailsTableRow().toString();   
   
   emailBodyHTML += "<p style='font-family:arial;'>The following papers have been scheduled for " + scheduleDate +
               ":<br/><br/><table><tr><strong><td>DOI</td><td>Author</td><td>Email address</td></strong></tr>" + outputMovedRows(movedRows) + "</table>";                             
 };
  
 /* Email with no paper details           
 if (movedRows.length > 0) {
   emailBody += "The following papers have been scheduled for " + scheduleDate +
         ":\n\n" + movedRows.toString().replaceAll(",", "\n");
   emailBodyHTML += "<p style='font-family:arial;'>The following papers have been scheduled for " + scheduleDate +
               ":<br/><br/>" + movedRows.toString().replaceAll(",", "<br/>");                            
 };
 */
  
 if (skippedRows.length > 0) {
   emailBody += "\nThe following papers were not scheduled due to an error " +
         ":\n\n" + skippedRows.toString().replaceAll(",", "\n");
   emailBodyHTML += "<p style='font-family:arial;'><br/>The following papers were not scheduled due to an error " +
               ":<br/><br/>" + skippedRows.toString().replaceAll(",", "<br/>");
 };
  
 if (movedRows.length == 0 && skippedRows.length == 0) { 
    emailBody = "No DOIs were inputted\n";
    emailBodyHTML = "No DOIs were <strike>inpat</strike>inputted<br/>";
 };
 
 MailApp.sendEmail({
   to: "scirep.production@nature.com", 
   subject: "Papers scheduled for " + scheduleDate, 
   body: emailBody + "\n\nLove and kisses,\nThe Scirep Robot",
   htmlBody: emailBodyHTML + "<br/><br/>Love and kisses,<br/>The Scirep Robot</p><br/>" + "<img src='cid:robotLogo'>",
   inlineImages: {robotLogo: robotBlob}                    
 });
 
 form.deleteAllResponses();
};
  
String.prototype.replaceAll = function(search, replace) {
    if (replace === undefined) {
        return this.toString();
    }
    return this.replace(new RegExp('[' + search + ']', 'g'), replace);
};

function outputMovedRows(arr) {
  var contents = ""
  for (var x = 0; x < arr.length; x++) {
  contents += arr[x].detailsTableRow();
  }
  return contents;
};

function test() {
  //Test content
  var scheduleDate = "19/05/2016";
  var dois = ["srep23887","srep25046"];
  
  //Open Go Live, get sheets
  var golive = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1F9dr6Jcy-V8epCeDN8uIdvno3emcWG9jGN6sXlQyt3o/edit");
  SpreadsheetApp.setActiveSpreadsheet(golive);
  var sheetInProduction = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("In Production");
  var sheetScheduled = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scheduled");
  var sheetLookup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lookup");
  var movedRows = [];
  var skippedRows = [];
  
        var dataEmailAddress = "";
        var lastRowLookup = sheetLookup.getLastRow();
        var doisLookup = sheetLookup.getRange(2,6,lastRowLookup,1).getValues();
        for (k = 0; k < doisLookup.length; k++) {
          var doiMatch2 = doisLookup[k].toString().search(dois[0]);
          if (doiMatch2 >= 0) {
            dataEmailAddress = sheetLookup.getRange(k+1,5).getValue().toString();
            Logger.log(dataEmailAddress);
            break;
          };
        };

};
  
