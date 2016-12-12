function main() {
  var email = Session.getActiveUser().getEmail();
  var threads = GmailApp.getInboxThreads();
  var doiMatch = /SREP\d{5}/i
  var proofOut = /(Proofs of \[SREP_SREP)\d{5}\]/
  var subjects = [];
  
  //Get subjects from threads array
  for (var i = 0; i < threads.length; i++) {
     subjects[i] = threads[i].getFirstMessageSubject()
  };
  
  //Open Go Live, get active sheet and pull cell values into 2d array 'data'
  var golive = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1F9dr6Jcy-V8epCeDN8uIdvno3emcWG9jGN6sXlQyt3o/edit");
  SpreadsheetApp.setActiveSpreadsheet(golive);
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
 
  for (var i = 0; i < subjects.length; i++) {
    //If matches proof out email
    if (proofOut.test(subjects[i])) {
        doi = doiMatch.exec(subjects[i])
        doi = doi.toString().toLowerCase();
        Logger.log(doi);
      //Run through DOI column of data array, search for string match for sliced DOI
      for (var j = 1; j < data.length; j++) {
        var n = data[j][1].trim().toLowerCase().search(doi);
        //If subject DOI is found in DOI column, get corresponding cell in ProofOut col and set value to today's date
        if (n >= 0) {
          var cell = sheet.getRange(j+1, 7);
          var d = new Date().toDateString().slice(4);
          cell.setValue(d);
          //Archive email
          GmailApp.markThreadRead(threads[i]);
          GmailApp.moveThreadToArchive(threads[i]);
        };
      };
    };
  };
};



 
