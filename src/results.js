function readResults(){
    Logger.log("Reading results from spreadsheet");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Results Summary Sheet");
    const data = sheet.getDataRange().getValues();

    //get settings
    const settings = getSettings(ss.getId());

    // Find the header row
    let headerRow = -1; // Initialize to an invalid row number
    for (let i = 0; i < data.length; i++) {
        if (data[i][0] === "Number") {
        headerRow = i;
        break; 
        }
    }
    if (headerRow === -1) {
        Logger.log("Header row not found");
        return;
    }
    // Extract the data from the header row
    const headers = data[headerRow];
    //get header indexes
    const nameIndex = headers.indexOf("Name");
    const gradeIndex = headers.indexOf("Grade");
    // Find the end row
    let endRow = -1;
    if (headerRow !== -1) { // Only search if the header row was found
        for (let i = headerRow + 1; i < data.length; i++) { 
        if (data[i][0] === "Tutor:") {
            endRow = i; // End row is one before "Tutor:"
            break;
        }
        }
    }
    // Process the data
    let learners = [];
    for (i=headerRow+1; i<endRow; i++){
        const name = data[i][nameIndex];
        const grade = data[i][gradeIndex];
        const learner = buildLearnerObject(name, ss);
        learner.grade = grade;
        learner.settings = settings;
        Logger.log(learner);
        learners.push(learner);
        //Check if a grade is present and learner has paid
        if(learner.grade != "" && learner.paid){
            //send the results
            emailResults(learner);
            //mark the results sent
            markResultsSent(name, ss);
        } else{
            Logger.log("Not sending results for " + name);
            Logger.log("Grade: " + learner.grade);
            Logger.log("Paid: " + learner.paid);
        }
    }
}

function testReadResults(){
    const ss = SpreadsheetApp.openById("1RPGTUZiJ4BUr0mpZLga_udpAa5izSnkn4pBXy5h__Dc");
    readResults(ss);
}

function markResultsSent(name, ss){
    Logger.log("Marking results sent for " + name);
    const sheet = ss.getSheetByName("Document Generator");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameIndex = headers.indexOf("Name");
    const resultsIndex = headers.indexOf("Results Sent");
    Logger.log("Results Index: " + resultsIndex);
    const row = data.findIndex(row => row[nameIndex] === name);
    Logger.log(name+" found at row " + row);
    if (row<1) {
      Logger.log(`Name "${name}" not found in the spreadsheet.`);
    } else {
        // Mark the results sent
        let resultsCell = sheet.getRange(row+1, resultsIndex+1);
        resultsCell.setValue(true);
        Logger.log("Marked results sent in cell " + resultsCell.getA1Notation());
    }
}

function addMissingResultsSentField(ss) {
    const sheet = ss.getSheetByName("Document Generator");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sentColumnIndex = headers.indexOf("Sent") + 1; // Get the index of "Sent" column
  
    if (sentColumnIndex === 0) {
      Logger.log("Column 'Sent' not found.");
      return; // Exit if "Sent" column doesn't exist
    }
  
    sheet.insertColumnAfter(sentColumnIndex); // Insert a new column after "Sent"
  
    // Set Header
    const headerRange = sheet.getRange(1, sentColumnIndex + 1);
    headerRange.setValue("Results Sent");
  
    // Set Data  
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const cell = sheet.getRange(i, sentColumnIndex + 1);
      cell.setValue(false);
      cell.insertCheckboxes();
    }
  }