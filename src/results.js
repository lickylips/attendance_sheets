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

function processResults(){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(ss);
    const learners = getResultsFromSheet(ss);
    for(let learner of learners){
        Logger.log(learner);
        if(learner.grade != "" && learner.paid){
            //send the results
            emailResults(learner);
            //mark the results sent
            markResultsSent(learner.getName(), ss);
        } else{
            Logger.log("Not sending results for " + learner.getName());
            Logger.log("Grade: " + learner.grade);
            Logger.log("Paid: " + learner.paid);
        }
    }

}

function getResultsFromSheet(ss){
    Logger.log("Reading results from spreadsheet");
    const settings = getSettings(ss.getId());
    let learners = [];
    const sheets = ss.getSheets();
    for(let sheet of sheets){
        if(sheet.getName().includes("Results Summary Sheet")){
            let data = sheet.getDataRange().getValues();
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
            Logger.log("headerRow: " + headerRow + " endRow: " + endRow);
            //get rows:
            let rows = [];
            for (i=headerRow+1; i<endRow; i++){
                Logger.log("adding row: " + i);
                Logger.log(data[i]);
                let row = data[i];
                rows.push(row);
            }
            for(let row of rows){
                const name = row[nameIndex];
                const grade = row[gradeIndex];
                const learner = buildLearnerObject(name, ss);
                learner.grade = grade;
                learner.settings = settings;
                learners.push(learner);
            }
        }
    }
    return learners;
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

  function generateResultsSummarySheet(){
    //Open the spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    //get the learners array
    const learners = getLearnerArray(ss);
    //check if there is an existing results summary sheet
    let existingSummarySheets = [];
    const sheets = ss.getSheets();
    for(sheet of sheets){
        if(sheet.getName().includes("Results Summary Sheet")){
            Logger.log("Results Summary Sheet already exists");
            existingSummarySheets.push(sheet);
        }
    }

    let maxSheetNumber = 0;
    if (existingSummarySheets.length > 0) {
        for (let i = 0; i < existingSummarySheets.length; i++) {
            const sheetName = existingSummarySheets[i].getName();
            const match = sheetName.match(/Results Summary Sheet (\d+)/); // Extract the number
            if (match) {
                const sheetNumber = parseInt(match[1], 10); 
                maxSheetNumber = Math.max(maxSheetNumber, sheetNumber);
            }
        }
    }

    // This will give you the next sheet number to use
    const newSheetNumber = maxSheetNumber + 1; 
    const newSheetName = `Results Summary Sheet ${newSheetNumber}`;
    Logger.log("New sheet number: " + newSheetNumber); 

    // Filter Learners who belong on the results summary sheet
    let filteredLearners = []
    for (let learner of learners){
        //Check if the learner is already in a results summary sheet
    let learnerFound = false; // Flag to track if learner is found

    for(sheet of existingSummarySheets){
        let data = sheet.getDataRange().getValues();

        // Assuming the first row (index 0) contains headers
        for (let i = 1; i < data.length; i++) { // Start from the second row
            const row = data[i];
            const nameInSheet = row[1]; // Assuming the name is in the first column (index 0)

            if (nameInSheet === learner.getName()) {
                learnerFound = true;
                break; // Exit the inner loop if learner is found
            }
        }

        if (learnerFound) {
            break; // Exit the outer loop if learner is found in any sheet
        }
    }

    if (learnerFound) {
        continue; // Skip this learner if already found in a summary sheet
    }
        let attendanceData = learner.getAttendanceRecords(ss);
        Logger.log(attendanceData)
        //if the learner has not submitted assignments, skip them
        if(attendanceData.assignmentSubmitted){
            if(!learner.lateSubmission){
                filteredLearners.push(learner);
            } else {
                //todo, handle late submission payment check
                Logger.log("Late submission for " + learner.name);
            }
            
        } else{
            continue;
        }
        
    }
    Logger.log("Filtered Learner Settings: " + learners.length);
    buildResultsSummarySheet(ss, filteredLearners, newSheetName);
}

/**
 * Creates a summary sheet in the given spreadsheet document.
 *
 * @param {string} docId - The ID of the spreadsheet document.
 * @param {Object} course - An object containing course details.
 */
function buildResultsSummarySheet(ss, learners, sheetName){
    Logger.log("creating results summary sheet");
    const summarySheet = ss.insertSheet(sheetName);
    
    //Add Document Headder
    let startRow = 1;
    summarySheet.getRange(startRow,1).setValue("Course Results Summary")
    .setFontSize(20)
    .setHorizontalAlignment("center");
    summarySheet.getRange(startRow,1,1,10).merge();
    startRow++;
  
    //Logo
  const logoUrl = "https://glinnationalcollege.ie/wp-content/uploads/2024/11/21-small.png";
  const image = SpreadsheetApp.newCellImage()
                              .setSourceUrl(logoUrl)
                              .build();
  summarySheet.getRange(startRow,1).setValue(image)
                               .setHorizontalAlignment("center");
  summarySheet.setRowHeight(startRow, 100);
  summarySheet.getRange(startRow,1, 1, 10).merge();
  startRow++;
  
  //Marks Level Tables
  let marks = [
    ["Marks", "Grade"],
    ["0-49", "Fail"],
    ["50-64", "Pass"],
    ["65-79", "Merit"],
    ["80-100", "Distinction"],
  ]
  const marksTable = summarySheet.getRange(startRow,2, 5, 2).setValues(marks);
  let borderStyle = SpreadsheetApp.BorderStyle.SOLID_THICK; // Choose your style
  let borderColor = "#4B3A71"; // GNC color
  marksTable.setHorizontalAlignment("center");
  marksTable.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
  
  //Course Details
  let courseDetails = [
    ["Start Date", learners[0].settings.startDate],
    ["Venue", ""],
    ["Group Name", ""],
    ["Module Title", learners[0].courseName],
    ["Module Code", ""],
  ]
  const courseDetailsTable = summarySheet.getRange(startRow,6, 5, 2).setValues(courseDetails);
  courseDetailsTable.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
  courseDetailsTable.setHorizontalAlignment("center");
  let detailsRow = startRow;
  for(i=0; i<5; i++){
    summarySheet.getRange(detailsRow, 7, 1, 2).merge();
    detailsRow++;
  }
  startRow += 7;
  
  //Student Results Table
  //Header
  summarySheet.getRange(startRow,1).setValue("Number");
  summarySheet.getRange(startRow,2).setValue("Name");
  summarySheet.getRange(startRow,3).setValue("");
  summarySheet.getRange(startRow,4).setValue("Project\nXX%");
  summarySheet.getRange(startRow,5).setValue("Exam\nXX%");
  summarySheet.getRange(startRow,6).setValue("Skills Demo\nXX%");
  summarySheet.getRange(startRow,7).setValue("Total");
  summarySheet.getRange(startRow,8).setValue("Grade");
  const headerRow = summarySheet.getRange(startRow,1, 1, 8);
  headerRow.setFontWeight("bold");
  headerRow.setHorizontalAlignment("center");
  headerRow.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
  summarySheet.getRange(startRow,2, 1, 2).merge();
  
  startRow++;
  
  //Student rows
  let number = 1;
  //figure out if we're using the old studentDetails or the new getLearners()
  for(learner of learners){
    summarySheet.getRange(startRow,1).setValue(number);
    summarySheet.getRange(startRow,2).setValue(learner.getName());
    let studentRange = summarySheet.getRange(startRow,1, 1, 8);
    studentRange.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
    // Check if the row number is even (0-indexed)
    if (startRow % 2 === 0) {
        studentRange.setBackground("#DDE5ED"); // Set background for the entire row
    }
    summarySheet.getRange(startRow,2, 1, 2).merge();
    // Add the formula to the Grade column (column 8)
    let gradeFormula = `=IFS(G${startRow}=0, "No Grade", G${startRow}<50,"Fail", G${startRow}<65,"Pass", G${startRow}<80,"Merit", G${startRow}<=100,"Distinction", TRUE,"Invalid Mark")`;
    summarySheet.getRange(startRow, 8).setFormula(gradeFormula);
    startRow++;
    number++;
  }
  summarySheet.getRange(startRow,1).setValue("Tutor:");
  summarySheet.getRange(startRow,2).setValue(learners[0].tutor);
  let tutorRange = summarySheet.getRange(startRow,1, 1, 8);
  tutorRange.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
  tutorRange.setFontWeight("bold");
  summarySheet.getRange(startRow,2, 1, 2).merge();
  Logger.log("Summary Sheet Created");
  }