/**
 * createCertGenerator
 * Function that creates a sheet with the students required for cert generation
 * Sheet is now called document generator
 * @param {string} docId 
 * @param {object} course 
 */
function createCertGenerator(docId, course){
    Logger.log("Creating Document Generator Sheet");
    const ss = SpreadsheetApp.openById(docId);
    const certSheet = ss.insertSheet("Document Generator");
    Logger.log("Created Document Generator Sheet :"+ certSheet.getName());
    Logger.log(course.getEnd());
    //Header row
    let headers = ["Name", "Email", "Sponsor Contact", 
                   "Tutor", "Date", "Paid", "Course Passed", 
                   "Sent", "Results Sent","Cert", "Letter",
                   "Address", "Phone", "Booking ID", "Person Number"];
    certSheet.appendRow(headers);
    //Student Rows
    let rowNumber = 2
    //figure out if we're using the old studentDetails or the new getLearners()
    if(course.studentDetails == null){
      course.studentDetails = course.getLearners();
    }
    for(student of course.studentDetails){
      let studentRow = [
        student.getName(),
        student.email,
        student.sponsor,
        course.tutorName,
        course.getEnd(),
        false,
        false,
        false,
        false,
        "",
        "",
        student.address,
        student.phone,
        student.bookingId,
      ];
      if(student.personNumber){
        studentRow.push(student.personNumber);
      }
      else{
        studentRow.push("");
      }
      newRow = certSheet.appendRow(studentRow);
      range = certSheet.getRange(rowNumber, 6, 1, 4);
      range.insertCheckboxes();
      rowNumber++;
    }
    ss.setFrozenRows(1);
    certSheet.getDataRange().setVerticalAlignment("TOP");
    ss.moveActiveSheet(ss.getNumSheets());
    //protect sheet
    let protection = certSheet.protect().setDescription("This sheet is protected. Access to sales team only");
    protection.addEditor("sales@ncutraining.ie")
    Logger.log("Finished Creating Document Generator Sheet");
  }

  /**
   * createAttendanceSheet
   * This function builds the worksheet that has the attendance record on it
   * @param {string} docId 
   * @param {object} course 
   */
  function createAttendanceSheet(docId, course){
  //open the spreadsheet for editing
  const ss = SpreadsheetApp.openById(docId);
  const sheet = ss.insertSheet(course.moduleName);

  // Add Course headder
  sheet.getRange(1,1).setBackground("#4B3A71");
  sheet.getRange(1,3).setValue("Live Leaner Register for "+course.moduleName+" on "+course.startDate)
                     .setHorizontalAlignment("center")
                     .setBackground("#4B3A71")
                     .setFontSize(18)
                     .setFontColor("#FFFFFF")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(1,2,1,14).merge();
  //tutor name
  sheet.getRange(2,1).setValue("Tutor Name")
                     .setBackground("#8EE4F3")
                     .setHorizontalAlignment("center")
                     .setFontColor("#4B3A71")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
                     //tutor name
  sheet.getRange(2,1).setValue("Tutor Name")
  .setBackground("#8EE4F3")
  .setHorizontalAlignment("center")
  .setFontColor("#4B3A71")
  .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(2,2).setValue(course.tutorName)
                     .setFontWeight("bold")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(2,2,1,6).merge();
  //mode of delivery
  sheet.getRange(2,8).setValue("Mode of Delivery")
                     .setBackground("#8EE4F3")
                     .setHorizontalAlignment("center")
                     .setFontColor("#4B3A71")
                     .setWrap(true)
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(2,9).setValue(course.getDeliveryMethod())
                     .setFontWeight("bold")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(2,9,1,6).merge();

  //Learner and session headers
  sheet.getRange(3,1).setValue("Learner Name")
                     .setBackground("#0073DB")
                     .setFontColor("#FFFFFF")
                     .setHorizontalAlignment("center")
                     .setFontWeight("bold")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setColumnWidth(1, sheet.getColumnWidth(1)*2);
  sheet.getRange(3,1,4).merge();
  sheet.getRange(3,2).setValue("Learner Email")
                     .setBackground("#0073DB")
                     .setFontColor("#FFFFFF")
                     .setHorizontalAlignment("center")
                     .setFontWeight("bold")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.setColumnWidth(2, sheet.getColumnWidth(2)*2);
  sheet.getRange(3,2,4).merge();
  sheet.getRange(3,3).setValue("Assignment Submitted")
                     .setBackground("#FFC980")
                     .setFontColor("#4B3A71")
                     .setHorizontalAlignment("center")
                     .setFontWeight("bold")
                     .setWrap(true)
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(3,3,4).merge();
  sheet.getRange(3,4).setValue("Course Completed")
                     .setBackground("#FFC980")
                     .setFontColor("#4B3A71")
                     .setHorizontalAlignment("center")
                     .setFontWeight("bold")
                     .setWrap(true)
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(3,4,4).merge();
  sheet.getRange(3,5).setValue("Late Submission")
                     .setBackground("#FFC980")
                     .setFontColor("#4B3A71")
                     .setHorizontalAlignment("center")
                     .setFontWeight("bold")
                     .setWrap(true)
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(3,5,4).merge();
  //weeks and sessions
  let startCol = 6;
  let numberOfSessions = course.sessions();
  Logger.log("Course ID: "+ course.courseId());
  Logger.log("Number of Sessions: "+numberOfSessions);
  //Sessions
  Logger.log(numberOfSessions);
  for(i=0; i<numberOfSessions; i++){
    let sessionNumber = i+1;
    let sessionRange = sheet.getRange(3, startCol);
    sessionRange.setValue("Session "+sessionNumber)
                .setBackground("#0073DB")
                .setHorizontalAlignment("center")
                .setFontColor("#FFFFFF")
                .setFontWeight("bold")
                .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    let mergeRange = sheet.getRange(3, startCol, 2, 2);
    mergeRange.merge();
    sheet.getRange(5, startCol).setValue("Present")
                               .setHorizontalAlignment("center")
                               .setBackground("#8EE4F3");
    sheet.getRange(5, startCol,2,1).merge();
    sheet.getRange(5, startCol+1).setValue("Tutor Notes")
                                 .setHorizontalAlignment("center")
                                 .setBackground("#8EE4F3");
    sheet.getRange(5, startCol+1,2,1).merge();
    startCol+=2
  }
  sheet.getRange(6, startCol).setValue("BookingID");
  sheet.getRange(6, startCol+1).setValue("Person Number");

  //Add Learner Details
  let studentRow = 7;
  //figure out if we're using the old studentDetails or the new getLearners()
  Logger.log("Student Details: "+course.studentDetails)
  if(course.studentDetails == null){
    course.studentDetails = course.getLearners();
  }
  Logger.log("Student Details: " + course.studentDetails);
  for(student of course.studentDetails){
    //paste first student row
    let newCheckBoxRange = sheet.getRange(studentRow, 4,1,sheet.getLastColumn());
    let studentRange = sheet.getRange(studentRow, 1).setValue(student.getName());
    sheet.getRange(studentRow, 2).setValue(student.email);
    sheet.getRange(studentRow, 3).insertCheckboxes();
    sheet.getRange(studentRow, 4).insertCheckboxes();
    sheet.getRange(studentRow, 5).insertCheckboxes();
    for(i=6; i<sheet.getLastColumn(); i++){
      let test = sheet.getRange(5, i).getValues();
      if(test[0][0].toString().includes("Present")){
        sheet.getRange(studentRow, i).insertCheckboxes();
      }
    }
    sheet.getRange(studentRow, startCol).setValue(student.bookingId);
    sheet.getRange(studentRow, startCol+1).setValue(student.personNumber);
    studentRow++
  }
  const lastColumn = sheet.getLastColumn();
  const formulaCell = sheet.getRange(7, lastColumn + 1); 
  formulaCell.setFormula('=ARRAYFORMULA(IF(A7:A="",,REGEXMATCH(A7:A,"(?i)"&TEXTJOIN("|",TRUE,\'Form responses 1\'!C:C)&".*(?i)"&TEXTJOIN("|",TRUE,\'Form responses 1\'!D:D))))');
  const formulaCellRange =  formulaCell.getA1Notation();
  const learnerNameRange = sheet.getRange(7, 1, sheet.getLastRow() - 6, 1); // Get all learner names starting from row 7
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('='+formulaCellRange+' = TRUE') // Formula to check the last column
    .setBackground("#00FF00") // Set background color to green
    .setRanges([learnerNameRange])
    .build();
  sheet.setConditionalFormatRules([rule]);
  //course footer
  sheet.getRange(studentRow,1).setBackground("#4B3A71");
  sheet.getRange(studentRow,2).setValue("Additional Tutor or Sales Team Comments")
                     .setHorizontalAlignment("center")
                     .setBackground("#4B3A71")
                     .setFontSize(18)
                     .setFontColor("#FFFFFF")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(studentRow,2,1,14).merge();
  sheet.getRange(studentRow+1,2).setValue(" ")
                     .setHorizontalAlignment("center")
                     .setBackground("#FFFFFF")
                     .setFontSize(18)
                     .setFontColor("#000000")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(studentRow+1,2,6,14).merge();
  sheet.setFrozenColumns(1);
  ss.moveActiveSheet(0);
}


/**
 * createChainOfCustody
 * Function to create a printable chain of custody sheet
 * @param {string} docId 
 * @param {object} course 
 */
function createChainOfCustody(docId, course){
  const ss = SpreadsheetApp.openById(docId);
  const cocSheet = ss.insertSheet("Chain of Custody");
  //Add Document Header
  
  let startRow = 1;
  //First Row / Title
  cocSheet.getRange(startRow,1).setValue("Assignment Chain of Custody")
                        .setFontSize(19)
                        .setHorizontalAlignment("center");
  cocSheet.getRange(startRow,1,1,10).merge();
  startRow++;

  //Logo
  const logoUrl = "https://glinnationalcollege.ie/wp-content/uploads/2024/11/21-small.png";
  const image = SpreadsheetApp.newCellImage()
                              .setSourceUrl(logoUrl)
                              .build();
  cocSheet.getRange(startRow,1).setValue(image)
                               .setHorizontalAlignment("center");
  cocSheet.setRowHeight(startRow, 99);
   cocSheet.getRange(startRow,1, 1, 10).merge();
  startRow++

  //Title, Date, QQI Code
  cocSheet.getRange(startRow,1).setValue("Course Title: ");
  cocSheet.getRange(startRow,2).setValue(course.moduleName);
  cocSheet.getRange(startRow,2,1,3).merge();
  cocSheet.getRange(startRow,6).setValue("QQI Code: ");
  cocSheet.getRange(startRow,7).setValue(course.courseId());
  cocSheet.getRange(startRow,7,1,3).merge();

  startRow++;
  //Start date and tutor name & signature
  cocSheet.getRange(startRow,1).setValue("Start Date:");
  Logger.log("formatting course date "+course.startDate);
  let startDate = new Date(course.startDate);
  Logger.log(startDate);
  let formattedDate = Utilities.formatDate(startDate, "GMT", "dd/MM/yyyy");
  cocSheet.getRange(startRow,2).setValue(formattedDate)
          .setHorizontalAlignment("left");
  cocSheet.getRange(startRow,2,1,3).merge();
  cocSheet.getRange(startRow,6).setValue("Tutor Signature:")
  cocSheet.getRange(startRow,7).setValue("___________________");
  cocSheet.getRange(startRow,7,1,3).merge();
  startRow++;
  
  cocSheet.getRange(startRow,7).setValue(course.moduleName);
  cocSheet.getRange(startRow,7,1,3).merge();
  startRow++;

  //Instruction on block caps
  cocSheet.getRange(startRow,1).setValue("PLEASE WRITE FULL NAME IN BLOCK CAPITALS")
          .setHorizontalAlignment("center");
  cocSheet.getRange(startRow,1,1,10).merge();
  startRow++;

  //Clear boarders first
  cocSheet.getRange(1,1,cocSheet.getLastRow(),cocSheet.getLastColumn())
          .setBorder(false, false, false, false, false, false);

  //Column Headers
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID;
  const borderColor = "#3B3A71";
  const borderWidth = 0;
  cocSheet.getRange(startRow,1).setValue("Number")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID)
  cocSheet.getRange(startRow,2).setValue("Student Name")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,3).setValue("Assignement Signed In")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,4).setValue("Date")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,5).setValue("Tutor Assignment Collection (Y/N)")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,6).setValue("Date")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,7).setValue("Turor Returned Assignment (Y/N)")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,8).setValue("Date")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,9).setValue("QQI Certificate Collected (Please Sign)")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,10).setValue("Date")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
  startRow++;
  //Student rows
  //figure out if we're using the old studentDetails or the new getLearners()
  if(course.studentDetails == null){
    course.studentDetails = course.getLearners();
  }
  for(i=0; i<course.studentDetails.length; i++){
    cocSheet.getRange(startRow, 1).setValue(i+1)
                                  .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
    cocSheet.getRange(startRow, 2).setValue(course.studentDetails[i].name)
                                  .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
    cocSheet.getRange(startRow,3,1,8).setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID)
    startRow++;
  }
  cocSheet.autoResizeColumn(2);
  cocSheet.getRange(1,1,cocSheet.getLastRow(),cocSheet.getLastColumn())
          .setFontColor("#3B3A71");
  cocSheet.setHiddenGridlines(true);
}

/**
 * createSettingsSheet
 * Function to put in a setting sheet as required by sheet generation
 * @param {string} docId 
 * @param {object} course 
 * @param {string} folderId 
 */
function createSettingsSheet(docId, course, folderId){
  //create sheet
  Logger.log("Creating settings sheet");
    const ss = SpreadsheetApp.openById(docId);
    const settingsSheet = ss.insertSheet("Settings");
    //headers
    const headers = ["Variable", "Value", "Note"];
    settingsSheet.appendRow(headers);
    //course name
    const nameRow = ["courseName", course.moduleName, "Full name of course"];
    settingsSheet.appendRow(nameRow);
    //date type
    const dateTypes = ["Date of Renewal", "Date of Issue"];
    const dateTypeRow = ["dateType", "Date of Issue" , "Renews on or issued on"];
    settingsSheet.appendRow(dateTypeRow);
    const rule = SpreadsheetApp.newDataValidation()
                                .requireValueInList(dateTypes)
                                .setAllowInvalid(false)
                                .build();
    settingsSheet.getRange("B3").setDataValidation(rule);
    //renewal duration
    const renewalRow=["renewalDuration", "3", "If type is renews on enter the renewal duration in years"];
    settingsSheet.appendRow(renewalRow);
    //course Dettails
    const detailsRow = ["courseDetails", "", "Details of course to appear on cert"];
    settingsSheet.appendRow(detailsRow);
    //folder details
    const folderRow = ["exportFolder", folderId, "ID Of the folder where the certs should go"];
    settingsSheet.appendRow(folderRow);
    //email cert
    const emailRow = ["emailCert", false, "Should the cert be emailed upon generation"];
    settingsSheet.appendRow(emailRow);
    settingsSheet.getRange("B7").insertCheckboxes();
    //delivery Mode 
    const deliveryModeRow = ["deliveryMode", course.getDeliveryMethod(), "Delivery mode of the cert, either Online or Printed"];
    settingsSheet.appendRow(deliveryModeRow);
    //tutor
    const tutorRow = ["tutor", course.tutorName, "Name of the tutor"];
    settingsSheet.appendRow(tutorRow);
    //start date
    const startDateRow = ["startDate", course.startDate, "Start date of the course"];
    settingsSheet.appendRow(startDateRow);
    //end date
    const endDateRow = ["endDate", course.getEnd(), "End date of the course"];
    settingsSheet.appendRow(endDateRow);
    if(course.productId){
      const productRow = ["productId", course.productId, "ID of the product this course relates to"];
      settingsSheet.appendRow(productRow);
    }
    let numSessions = course.sessions();
    if(numSessions > 0){
      const sessionsRow = ["sessions", numSessions, "Number of sessions in the course"];
      settingsSheet.appendRow(sessionsRow);
    }
    //EA Submission setting
    const eaRow = ["eaSubmission", false, "Has the course be submitted to EA"];
    settingsSheet.appendRow(eaRow);
    settingsSheet.getRange("B13").insertCheckboxes();

    //Course Code
    const courseCodeRow = ["courseCode", course.courseData, "Course Code"];
    settingsSheet.appendRow(courseCodeRow);
    
    settingsSheet.setFrozenRows(1);
    //move sheet to end
    ss.moveActiveSheet(ss.getNumSheets());
    //protect sheet
    let protection = settingsSheet.protect().setDescription("This sheet is protected. Access to sales team only");
    protection.addEditor("sales@ncutraining.ie")
    Logger.log("Created Settings Sheet");
}

/**
 * createSignInSheet
 * Function to build a printable sign in sheet for use in class
 * @param {string} docId 
 * @param {object} course 
 */
function createSignInSheet(docId, course){
  const ss = SpreadsheetApp.openById(docId);
  Logger.log("Building Printable Sign In Sheet");
  const siSheet = ss.insertSheet("Printable Sign In Sheet");
  //Add Document Header
  
  let startRow = 2;
  //First Row / Title
  siSheet.getRange(startRow,1).setValue("Sign In Sheet")
                        .setFontSize(20)
                        .setHorizontalAlignment("center");
  siSheet.getRange(startRow,1,1,10).merge();
  startRow++;

  //Logo
  const logoUrl = "https://glinnationalcollege.ie/wp-content/uploads/2024/11/21-small.png";
  const image = SpreadsheetApp.newCellImage()
                              .setSourceUrl(logoUrl)
                              .build();
  siSheet.getRange(startRow,1).setValue(image)
                               .setHorizontalAlignment("center");
  siSheet.setRowHeight(startRow, 100);
  siSheet.getRange(startRow,1, 1, 10).merge();
  startRow++

  //Title, Date, QQI Code
  siSheet.getRange(startRow,1).setValue("Course Title: ");
  siSheet.getRange(startRow,2).setValue(course.moduleName);
  siSheet.getRange(startRow,2,1,3).merge();
  siSheet.getRange(startRow,6).setValue("QQI Code: ");
  siSheet.getRange(startRow,7).setValue(course.courseId());
  siSheet.getRange(startRow,7,1,3).merge();

  startRow++;
  //Start date and tutor name & signature
  siSheet.getRange(startRow,1).setValue("Start Date:");
  let startDate = new Date(course.startDate);
  let formattedDate = Utilities.formatDate(startDate, "GMT", "dd/MM/yyyy");
  siSheet.getRange(startRow,2).setValue(formattedDate)
         .setHorizontalAlignment("left");
  siSheet.getRange(startRow,2,1,3).merge();
  siSheet.getRange(startRow,6).setValue("Tutor Signature:")
  siSheet.getRange(startRow,7).setValue("___________________");
  siSheet.getRange(startRow,7,1,3).merge();
  startRow++;
  
  siSheet.getRange(startRow,7).setValue(course.moduleName);
  siSheet.getRange(startRow,7,1,3).merge();
  startRow++;

  //Instruction on block caps
  siSheet.getRange(startRow,1).setValue("PLEASE WRITE FULL NAME IN BLOCK CAPITALS")
          .setHorizontalAlignment("center");
  siSheet.getRange(startRow,1,1,10).merge();
  startRow++;

  //Clear boarders first
  siSheet.getRange(1,1,siSheet.getLastRow(),siSheet.getLastColumn())
          .setBorder(false, false, false, false, false, false);

  //Column Headers
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID;
  const borderColor = "#4B3A71";
  const borderWidth = 1;
  siSheet.getRange(startRow,1).setValue("Number")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  siSheet.getRange(startRow,2).setValue("Student")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  Logger.log("Adding headders for "+course.sessions()+" sessions");
  let sessionNo = 1;
  let col = 3;
  for (let i=0; i<Number(course.sessions()); i++){
    Logger.log("adding session "+sessionNo);
    siSheet.getRange(startRow, col).setValue("Session "+sessionNo)
                                  .setHorizontalAlignment("center")
                                  .setWrap(true)
                                  .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    col+=1;
    sessionNo+=1;
  }

  startRow++;
  //Student rows
  //figure out if we're using the old studentDetails or the new getLearners()
  if(course.studentDetails == null){
      course.studentDetails = course.getLearners();
    }
  for(let i=0; i<course.studentDetails.length; i++){
    siSheet.getRange(startRow, 1).setValue(i+1)
                                  .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    siSheet.getRange(startRow, 2).setValue(course.studentDetails[i].name)
                                  .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    siSheet.getRange(startRow,3,1,course.sessions()).setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID)
    startRow++;
  }
  siSheet.autoResizeColumn(2);
  siSheet.getRange(1,1,siSheet.getLastRow(),siSheet.getLastColumn())
          .setFontColor("#4B3A71");
  siSheet.setHiddenGridlines(true);
}

/**
 * Creates a summary sheet in the given spreadsheet document.
 *
 * @param {string} docId - The ID of the spreadsheet document.
 * @param {Object} course - An object containing course details.
 */
function createSummarySheet(docId, course){
  Logger.log("creating summary sheet");
  const ss = SpreadsheetApp.openById(docId);
  const summarySheet = ss.insertSheet("Results Summary Sheet");
  
  //Add Document Headder
  let startRow = 1;
  summarySheet.getRange(startRow,1).setValue("Course Results Summary")
  .setFontSize(20)
  .setHorizontalAlignment("center");
  summarySheet.getRange(startRow,1,1,10).merge();
  startRow++;

  //Logo
const logoUrl = "https://lickylip.net/wp-content/uploads/2023/09/21-small.png";
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
  ["0-49", "F"],
  ["50-64", "P"],
  ["65-79", "M"],
  ["80-100", "D"],
]
const marksTable = summarySheet.getRange(startRow,2, 5, 2).setValues(marks);
let borderStyle = SpreadsheetApp.BorderStyle.SOLID_THICK; // Choose your style
let borderColor = "#4B3A71"; // GNC color
marksTable.setHorizontalAlignment("center");
marksTable.setBorder(true, true, true, true, true, true, borderColor, borderStyle);

//Course Details
let courseDetails = [
  ["Start Date", course.startDate],
  ["Venue", ""],
  ["Group Name", ""],
  ["Module Title", course.moduleName],
  ["Module Code", course.courseId()],
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
if(course.studentDetails == null){
  course.studentDetails = course.getLearners();
}
for(student of course.studentDetails){
  summarySheet.getRange(startRow,1).setValue(number);
  summarySheet.getRange(startRow,2).setValue(student.getName());
  let studentRange = summarySheet.getRange(startRow,1, 1, 8);
  studentRange.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
  // Check if the row number is even (0-indexed)
  if (startRow % 2 === 0) {
      studentRange.setBackground("#DDE5ED"); // Set background for the entire row
  }
  summarySheet.getRange(startRow,2, 1, 2).merge();
  startRow++;
  number++;
}
summarySheet.getRange(startRow,1).setValue("Tutor:");
summarySheet.getRange(startRow,2).setValue(course.tutorName);
let tutorRange = summarySheet.getRange(startRow,1, 1, 8);
tutorRange.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
tutorRange.setFontWeight("bold");
summarySheet.getRange(startRow,2, 1, 2).merge();
Logger.log("Summary Sheet Created");
}

