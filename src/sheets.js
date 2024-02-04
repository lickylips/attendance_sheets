/**
 * createCertGenerator
 * Function that creates a sheet with the students required for cert generation
 * @param {string} docId 
 * @param {object} course 
 */
function createCertGenerator(docId, course){
    const ss = SpreadsheetApp.openById(docId);
    const certSheet = ss.insertSheet("Cert Generator");
    Logger.log(course.end);
    //Header row
    let headers = ["Name", "Email", "Sponsor Contact", "Date", "Paid", "Course Passed", "Sent"];
    certSheet.appendRow(headers);
    //Student Rows
    let rowNumber = 2
    for(student of course.studentDetails){
      let studentRow = [
        student.name,
        student.email,
        student.sponsor,
        course.end,
        false,
        false,
        false,
      ];
      newRow = certSheet.appendRow(studentRow);
      range = certSheet.getRange(rowNumber, 5, 1, 3);
      range.insertCheckboxes();
      rowNumber++;
    }
    ss.setFrozenRows(1);
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
  sheet.getRange(2,9).setValue(course.deliveryMode)
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
  //weeks and sessions
  let startCol = 4;
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

  //Add Learner Details
  let studentRow = 7;
  for(student of course.studentDetails){
    //paste first student row
    let newCheckBoxRange = sheet.getRange(studentRow, 4,1,sheet.getLastColumn());
    let studentRange = sheet.getRange(studentRow, 1).setValue(student.name);
    sheet.getRange(studentRow, 2).setValue(student.email);
    Logger.log("Adding "+student.name+" To "+studentRange.getA1Notation());
    sheet.getRange(studentRow, 3).insertCheckboxes();
    for(i=4; i<sheet.getLastColumn(); i++){
      let test = sheet.getRange(5, i).getValues();
      if(test[0][0].toString().includes("Present")){
        sheet.getRange(studentRow, i).insertCheckboxes();
      }
    }
    sheet.getRange(studentRow, sheet.getLastColumn()).inser
    studentRow++
  }
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
  const logoUrl = "https://lickylip.net/wp-content/uploads/2023/09/21-small.png";
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
    col+=1
    siSheet.getRange(startRow, col).setValue("Notes")
                                  .setHorizontalAlignment("center")
                                  .setWrap(true)
                                  .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    col+=1;
    sessionNo+=1
  }

  startRow++;
  //Student rows
  for(let i=0; i<course.studentDetails.length; i++){
    siSheet.getRange(startRow, 1).setValue(i+1)
                                  .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    siSheet.getRange(startRow, 2).setValue(course.studentDetails[i].name)
                                  .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    siSheet.getRange(startRow,3,1,course.sessions()*2).setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID)
    startRow++;
  }
  siSheet.autoResizeColumn(2);
  siSheet.getRange(1,1,siSheet.getLastRow(),siSheet.getLastColumn())
          .setFontColor("#4B3A71");
  siSheet.setHiddenGridlines(true);
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
  const logoUrl = "https://lickylip.net/wp-content/uploads/2023/09/21-small.png";
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
  for(i=0; i<course.studentDetails.length; i++){
    cocSheet.getRange(startRow, 1).setValue(i+1)
                                  .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
    cocSheet.getRange(startRow, 2).setValue(course.studentDetails[i].name)
                                  .setBorder(true, true, true, true, true, true, "#3B3A71", SpreadsheetApp.BorderStyle.SOLID);
    cocSheet.getRange(startRow,3,1,8).setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID)
    startRow++;
  }
  cocSheet.autoResizeColumn(1);
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
    const emailRow = ["emailCert", true, "Should the cert be emailed upon generation"];
    settingsSheet.appendRow(emailRow);
    settingsSheet.getRange("B7").insertCheckboxes();
    settingsSheet.setFrozenRows(1);
    ss.moveActiveSheet(ss.getNumSheets());
}