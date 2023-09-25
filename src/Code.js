//import "google-apps-script";

function buildAttendanceSheet(course) {
  const date = new Date();
  //const course = mockClass();
  const ss = SpreadsheetApp.create(course.moduleName);
  const sheet = ss.insertSheet("CourseTitle");

  //Course Header
  sheet.getRange(1,3).setValue("Live Leaner Register "+date.getFullYear())
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
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID)
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
  let sessionNumber = 1;
  //weeks
  for(i=0; i<course.duration; i++){
    let weekNumber = i+1;
    let weekRange = sheet.getRange(3, startCol);
    let mergeRange = sheet.getRange(3, startCol, 1, course.sessionsPerWeek+1);    
    weekRange.setValue("Week "+weekNumber)
             .setBackground("#0073DB")
             .setHorizontalAlignment("center")
             .setFontColor("#FFFFFF")
             .setFontWeight("bold")
             .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    mergeRange.merge();
    for(j=0; j<course.sessionsPerWeek; j++){
      sheet.getRange(4, startCol).setValue("Date/Time")
                                   .setBackground("#8EE4F3")
                                   .setHorizontalAlignment("center")
                                   .setFontColor("#4B3A71")
                                   .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(6, startCol).setValue("Session "+sessionNumber)
                                   .setBackground("#8EE4F3")
                                   .setHorizontalAlignment("center")
                                   .setFontColor("#4B3A71")
                                   .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
      startCol++;
      sessionNumber++;
    }
    sheet.getRange(4, startCol).setValue("Tutor Notes")
                               .setBackground("#FFFFFF")
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(4, startCol, 3).merge();
    startCol++;
  }
  //Add Learner Details
  let studentRow = 7;
  //copy row of check boxes
  let checkBoxes = sheet.getRange(studentRow, 4, 1, sheet.getLastColumn()).getValues();
  for(student of course.studentDetails){
    //paste first student row
    let newCheckBoxRange = sheet.getRange(studentRow, 4,1,sheet.getLastColumn());
    newCheckBoxRange.setValues(checkBoxes);
    let studentRange = sheet.getRange(studentRow, 1).setValue(student.name);
    sheet.getRange(studentRow, 2).setValue(student.email);
    Logger.log("Adding "+student.name+" To "+studentRange.getA1Notation());
    sheet.getRange(studentRow, 3).insertCheckboxes();
    for(i=4; i<sheet.getLastColumn(); i++){
      let test = sheet.getRange(6, i).getValues();
      if(test[0][0].toString().includes("Session")){
        sheet.getRange(studentRow, i).insertCheckboxes();
      }
    }
    studentRow++
  }
  //course footer
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
  //Clean Up
  ss.moveActiveSheet(0);
  const destinationFolderId = "1fv7VcfjvOrfw7EmXPowwsH5XGFPJTIs_";
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  const opSheet = DriveApp.getFileById(ss.getId());
  const opSheetUrl = opSheet.getUrl();
  opSheet.moveTo(destinationFolder);
  return opSheetUrl;
}

function extractName(input) {
  // Split the input into lines
  var lines = input.split('\n');

  // Initialize the name variable
  var name = "Name not found";

  // Loop through the lines to find the name
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();

    // Check if the line is not empty and does not start with ###
    if (line && !line.startsWith("###")) {
      name = line;
      break; // Exit the loop after finding the name
    }
  }

  return name;
}

function extractEmail(input) {
  // Split the input into lines
  var lines = input.split('\n');

  // Initialize the email variable
  var email = "Email not found";

  // Loop through the lines to find the email
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();

    // Check if the line contains "@" to identify an email
    if (line.includes("@")) {
      email = line;
      break; // Exit the loop after finding the email
    }
  }

  return email;
}

function buildCourse(ssId){
  class CourseDetails {
  constructor(moduleName, duration, deliveryMode, sessionsPerWeek, tutorName, studentDetails) {
      this.moduleName = moduleName;
      this.duration = duration;
      this.sessionsPerWeek = sessionsPerWeek;
      this.tutorName = tutorName;
      this.studentDetails = studentDetails;
      this.deliveryMode = deliveryMode
    }
    totalSessions(){
      total = this.sessionsPerWeek*this.duration;
      return total;
    }
  }
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName("Main");
  Logger.log(sheet)
  const data = sheet.getDataRange().getValues();
  //missing variables
  const duration = 4;
  const sessionsPerWeek = 2;
  //find indexes for required fields
  let courseIndex, participantIndex, locationIndex, tutorIndex, firstNameIndex;
  for(i in data[0]){
    if(data[0][i].includes("Course")){courseIndex = Number(i)}
    if(data[0][i].includes("Participants (details)")){participantIndex = Number(i)}
    if(data[0][i].includes("Location")){locationIndex = Number(i)}
    if(data[0][i].includes("Tutor")){tutorIndex = Number(i)}
    if(data[0][i].includes("First name (participant)")){firstNameIndex = Number(i)}
    if(data[0][i].includes("Last name (participant)")){lastNameIndex = Number(i)}
    if(data[0][i].includes("Email address (participant)")){emailIndex = Number(i)}
  }
  data.shift(); //drop header row
  //find all courses on this date
  const courseNames = [];
  const courses = [];
  for(row of data){
    if(courseNames.indexOf(row[courseIndex]) == -1){
      Logger.log("New course "+row[courseIndex]+" Being Created")
      let student = {
        name: row[firstNameIndex]+" "+row[lastNameIndex],
        email: row[emailIndex]
      };
      courseNames.push(row[courseIndex]);
      let course = new CourseDetails(
        row[courseIndex],
        4,
        row[locationIndex],
        2,
        row[tutorIndex],
        [student]
      );
      courses.push(course);
    }
    else{
      let course;
      for(line of courses){
        Logger.log("Checking "+line.moduleName);
        if(line.moduleName == row[courseIndex]){
          course = line;
          student = {
            name: row[firstNameIndex]+" "+row[lastNameIndex],
            email: row[emailIndex]
          };
          course.studentDetails.push(student);
        }
      }
    }
  }
  return(courses);
}

function convertExcelToGoogleSheets(xlsId) {
  let file = DriveApp.getFileById(xlsId);
  let blob = file.getBlob();
  let folder = "16L0EUJ4KPhnu9TfTqMURR4AuKLfDiUeW";
  let config = {
    title: "[Google Sheets] " + file.getName(),
    parents: [{id: folder}],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  let spreadsheet = Drive.Files.insert(config, blob);
  return spreadsheet.id;
}

function processUpload(e){
  Logger.log(e.values);
  const xlsUrl = e.values[2];
  const email = e.values[1];
  const xlsId = xlsUrl.substring(33, xlsUrl.length);
  const ssId = convertExcelToGoogleSheets(xlsId);
  let courses = buildCourse(ssId);
  let opSheets = [];
  for(course of courses){
    let opSheet = buildAttendanceSheet(course);
    let opCourse = {
      course: course,
      sheet: opSheet
    };
    opSheets.push(opCourse);
  }
  emailAttendanceSheets(email, opSheets);
}

function emailAttendanceSheets(email, opSheets){
  Logger.log("Emailing Results to "+ email);
  Logger.log(opSheets);
  let template = HtmlService.createTemplateFromFile("emailBody");
  const urls = [];
  const locations = [];
  const tutor = [];
  const courseName = [];
  for(i in opSheets){
    urls.push(opSheets[i].sheet);
    locations.push(opSheets[i].course.deliveryMode);
    tutor.push(opSheets[i].course.tutorName);
    courseName.push(opSheets[i].course.moduleName);
  }
  Logger.log(urls);
  const messageContent = {
    urls: urls,
    locations: locations,
    tutor: tutor,
    courseName: courseName
  }
  Logger.log(messageContent)
  template.messageContent = messageContent;
  const message = template.evaluate().getContent();
  const mail = {
    to: email,
    replyTo: "info@ncultd.ie",
    subject: "Upcoming Course Attendance Sheets",
    htmlBody: message
  }
  MailApp.sendEmail(mail);
}