//import "google-apps-script";

function testCreation(){
  const course = mockClass();
  buildAttendanceSheet(course);
}

function buildAttendanceSheet(course) {
  const date = new Date();
  //const course = mockClass();
  const ss = SpreadsheetApp.create(course.moduleName+" - "+course.startDate);
  const sheet = ss.insertSheet("CourseTitle");

  //Course Header
  sheet.getRange(1,1).setBackground("#4B3A71");
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
  //Clean Up
  ss.moveActiveSheet(0);
  const destinationFolderId = findDestinationFolder(course);
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  const opSheet = DriveApp.getFileById(ss.getId());
  const opSheetUrl = opSheet.getUrl();
  opSheet.moveTo(destinationFolder);
  createChainOfCustody(SpreadsheetApp.openById(opSheet.getId()), course);
  return opSheetUrl;
}

function findDestinationFolder(course) {
  const rootFolderId = "1S4OWYJNRCEev0e9IxLuQO1v6DcvQsu2i";
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const subFolders = rootFolder.getFolders();
  Logger.log("Subfolders: "+subFolders);
  let destinationFolderId;

  if (subFolders !== undefined && subFolders !== null) {
    while (subFolders.hasNext()) {
      let subFolder = subFolders.next();
      if (subFolder.getName() === course.moduleName) {
        destinationFolderId = subFolder.getId();
        break;
      }
    }
  }

  if (destinationFolderId === undefined) {
    let newFolder = rootFolder.createFolder(course.moduleName);
    destinationFolderId = newFolder.getId();
  }

  return destinationFolderId;
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
  const dataSs = SpreadsheetApp.openById("1oC8wzfx9ORiB-VEqrylhX6fOZIbL6g6fTqfmp2wd2kA");
  const dataSheet = dataSs.getSheetByName("data");
  const courseData = dataSheet.getDataRange().getValues();
  class CourseDetails {
  constructor(moduleName, deliveryMode, tutorName, studentDetails,  courseData, startDate) {
      this.moduleName = moduleName;
      this.tutorName = tutorName;
      this.studentDetails = studentDetails;
      this.deliveryMode = deliveryMode
      this.courseData = courseData
      this.startDate = startDate
    }
    courseId(){
      //get headders
      let courseId = "NA";
      for(i in this.courseData){
        if(this.courseData[i][0].trim().includes(this.moduleName.trim())){
          courseId = this.courseData[i][2];
        }
      }
      return courseId;
    }
    sessions(){
      let sessions = 4;
      for(i in this.courseData){
        if(this.courseData[i][0].trim().includes(this.moduleName.trim())){
          sessions = this.courseData[i][1];
        }
      }
      return sessions;
    }
  }
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName("Main");
  const data = sheet.getDataRange().getValues();
  //missing variables
  const duration = 4;
  const sessionsPerWeek = 2;
  //find indexes for required fields
  let courseIndex, participantIndex, locationIndex, tutorIndex, firstNameIndex, startDateIndex;
  for(i in data[0]){
    if(data[0][i].includes("Course")){courseIndex = Number(i)}
    if(data[0][i].includes("Participants (details)")){participantIndex = Number(i)}
    if(data[0][i].includes("Location")){locationIndex = Number(i)}
    if(data[0][i].includes("Tutor")){tutorIndex = Number(i)}
    if(data[0][i].includes("First name (participant)")){firstNameIndex = Number(i)}
    if(data[0][i].includes("Last name (participant)")){lastNameIndex = Number(i)}
    if(data[0][i].includes("Email address (participant)")){emailIndex = Number(i)}
    if(data[0][i].includes("Start")){startDateIndex = Number(i);}
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
      Logger.log("Adding "+student.name+" to "+row[courseIndex]);
      courseNames.push(row[courseIndex]);
      let course = new CourseDetails(
        row[courseIndex],
        row[locationIndex],
        row[tutorIndex],
        [student],
        courseData,
        row[startDateIndex]
      );
      Logger.log("Course ID: "+course.courseId());
      Logger.log("Number of Sessions: "+course.sessions())
      courses.push(course);
    }
    else{
      let course;
      for(line of courses){
        if(line.moduleName == row[courseIndex]){
          course = line;
          student = {
            name: row[firstNameIndex]+" "+row[lastNameIndex],
            email: row[emailIndex]
          };
          Logger.log("Adding "+student.name+" to "+line.moduleName)
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
  let folder = "18nt7cn0m-NZW24DERYbF46bcWu7gRZ7a";
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

function createChainOfCustody(ss, course){
  const cocSheet = ss.insertSheet("Chain of Custody");
  //Add Document Header
  
  let startRow = 2;
  //First Row / Title
  cocSheet.getRange(startRow,1).setValue("Assignment Chain of Custody")
                        .setFontSize(20)
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
  cocSheet.setRowHeight(startRow, 100);
   cocSheet.getRange(startRow,1, 1, 10).merge();
  startRow++

  //Title, Date, QQI Code
  cocSheet.getRange(startRow,1).setValue("Course Title: ");
  cocSheet.getRange(startRow,2).setValue(course.moduleName);
  cocSheet.getRange(startRow,2,1,3).merge();
  cocSheet.getRange(startRow,6).setValue("QQI Code: ");
  //TODO: Implement QQI code in course creation
  cocSheet.getRange(startRow,7).setValue(course.courseId());
  cocSheet.getRange(startRow,7,1,3).merge();

  startRow++;
  //Start date and tutor name & signature
  cocSheet.getRange(startRow,1).setValue("Start Date:");
  //TODO: Implement Start date in Course Creation
  cocSheet.getRange(startRow,2).setValue("START DATE");
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
  const borderColor = "#4B3A71";
  const borderWidth = 1;
  cocSheet.getRange(startRow,1).setValue("Number")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID)
  cocSheet.getRange(startRow,2).setValue("Student Name")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,3).setValue("Assignement Signed In")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,4).setValue("Date")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,5).setValue("Tutor Assignment Collection (Y/N)")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,6).setValue("Date")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,7).setValue("Turor Returned Assignment (Y/N)")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,8).setValue("Date")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,9).setValue("QQI Certificate Collected (Please Sign)")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  cocSheet.getRange(startRow,10).setValue("Date")
                               .setHorizontalAlignment("center")
                               .setWrap(true)
                               .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  startRow++;
  //Student rows
  for(i=0; i<course.studentDetails.length; i++){
    cocSheet.getRange(startRow, 1).setValue(i+1)
                                  .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    cocSheet.getRange(startRow, 2).setValue(course.studentDetails[i].name)
                                  .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    cocSheet.getRange(startRow,3,1,8).setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID)
    startRow++;
  }
  cocSheet.autoResizeColumn(2);
  cocSheet.getRange(1,1,cocSheet.getLastRow(),cocSheet.getLastColumn())
          .setFontColor("#4B3A71");
  cocSheet.setHiddenGridlines(true);
}

function mockClass() {
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
  const student1 = {
    name: "Seán O'Brien",
    email: "sean.obrien@ncutraining.ie"
  }
  const student2 = {
    name: "Catherine Keegan",
    email: "catherine@blah.com"
  }
  const mockStudents = [student1, student2];
  const moduleName = "Hard Knocks 101"; 
  const duration = 8;
  const sessionCount = 2;
  const tutorName = "Suzanne";
  const studentDetails = mockStudents;
  const deliveryMode = "On Site in Glin Centre"
  const course = new CourseDetails(moduleName, duration, deliveryMode, sessionCount, tutorName, studentDetails)
  return course;
}
