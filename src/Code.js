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

function buildAttendanceSheet() {
  const date = new Date();
  const course = mockClass();
  const ss = SpreadsheetApp.create(course.moduleName);
  const sheet = ss.insertSheet("CourseTitle");

  //Course Header
  sheet.getRange(1,1).setValue("Live Leaner Register "+date.getFullYear())
                     .setHorizontalAlignment("center")
                     .setBackground("#4B3A71")
                     .setFontSize(18)
                     .setFontColor("#FFFFFF")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(1,1,1,14).merge();
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
  //Course Header
  sheet.getRange(studentRow,1).setValue("Additional Tutor or Sales Team Comments")
                     .setHorizontalAlignment("center")
                     .setBackground("#4B3A71")
                     .setFontSize(18)
                     .setFontColor("#FFFFFF")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(studentRow,1,1,14).merge();
  sheet.getRange(studentRow+1,1).setValue(" ")
                     .setHorizontalAlignment("center")
                     .setBackground("#FFFFFF")
                     .setFontSize(18)
                     .setFontColor("#000000")
                     .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(studentRow+1,1,6,14).merge();
  //Clean Up
  ss.moveActiveSheet(0);
  const destinationFolderId = "1fv7VcfjvOrfw7EmXPowwsH5XGFPJTIs_";
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  const opSheet = DriveApp.getFileById(ss.getId());
  opSheet.moveTo(destinationFolder);
}
