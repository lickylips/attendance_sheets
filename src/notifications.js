function extractEmail(input) {
    Logger.log("Extracting tutor email from "+input);
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

  function publishAttendanceSheets(opSheets){
    Logger.log("Publishing Attendance Sheets");
    const docId = "1jIZB4ywPC2CDlgSbWx7Muqbsm27Q9DoVBg4uVEvai_0";
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const date = new Date(opSheets[0].course.startDate.getTime());
    const dateParagraphIndex = findOrInsertDate(body, date);
    for(i in opSheets){
      sheet = opSheets[i];
      const courseLine = sheet.course.moduleName + " - " + sheet.course.deliveryMode;
      const courseParagraph = body.insertParagraph(dateParagraphIndex+1, courseLine);
      courseParagraph.setLinkUrl(sheet.sheet);
    }
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
      cc: "sean.obrien@ncutraining.ie",
      replyTo: "info@ncultd.ie",
      subject: "Upcoming Course Attendance Sheets",
      htmlBody: message
    }
    MailApp.sendEmail(mail);
  }

  function tutorNotificationEmail(course, url){
    Logger.log("Sending tutor notification email");
    //tutor email look up
    const tutorSheetId = "1FtKPRTCxCZSSv2vGngOJnx8n-AA6POCz2pXPAvW5jZ0";
    const tutorSS = SpreadsheetApp.openById(tutorSheetId);
    const tutorSheet = tutorSS.getSheetByName("Sheet1");
    const tutorData = tutorSheet.getDataRange().getValues();
    let primaryEmail, secondaryEmail;
    for(row of tutorData){
      if(row[0].includes(course.tutorName)  && row[0].trim() !== ""){
        Logger.log("Found Tutor "+row[0])
        if (row[2].charAt(0) === '<'){
          primaryEmail = row[2].replace(/[<>]/g, '');
        } else {
          primaryEmail = row[2];
        }
        Logger.log("Primary Email: "+primaryEmail);
        if (row[1].charAt(0) === '<'){
          secondaryEmail = row[1].replace(/[<>]/g, '');
        } else {
          secondaryEmail = row[1];
        }
        Logger.log("Secondary Email: "+secondaryEmail);
      }
    }

    let tutorEmail;
    if(primaryEmail != "" && secondaryEmail != ""){
      tutorEmail = primaryEmail+", "+secondaryEmail;
    } else if(primaryEmail != "" && secondaryEmail==""){
      tutorEmail = primaryEmail;
    } else if(secondaryEmail != ""  && primaryEmail == ""){
      tutorEmail = secondaryEmail;
    } else {
      tutorEmail = "sales@ncutraining.ie";
    }
    course.tutorEmail = tutorEmail;
    Logger.log("Tutor Emails: "+tutorEmail)
    let template = HtmlService.createTemplateFromFile("tutorEmail");
    const messageContent = {
      url: url,
      course: course
    }
    template.messageContent = messageContent;
    const message = template.evaluate().getContent();
    const mail = {
      to: tutorEmail,
      cc: "sales@ncutraining.ie",
      //cc: "sean.obrien@ncutraining.ie", //for testing
      replyTo: "info@ncutraining.ie",
      subject: "Upcoming GNC Course Details",
      htmlBody: message
    }
    MailApp.sendEmail(mail);
  }

function emailErrorLog(error){
  Logger.log("Sending error log email");
  const today = new Date();
  const email = "sean.obrien@ncutraining.ie";
  const mail = {
    to: email,
    replyTo: "info@ncultd.ie",
    subject: "Error Log "+today,
    body: error
  };
  MailApp.sendEmail(mail);
}

function emailLetter(pdf, student, settings){
  const attachment = pdf;
  let template = HtmlService.createTemplateFromFile("letterEmail");
  template.student = student;
  template.settings = settings;
  const message = template.evaluate().getContent();
  const email = {
    to: student.email,
    replyTo: "info@ncutraining.ie",
    cc: "",
    subject: "Letter of Completion",
    htmlBody: message,
    attachments: [attachment]
  }
  if(student.sponsor){email.cc = student.sponsor;}
  MailApp.sendEmail(email);
}

function emailNewCert(pdf, student, settings){
  const attachment = pdf.getBlob();
  let template = HtmlService.createTemplateFromFile("certEmail");
  template.student = student;
  template.settings = settings;
  const message = template.evaluate().getContent();
  const email = {
    to: student.email,
    replyTo: "info@ncutraining.ie",
    cc: "",
    subject: "Certificate of Completion",
    htmlBody: message,
    attachments: [attachment]
  }
  if(student.sponsor){email.cc = student.sponsor;}
  MailApp.sendEmail(email);
}

function getAttendanceRecords(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss.getId());
  const sheet = ss.getSheets()[0];
  const attendanceData = sheet.getDataRange().getValues();
  const generatorSheet = ss.getSheetByName("Document Generator");
  const generatorData = generatorSheet.getDataRange().getValues();
  //find list of students and their attendance records
  const records = [];
  for(i=6; i<attendanceData.length; i++){
    let student = {
      name: attendanceData[i][0],
      settings: settings
    }
    for(j=1; j<attendanceData[i].length; j++){
      if(attendanceData[i][j] === true || attendanceData[i][j] === false){
        let session = attendanceData[2][j];
        student[session]=attendanceData[i][j];
        student[session+"notes"]=attendanceData[i][j+1];      
      }
    }
    for(row of generatorData){
      if(row[0] === student.name){
        student.sponsor = row[2];
        student.bookingId = row[12];
        student.number = row[13];
      }
    }
    if(student.name.trim() !== ""){
      records.push(student);
    } else {
    }
  }
  return records;
}

function emailDailyAttendanceRecord(){
  const records = getAttendanceRecords();
  const sponsors = [];
  const settings = records[0].settings;
  for(student of records){
    let index = sponsors.findIndex(sponsor => sponsor.email === student.sponsor);
    if(index === -1){
      let sponsor = {
        email: student.sponsor,
        students: [student]
      };
      sponsors.push(sponsor);
    } else {
      sponsors[index].students.push(student);
    }
  }
  if(!settings.sessions){
  if(compareTimestampsForSameDate(settings.startDate, settings.endDate)){
      settings.sessions = 1;
    }
  }
  Logger.log(sponsors);
  for(sponsor of sponsors){
    let customer = getCustomerDetails(sponsor.students[0].bookingId);
    Logger.log(customer);
    sponsor.name = customer.firstName+" "+customer.lastName;
    let template = HtmlService.createTemplateFromFile("attendanceRecordEmail");
    template.content = {
      settings: settings,
      sponsor: sponsor
    };
    const message = template.evaluate().getContent();
    const mail = {
      to: sponsor.email,
      //cc: "sales@ncutraining.ie",
      cc: "sean.obrien@ncutraining.ie", //for testing
      replyTo: "info@ncutraining.ie",
      subject: "Daily Attendance Record",
      htmlBody: message
    }
    MailApp.sendEmail(mail);
  }
}