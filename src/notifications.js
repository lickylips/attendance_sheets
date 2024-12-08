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
    let tutorEmails = getTutorEmail(course.tutorName);
    let tutorEmail = tutorEmails.primaryEmail;
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
  Logger.log("Getting Attendance Records");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss.getId());
  const sheet = ss.getSheets()[0];
  const attendanceData = sheet.getDataRange().getValues();
  const generatorSheet = ss.getSheetByName("Document Generator");
  const generatorData = generatorSheet.getDataRange().getValues();
  //find list of students and their attendance records
  const records = [];
  const sessions = {}; // Object to track attendance for each session

  for (let i = 6; i < attendanceData.length; i++) {
    if (attendanceData[i][1].includes("Additional Tutor or Sales Team Comments")) {
      break;
    }
    let student = {
      name: attendanceData[i][0],
      settings: settings,
      assignmentSubmitted: attendanceData[i][2]
    };
    for (let j = 1; j < attendanceData[i].length; j++) {
      if (attendanceData[i][j] === true || attendanceData[i][j] === false) {
        let session = attendanceData[2][j];
        student[session] = attendanceData[i][j];
        student[session + "notes"] = attendanceData[i][j + 1];
        // Initialize session attendance if not already
        if (!sessions[session]) {
          sessions[session] = false; 
        }
        // If any student attends the session, mark it as true
        if (attendanceData[i][j] === true) {
          sessions[session] = true;
        }
      }
    }
    for (row of generatorData) {
      if (row[0].trim() === student.name.trim()) {
        student.sponsor = row[2];
        student.bookingId = row[12];
        student.number = row[13];
        break;
      }
    }
    records.push(student);
  }

  // Filter out sessions where all students are marked false
  for (let i = 0; i < records.length; i++) {
    for (let session in sessions) {
      if (sessions[session] === false) {
        delete records[i][session];
        delete records[i][session + "notes"];
      }
    }
  }

  return records;
}

function emailDailyAttendanceRecord(){
  const records = getAttendanceRecords();
  Logger.log(records);
  const sponsors = [];
  const settings = records[0].settings;
  for(student of records){
    let index = sponsors.findIndex(sponsor => sponsor.email === student.sponsor);
    if(index === -1){
      let sponsor = {
        email: student.sponsor,
        students: [student]
      };
      if(sponsor.email != null){
        sponsors.push(sponsor);
      }
      
    } else {
      sponsors[index].students.push(student);
    }
  }
  if(!settings.sessions){
  if(compareTimestampsForSameDate(settings.startDate, settings.endDate)){
      settings.sessions = 1;
    }
  }
  for(sponsor of sponsors){
    Logger.log(sponsor.email);
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
      cc: "sales@ncutraining.ie",
      //cc: "sean.obrien@ncutraining.ie", //for testing
      replyTo: "sales@ncutraining.ie",
      subject: "Daily Attendance Record",
      htmlBody: message
    }
    MailApp.sendEmail(mail);
  }
}

function emailEaSubmission(settings, selectedFolder) {
  // Get the active spreadsheet's URL
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ssUrl = ss.getUrl();

  // Get the selected folder's URL
  let selectedFolderUrl = selectedFolder.getUrl();

  // Construct the email subject
  let subject = `EA Submission for ${settings.courseName}`;

  // Construct the email body
  let body = `The EA submission for <a href="${ssUrl}">${settings.courseName}</a> 
              (starting on ${settings.startDate}) has been completed 
              and is available in the following folder: <a href="${selectedFolderUrl}">${selectedFolder.getName()}</a>

              You can access the attendance spreadsheet here: <a href="${ssUrl}">${ss.getName()}</a>`;

  // Send the email
  MailApp.sendEmail({
    to: "sales@ncutraining.ie",
    subject: subject,
    htmlBody: body
  });
  Logger.log("Sales team notified by email.");
}

function emailRegForm(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formUrl = ss.getFormUrl(); 
  const emails = [];

  //add emails from doc gen sheet
  const docGenSheet = ss.getSheetByName("Document Generator");
  const docGenData = docGenSheet.getDataRange().getValues();
  let emailCol = docGenData[0].indexOf("Email");
  for(row of docGenData){
    addEmail(row[emailCol], emails);
  }

  //add emails from main sheet
  const mainSheet = ss.getSheets()[0];
  const mainSheetData = mainSheet.getDataRange().getValues();
  emailCol = mainSheetData[3].indexOf("Learner Email");
  for(row of mainSheetData){
    addEmail(row[emailCol], emails)
  }

  Logger.log("Learner Emails: "+emails);

  // Send emails to collected addresses
  const subject = "Registration Form";
  const body = `Please fill out the following registration form: ${formUrl}`;

  for (const email of emails) {
    try {
      MailApp.sendEmail(email, subject, body);
      Logger.log("Email sent to " + email);
    } catch (error) {
      Logger.log("Error sending email to " + email + ": " + error);
    }
  }
}

function emailResults(learner){
  Logger.log("sending email for "+learner.firstName+" "+learner.lastName);
  //check if results already sent
  if(learner.resultsSent){
    Logger.log("results already sent");
    return;
  }
  if(learner.email){
    const template = HtmlService.createTemplateFromFile("resultsEmail");
    template.learner = learner;
    const message = template.evaluate().getContent();
    const subject = "Results for "+learner.firstName+" "+learner.lastName;

    // Use the 'htmlBody' parameter to send HTML content
    MailApp.sendEmail({
      to: learner.email,
      cc: learner.sponsor,
      subject: subject,
      htmlBody: message 
    });
  }
}