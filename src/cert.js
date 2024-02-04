function generateCert(content, settings){
  //Get today's Date
  const today = new Date();
  const todayFormatted = Utilities.formatDate(today, "GMT", "yyyy-M-d");
  let date
  if(settings.dateType.includes("Date of Renewal")){
    date = content.renewsOn();
  }
  else{
    date = content.issuedOn();
  }
  const dateFormatted = Utilities.formatDate(date, "GMT", "MMMMM dd, yyyy")
  //Get cert template & copy
  const outputFolderId = settings.exportFolder
  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const templateId = "157JQpm3_-es0zCTpe4Zyl_fRUUkdtnI-ybKDMpnOUF8";
  
  const template = DriveApp.getFileById(templateId);
  const newCertDeck = template.makeCopy().setName(todayFormatted+" - "+content.name+" - "+settings.courseName);
  const newCertId = newCertDeck.getId();
  newCertDeck.moveTo(outputFolder)
  
  //Open new cert
  const newCert = SlidesApp.openById(newCertId);
  const slide = newCert.getSlides()[0];
  const shapes = slide.getShapes();
  for(i in shapes){
    let textBox = shapes[i];
    let text = textBox.getText().asString();
    //Find Name text box
    if(text.includes("{{NAME}}")){
      textBox.getText().setText(content.name);
    }
    if(text.includes("{{COURSE NAME}}")){
      textBox.getText().setText(settings.courseName);
    }
    if(text.includes("{{DATE TYPE}}")){
      textBox.getText().setText(settings.dateType);
    }
    if(text.includes("{{DATE}}")){
      textBox.getText().setText(dateFormatted);
    }
    if(text.includes("{{COURSE DETAILS}}")){
      textBox.getText().setText(settings.courseDetails);
    }
  }
  newCert.saveAndClose();
  //create pdf
  const pdfBlob = DriveApp.getFileById(newCertId).getBlob();
  const pdf = DriveApp.createFile(pdfBlob);
  pdf.moveTo(outputFolder);
  return pdf;
}

function getSettings(docId){
  const ss = SpreadsheetApp.openById(docId);
  const settingsSheet = ss.getSheetByName("Settings");
  const settingsArray = settingsSheet.getDataRange().getValues();
  const settings = {};
  for(i in settingsArray){
    settings[settingsArray[i][0]] = settingsArray[i][1];
  }
  return settings;
}

function readSheet(docId){
  const ss = SpreadsheetApp.openById(docId);
  const studentsSheet = ss.getSheetByName("Cert Generator");
  const studentsArray = studentsSheet.getDataRange().getValues();
  return studentsArray;
}

function buildStudentObject(studentArray, settings){
  class Student{
    constructor(name, email, date, paid, coursePassed, sent){
      this.name = name;
      this.email = email;
      this.date = date;
      this.paid = paid;
      this.coursePassed = coursePassed;
      this.sent = sent;
      this.sponsor = "";
    }
    issuedOn(){
      let issuedOnDate = new Date(this.date);
      return issuedOnDate;
    }
    renewsOn(){
      let issuedOnDate = new Date(this.date);
      let renewsOnDate = new Date(issuedOnDate.setFullYear(issuedOnDate.getFullYear()+settings.renewalDuration))
      return renewsOnDate;
    }
  }
  //get headding cols
  let headers = studentArray.shift()
  let nameCol, emailCol, dateCol, paidCol, coursePassedCol, sentCol, sponsorCol;
  for(i in headers){
    if(headers[i].includes("Name")){ nameCol = Number(i);}
    if(headers[i].includes("Email")){ emailCol = Number(i);}
    if(headers[i].includes("Date")){ dateCol = Number(i);}
    if(headers[i].includes("Paid")){ paidCol = Number(i);}
    if(headers[i].includes("Course Passed")){ coursePassedCol = Number(i);}
    if(headers[i].includes("Sent")){ sentCol = Number(i);}
    if(headers[i].includes("Sponsor Contact")){sponsorCol = Number(i);}
  }
  const students = [];
  for(i in studentArray){
    let name = studentArray[i][nameCol];
    let email = studentArray[i][emailCol];
    let date = studentArray[i][dateCol];
    let paid = studentArray[i][paidCol];
    let coursePassed = studentArray[i][coursePassedCol];
    let sent = studentArray[i][sentCol]
    let student = new Student(name, email, date, paid, coursePassed, sent);
    if(studentArray[i][sponsorCol] != null){student.sponsor = studentArray[i][sponsorCol];}
    students.push(student)
  }
  return students;
}

function createContent(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docId = ss.getId();
  const settings = getSettings(docId);
  const studentsArray = readSheet(docId);
  const content = buildStudentObject(studentsArray, settings);
  return content;
}

function markSent(student){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName("Cert Generator");
  const studentArray = studentSheet.getDataRange().getValues();
  let nameCol, emailCol, sentCol;
  for(i in studentArray[0]){
    if (studentArray[0][i].includes("Name")){nameCol = Number(i);}
    if (studentArray[0][i].includes("Email")){emailCol = Number(i);}
    if (studentArray[0][i].includes("Sent")){sentCol = Number(i);}
  }
  for(i in studentArray){
    if(studentArray[i][nameCol]== student.name && studentArray[i][emailCol]==student.email){
      let tickBoxRange = studentSheet.getRange(Number(i)+1, sentCol+1);
      tickBoxRange.setValue(true);
    }
  }
}

function getOrCreateCertsFolder(parentFolderId) {
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var existingCerts = parentFolder.getFoldersByName("Certs");

  if (existingCerts.hasNext()) {
    // "Certs" folder exists, return its ID
    return existingCerts.next().getId(); 
  } else {
    // "Certs" folder doesn't exist, create it and return the ID
    var newCertsFolder = parentFolder.createFolder("Certs");
    return newCertsFolder.getId();
  }
}

function emailNewCert(pdf, student, settings){
  const attachment = pdf.getBlob();
  let template = HtmlService.createTemplateFromFile("certEmail");
  Logger.log(student);
  template.student = student;
  template.settings = settings;
  const message = template.evaluate().getContent();
  const email = {
    to: student.email,
    replyTo: "info@ncultd.ie",
    cc: "",
    subject: "Certificate of Completion",
    htmlBody: message,
    attachments: [attachment]
  }
  if(student.sponsor){email.cc = student.sponsor;}
  MailApp.sendEmail(email);
}

function buildCompletionLetter(content, settings){
  Logger.log("generating letter");
  //Make New Letter File
  const outputFolder = DriveApp.getFolderById(settings.exportFolder);
  const letterTemplateId = "0er2HGhQ0_I3QahJIDaRRyvdFoAT8VgbWORG3wLc8ivc"
  const letterTemplate = DriveApp.getFileById(letterTemplateId);
  let newLetter = letterTemplate.makeCopy();
  newLetter.setName("Course Completion Letter "+content.name);
  newLetter.moveTo(outputFolder);
  const newLetterId = newLetter.getId();
  //Open new Letter file as document

  newLetter = DocumentApp.openById(newLetterId);
  const body = newLetter.getBody();
  const dateFormatted = Utilities.formatDate(content.date, "GMT", "EEE MMM dd yyyy");
  body.replaceText("{{STUDENT NAME}}", content.name);
  body.replaceText("{{COURSE NAME}}", settings.courseName);
  body.replaceText("{{DATE}}", dateFormatted);
  body.replaceText("{{COURSE DETAILS}}", settings.courseDetails);
}