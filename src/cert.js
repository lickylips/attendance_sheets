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
  //find output folder (learner name)
  const parentFolderId = getSpreadsheetFolder();
  const learnerFolderID = findOrCreateLearnersFolder(parentFolderId, content.name);
  const outputFolder = DriveApp.getFolderById(learnerFolderID);
  //Get cert template & copy
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





function buildStudentObject(studentArray, settings){
  class Student{
    constructor(name, email, date, paid, coursePassed, sent, letter, cert, tutor){
      this.name = name;
      this.email = email;
      this.date = date;
      this.paid = paid;
      this.coursePassed = coursePassed;
      this.sent = sent;
      this.sponsor = "";
      this.letter = letter;
      this.cert = cert;
      this.tutor = tutor;
    }
    issuedOn(){
      let issuedOnDate = new Date(this.date);
      issuedOnDate.setHours(12, 0, 0, 0);//Set the date to noon to account for DST changes
      return issuedOnDate;
    }
    renewsOn(){
      let issuedOnDate = new Date(this.date);
      let renewsOnDate = new Date(issuedOnDate.setFullYear(issuedOnDate.getFullYear()+settings.renewalDuration))
      renewsOnDate.setHours(12, 0, 0, 0);//Set the date to noon to account for DST changes
      return renewsOnDate;
    }
    courseCompleted(){
      //get attendance Sheet data
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const attendanceSheet = ss.getSheets()[0];
      const data = attendanceSheet.getDataRange().getValues();
      Logger.log(data[0]);
      //find course complete col
      let courseCompleteIndex = data[2].indexOf("Course Completed");
      let courseComplete = false;
      if(this.name != null){
        let studentRow = data.find(row => row[0] == this.name);
        courseComplete = studentRow[courseCompleteIndex];
      }
      return courseComplete;
    }
  }
  //get headding cols
  let headers = studentArray.shift()
  let nameCol, emailCol, dateCol, paidCol, coursePassedCol, sentCol, sponsorCol , letterCol, certCol;
  for(i in headers){
    if(headers[i].includes("Name")){ nameCol = Number(i);}
    if(headers[i].includes("Email")){ emailCol = Number(i);}
    if(headers[i].includes("Date")){ dateCol = Number(i);}
    if(headers[i].includes("Paid")){ paidCol = Number(i);}
    if(headers[i].includes("Course Passed")){ coursePassedCol = Number(i);}
    if(headers[i].includes("Sent")){ sentCol = Number(i);}
    if(headers[i].includes("Sponsor Contact")){sponsorCol = Number(i);}
    if(headers[i].includes("Letter")){letterCol = Number(i);}
    if(headers[i].includes("Cert")){certCol = Number(i);}
    if(headers[i].includes("Tutor")){tutorCol = Number(i);}
    if(headers[i].includes("Skills Demo")){skillsCol = Number(i);}
  }
  const students = [];
  for(i in studentArray){
    let name = studentArray[i][nameCol];
    let email = studentArray[i][emailCol];
    let date;
    if(settings.endDate == null){date = studentArray[i][dateCol];}
    else{date = settings.endDate;}
    let paid = studentArray[i][paidCol];
    let coursePassed = studentArray[i][coursePassedCol];
    let sent = studentArray[i][sentCol]
    let letter = studentArray[i][letterCol];
    let cert = studentArray[i][certCol];
    let tutor = studentArray[i][tutorCol];
    let student = new Student(name, email, date, paid, coursePassed, sent, letter, cert, tutor);
    if(studentArray[i][sponsorCol] != null){student.sponsor = studentArray[i][sponsorCol];}
    try{
      if(skillsCol != null){
        student.skills = studentArray[i][skillsCol];
      }
    }
    catch(e){
      student.skills = true;
    }
    students.push(student)
  }
  return students;
}

function createContent(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docId = ss.getId();
  const settings = getSettings(docId);
  const studentsSheet = getStudentArray(docId);
  const studentsArray = studentsSheet.getDataRange().getValues();
  const content = buildStudentObject(studentsArray, settings);
  return content;
}

function markSent(student, sheetName){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if(!sheetName){sheetName = "Document Generator";}
  const studentSheet = ss.getSheetByName(sheetName);
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




function linkPdf(student, pdf){
  const url = pdf.getUrl();
  const studentSheet = getStudentArray();
  const studentArray = studentSheet.getDataRange().getValues();
  let nameCol, emailCol, certCol;
  for(i in studentArray[0]){
    if (studentArray[0][i].includes("Name")){nameCol = Number(i);}
    if (studentArray[0][i].includes("Email")){emailCol = Number(i);}
    if (studentArray[0][i].includes("Cert")){certCol = Number(i);}
  }
  for(i in studentArray){
    if(studentArray[i][nameCol]== student.name && studentArray[i][emailCol]==student.email){
      let linkCell = studentSheet.getRange(Number(i)+1, certCol+1);
      linkCell.setFormula("=HYPERLINK(\""+url+"\", \"Cert\")");
    }
  }
}