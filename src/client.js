function menu() {
  var ui = SpreadsheetApp.getUi();
  var mainMenu = ui.createMenu('Attendance Sheets');

  // Cert Generation
  var submenu1 = ui.createMenu('Generate Certificates');
  submenu1.addItem('Process Certs', 'processContent');
  submenu1.addItem("Process Letters", "letters");
  submenu1.addItem("Create Labels", "labels");
  mainMenu.addSubMenu(submenu1);

  // Sending Notifications
  var submenu2 = ui.createMenu('Notifications');
  submenu2.addItem('Send Sponsor Emails', 'sponsorEmail')
  mainMenu.addSubMenu(submenu2);

  // Updates
  var submenu3 = ui.createMenu('Updates');
  submenu3.addItem('Update Learners from Bookeo', 'updateSheetFromBookeo');
  submenu3.addItem('Update Bookeo from Sheet', 'updateBookeoFromSheet');
  mainMenu.addSubMenu(submenu3);

  // EA Submission
  var submenu4 = ui.createMenu('EA Submission');
  submenu4.addItem('Submit EA', 'submitEa');
  mainMenu.addSubMenu(submenu4);
  mainMenu.addToUi();
}

function processCerts(){
  const docId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const content = createContent();
  const settings = getSettings(docId);
  let pdf;
  for(i in content){
    let student = content[i];
    if(content[i].sent === false && content[i].paid === true && content[i].coursePassed === true){
      pdf = generateCert(content[i], settings);
      markSent(student)
      linkPdf(student, pdf);
      if(settings.emailCert === true){
        emailNewCert(pdf, student, settings);
      }
    }
  }
}

function processLetters(){
  Logger.log("Processing Letters");
  const docId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const content = createContent();
  const settings = getSettings(docId);
  let letterUrl;
  for(i in content){
    let student = content[i];
    Logger.log("Checking Learner: "+content[i].name);
    Logger.log("Letter: "+content[i].letter);
    Logger.log("Course Completed: "+content[i].courseCompleted());
    Logger.log("Paid: "+content[i].paid);
    if(content[i].letter == "" && content[i].paid === true && content[i].courseCompleted() === true){
      Logger.log("Generating Letter for: "+content[i].name);
      letterUrl = buildCompletionLetter(content[i], settings);
      linkLetter(student, letterUrl);
    }
  }
}

function processLabels(){
  generateLabels();
}

function sendSponsorEmail(){
  emailDailyAttendanceRecord();
}

function updateSheetFromBookeo(){
  updateFromBookeo();
}