function menu() {
  var ui = SpreadsheetApp.getUi();
  // Create a main menu item
  ui.createMenu('Generate Certificate')
      .addItem('Process Certs', 'processContent')
      .addItem("Process Letters", "letters")
      .addItem("Create Labels", "labels")
      .addToUi();
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
  const docId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const content = createContent();
  const settings = getSettings(docId);
  let letterUrl;
  for(i in content){
    let student = content[i];
    if(content[i].letter == "" && content[i].paid === true && content[i].coursePassed === true){
      letterUrl = buildCompletionLetter(content[i], settings);
      linkLetter(student, letterUrl);
    }
  }
}

function processLabels(){
  generateLabels();
}
