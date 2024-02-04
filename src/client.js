function menu() {
  var ui = SpreadsheetApp.getUi();
  // Create a main menu item
  ui.createMenu('Generate Certificate')
      .addItem('Process Certs', 'processContent')
      .addItem("Process Letters", "processLetters")
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
      if(settings.emailCert === true){
        emailNewCert(pdf, student, settings);
      }
    }
  }
}