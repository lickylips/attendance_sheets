function readSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet;
    try{
        sheet = ss.getSheetByName("Document Generator");
    } catch(err){
        Logger.log(err+" Sheet not found, trying Cert Generator");
      sheet = ss.insertSheet("Cert Generator");
    }
    const data = sheet.getDataRange().getValues();
    let headers = data[0];
    data.shift(); //remove headers from data
    let students = [];
    for(row of data){
      let student = {};
      let index = 0;
      for(header of headers){
        student[header] = row[index];
        index++;
      }
      students.push(student);
    }
    
}

function readAttendanceSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const sessionsRow = data[2];
  const numSessions = countSessions(sessionsRow);
  Logger.log(numSessions);
}

function countSessions(myArray) {
  let count = 0;
  for (let i = 0; i < myArray.length; i++) {
    if (myArray[i].toLowerCase().includes("session")) {
      count++;
    }
  }
  return count;
}
