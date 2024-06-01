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