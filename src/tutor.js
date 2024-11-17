function getTutorEmail(tutorName){
    Logger.log("Getting tutor email for "+tutorName);
    const tutorSheetId = "1FtKPRTCxCZSSv2vGngOJnx8n-AA6POCz2pXPAvW5jZ0";
    const tutorSS = SpreadsheetApp.openById(tutorSheetId);
    const tutorSheet = tutorSS.getSheets()[0];
    const tutorData = tutorSheet.getDataRange().getValues();
    let headerRow = tutorData[0];
    let primaryEmailIndex = headerRow.indexOf("NCU Training Email");
    let secondaryEmailIndex = headerRow.indexOf("Email Address");
    let nameIndex = headerRow.indexOf("Name");
    let primaryEmail;
    for(row of tutorData){
      if(row[nameIndex].includes(tutorName)  && row[nameIndex].trim() !== ""){
        Logger.log("Found Tutor "+row[nameIndex])
        if (row[primaryEmailIndex].charAt(0) === '<'){
          primaryEmail = row[primaryEmailIndex].replace(/[<>]/g, '');
        } else {
          primaryEmail = row[primaryEmailIndex];
        }
        //Logger.log("Primary Email: "+primaryEmail);
        if (row[secondaryEmailIndex].charAt(0) === '<'){
          secondaryEmail = row[secondaryEmailIndex].replace(/[<>]/g, '');
        } else {
          secondaryEmail = row[secondaryEmailIndex];
        }
        //Logger.log("Secondary Email: "+secondaryEmail);
      }
    }
    let tutorEmail={
        primaryEmail: primaryEmail,
        secondaryEmail: secondaryEmail
    };
    return tutorEmail;
}

function testGetTutorEmail(){
  const tutorName = "Sean O'Brien";
  const tutorEmail = getTutorEmail(tutorName);
  Logger.log("Tutor Email: "+tutorEmail.primaryEmail);
  if(tutorEmail.secondaryEmail != ""){
    Logger.log("Secondary Email: "+tutorEmail.secondaryEmail);
  }
}

