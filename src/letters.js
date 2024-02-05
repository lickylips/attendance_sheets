function buildCompletionLetter(content, settings){
    Logger.log("generating letter");
    //Make New Letter File
    const outputFolder = DriveApp.getFolderById(settings.exportFolder);
    const letterTemplateId = findLetterTemplate(settings.courseName);
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
    body.replaceText("{{TUTOR NAME}}", content.tutor);
    const url = newLetter.getUrl();
    return url;
  }

  function linkLetter(student, url){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = ss.getSheetByName("Cert Generator");
    const studentArray = studentSheet.getDataRange().getValues();
    let nameCol, emailCol, linkCol;
    for(i in studentArray[0]){
      if (studentArray[0][i].includes("Name")){nameCol = Number(i);}
      if (studentArray[0][i].includes("Email")){emailCol = Number(i);}
      if (studentArray[0][i].includes("Letter")){linkCol = Number(i);}
    }
    for(i in studentArray){
      if(studentArray[i][nameCol]== student.name && studentArray[i][emailCol]==student.email){
        let linkCell = studentSheet.getRange(Number(i)+1, linkCol+1);
        linkCell.setFormula("=HYPERLINK(\""+url+"\", \"Letter\")");
      }
    }
  }

  function findLetterTemplate(courseName){
    let templateId;
    if(courseName.includes("Safe Pass") || courseName.includes("safepass")){
        templateId = "19LyYJ8XPlVp6FZTQcsFOBXSbYwU964rAbeTE8CuUbcY"
    }
    else {
        templateId = "1er2HGhQ0_I3QahJIDaRRyvdFoAT8VgbWORG3wLc8ivc"
    }
    return templateId
  }