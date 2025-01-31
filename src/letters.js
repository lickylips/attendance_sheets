function buildCompletionLetter(content, settings) {
  Logger.log("generating letter");

  //Make New Letter File
  //find output folder (learner name)
  const parentFolderId = getSpreadsheetFolder();
  const learnerFolderID = findOrCreateLearnersFolder(parentFolderId, content.name);
  const outputFolder = DriveApp.getFolderById(learnerFolderID);
  const letterTemplateId = findLetterTemplate(settings.courseName);
  const letterTemplate = DriveApp.getFileById(letterTemplateId);
  let newLetter = letterTemplate.makeCopy();
  newLetter.setName("Course Completion Letter " + content.name);
  newLetter.moveTo(outputFolder);
  const newLetterId = newLetter.getId();

  //Open new Letter file as document

  newLetter = DocumentApp.openById(newLetterId);
  const body = newLetter.getBody();
  content.date.setHours(12, 0, 0, 0);//Set the date to noon to account for DST changes
  //changing date to be the day the letter is generated
  //this was requested by suzanne on 2024-08-16
  //in meeting with Chris & sales staff
  const date = new Date();
  const dateFormatted = Utilities.formatDate(date, "GMT", "EEE MMM dd yyyy");
  body.replaceText("{{STUDENT NAME}}", content.name);
  body.replaceText("{{COURSE NAME}}", settings.courseName);
  body.replaceText("{{DATE}}", dateFormatted);
  body.replaceText("{{COURSE DETAILS}}", settings.courseDetails);
  body.replaceText("{{TUTOR NAME}}", content.tutor);

  //add signature
  const signature = getSignatureImage(content.tutor);
  if (signature != null) {
    const elements = body.getParagraphs();
    for (let i = 0; i < elements.length; i++) {
      const paragraph = elements[i];
      const text = paragraph.getText();
      if (text.indexOf("{{SIGNATURE IMAGE}}") !== -1) {
        const inlineImage = paragraph.appendInlineImage(signature.blob);
        const width = 100; // Set your desired width
        const height = width * image.height / image.width;
        inlineImage.setWidth(width);
        inlineImage.setHeight(height);
        break; // Replace only the first occurrence
      }
    }
    body.replaceText("{{SIGNATURE IMAGE}}", "");
  }
  //Email letter if required
  Logger.log(settings.emailCert);
  Logger.log(content.email);
  Logger.log(content.letter);
  if (settings.emailCert && content.email && content.letter!="Letter") {
    Logger.log("emailing letter");
    //create pdf version of letter.
    newLetter.saveAndClose();
    const pdf= newLetter.getAs("application/pdf");
    emailLetter(pdf, content, settings);
  }
  
  const url = newLetter.getUrl();
  return url;
}


  function linkLetter(student, url){
    const studentSheet = getStudentArray();
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
    if(courseName.includes("Safe Pass") || courseName.includes("safepass") || courseName.includes("Safepass")){
        templateId = "19LyYJ8XPlVp6FZTQcsFOBXSbYwU964rAbeTE8CuUbcY"
    }
    else {
        templateId = "1er2HGhQ0_I3QahJIDaRRyvdFoAT8VgbWORG3wLc8ivc"
    }
    return templateId
  }

  function getSignatureImage(name) {
    // Replace "FOLDER_ID" with the actual ID of your "signatures" folder
    const folderId = "1Bs0-0nR_36bK8Uw_0vNCFFbhWj33W05e";
    
    // Search for the file with matching name
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(name + ".png");
    
    // Check if a file was found
    image = {};
    if (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();
      image.blob = file.getBlob(); // Get the image blob
      // Get image dimensions
      const imageProperties = ImgApp.getSize(image.blob);
      image.width = imageProperties.width;
      image.height = imageProperties.height;
      return image;
    } else {
      // Log a message if no file is found
      console.warn("Signature image not found for: " + name);
      return null; // Return null if no image is found
    }
  }

  function testGetSignatureImage() {
    const image = getSignatureImage("Stephen Murphy");
    Logger.log(image);
  }