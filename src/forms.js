function modifyForm(formId, inputData, spreadsheetId, folderId) {
  // 1. Open the Form and Spreadsheet
  let formTemplate = DriveApp.getFileById(formId);
  let formFile = formTemplate.makeCopy("Registration Form for " + inputData.moduleName);
  let outputFolder = DriveApp.getFolderById(folderId);
  formFile.moveTo(outputFolder);
  let newFormId = formFile.getId();
  let form = FormApp.openById(newFormId)
  let ss = SpreadsheetApp.openById(spreadsheetId);

  // 2. Create a new Section for the additional questions
  let newSection = form.addSectionHeaderItem() // Call addItem() on the form object
    .setTitle("Course Information"); 

  // 3. Add New Questions to the form (not the newSection)
  form.addListItem() // Call addItem() on the form object
    .setTitle('Course Code')
    .setChoiceValues([inputData.courseId()]) // Set the single choice value
    .setRequired(true);

  form.addListItem() // Call addItem() on the form object
    .setTitle('Course Name')
    .setChoiceValues([inputData.moduleName]) // Set the single choice value
    .setRequired(true);

  form.addListItem() // Call addItem() on the form object
    .setTitle('Course Start Date')
    .setChoiceValues([inputData.startDate]) // Set the date value
    .setRequired(true);

  form.addListItem() // Call addItem() on the form object
    .setTitle('Tutor Name')
    .setChoiceValues([inputData.tutorName]) // Set the single choice value
    .setRequired(true);

  // 4. Attach Form to Spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

  // 5. Move the Response Sheet to the Last Position
  let formResponseSheet = ss.getSheetByName("Form Responses 1");
  ss.moveActiveSheet(ss.getNumSheets()); // Move to the last position
  publishForm(form, ss, outputFolder);
  return form;
}

function testModifyForm(){
  let formId = "1YrqzKZ4x-J_anLcLGqCmwFEYNvrH5sc5dV1Kf5491Lw";
  let inputData = {
    courseId: "12345",
    moduleName: "Test Module",
    startDate: "2023-01-01",
    tutorName: "John Doe"
  };
  let spreadsheetId = "1uVuRsDy-VJCXzJ3jFuluVatKsQ0UPkFYXgHLGlAaGFk";
  let form = modifyForm(formId, inputData, spreadsheetId);
  Logger.log("Form ID: " + form.getId());
}

function publishForm(form, ss, folder) {
  // 1. Create a new Google Slides presentation
  const templateId = "1J0ZeiTgGvX0M8sHhL2hae4baKhNQHQeC444V_YlDVbo";
  const templateFile = DriveApp.getFileById(templateId);
  const presentationFile = templateFile.makeCopy();
  const presentationId = presentationFile.getId();
  presentationFile.moveTo(folder);
  const presentation = SlidesApp.openById(presentationId);
  presentation.setName("Registration Form Slide");

  // 2. Get the first slide (or create one if it doesn't exist)
  let slide = presentation.getSlides()[0];
  if (!slide) {
    slide = presentation.appendSlide();
  }

  // 3. Generate a short link to the form using TinyURL
  const longUrl = form.getPublishedUrl();
  const shortUrl = UrlFetchApp.fetch(`https://tinyurl.com/api-create.php?url=${encodeURIComponent(longUrl)}`)
  .getContentText();

  // 4. Insert the short link into the slide
  const linkTextBox = findTextBox(slide, "{{URL}}");
  linkTextBox.getText()
    .setText(shortUrl)
    .getTextStyle()
    .setLinkUrl(shortUrl);

  // 5. Generate a QR code image
  const qrCodeUrl = `https://quickchart.io/chart?cht=qr&chs=300x300&chl=${encodeURIComponent(shortUrl)}`;
  Logger.log(qrCodeUrl);
  const qrCodeBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob();

  // 6. Insert the QR code image into the slide
  slide.insertImage(qrCodeBlob, 400, 100, 300, 300);
}

function readRegistrationInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form responses 1");
  const data = sheet.getDataRange().getValues();
  //get header indexes
  const headers = data[0];
  const emailIndex = headers.indexOf("Email address");
  const firstNameIndex = headers.indexOf("First Name");
  const lastNameIndex = headers.indexOf("Surname");
  const dobIndex = headers.indexOf("Date of Birth");
  const ppsIndex = headers.indexOf("PPS Number");
  const genderIndex = headers.indexOf("Gender");
  const ceEmailIndex = headers.indexOf("CE Supervisor Email");
  const ceNameIndex = headers.indexOf("CE Supervisor Name");
  const understandingIndex = headers.indexOf("I have read, understood and I will comply with all information included on the Learner Information guidelines");
  const courseIndex = headers.indexOf("Course Code");
  const courseNameIndex = headers.indexOf("Course Name");
  const courseStartDateIndex = headers.indexOf("Course Start Date");
  const tutorIndex = headers.indexOf("Tutor Name");
  const medicalIndex = headers.indexOf("Medical Card Number");
  const phoneIndexIndex = headers.indexOf("Phone Number");
  const addressIndex = headers.indexOf("Address");
  const ceBinaryIndex = headers.indexOf("Are You A Participant in a Community Employment Scheme");
  const socialIndex = headers.indexOf("Are you in receipt of a social welfair payment?");
  // process rows
  data.shift();
  const learners = [];
  for(row of data){
    //check if row is empty
    if(row[0] == ""){
      continue;
    }
    let learner = {
      email: row[emailIndex],
      firstName: row[firstNameIndex],
      lastName: row[lastNameIndex],
      dob: row[dobIndex],
      pps: row[ppsIndex],
      gender: row[genderIndex],
      ceEmail: row[ceEmailIndex],
      ceName: row[ceNameIndex],
      understanding: row[understandingIndex],
      course: row[courseIndex],
      courseName: row[courseNameIndex],
      courseStartDate: row[courseStartDateIndex],
      tutor: row[tutorIndex],
      medical: row[medicalIndex],
      phone: row[phoneIndexIndex],
      address: row[addressIndex],
      ceBinary: row[ceBinaryIndex],
      social: row[socialIndex]
    };
    learners.push(learner);
  }
  return learners;
}

function createRegDoc(){
  //read registration info
  const learners = readRegistrationInfo();
  //create document and move to correct folder
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const file = DriveApp.getFileById(ssId);
  const folder = file.getParents().next();
  const doc = DocumentApp.create("Learner Registration");
  const docId = doc.getId();
  const docFile = DriveApp.getFileById(docId);
  docFile.moveTo(folder);
  //add content to document
  const body = doc.getBody();
  const styles = getStyles();
  const url = "https://glinnationalcollege.ie/wp-content/uploads/2024/11/21-small.png";
  const image = downloadImage(url);
  //Build the body
  //Loop through student pages
  for(student of learners){
    let logo = body.appendImage(image);
    logo.setWidth(200).setHeight(100);
    logo.setAttributes(styles.logo);
    body.appendParagraph("Learner Registration").setAttributes(styles.title);
    body.appendParagraph("Course Title: ").setAttributes(styles.header);
    body.appendParagraph(student.courseName).setAttributes(styles.text);
    body.appendParagraph("Course Code: ").setAttributes(styles.header);
    body.appendParagraph(student.courseCode).setAttributes(styles.text);
    body.appendParagraph("Course Start Date: ").setAttributes(styles.header);
    body.appendParagraph(Utilities.formatDate(new Date(student.courseStartDate), "GMT", "d MMMM yyyy")).setAttributes(styles.text);
    body.appendParagraph("Name: ").setAttributes(styles.header);
    body.appendParagraph(student.firstName+" "+student.lastName).setAttributes(styles.text);
    if(student.maidenName!= ""){
      body.appendParagraph("Maiden Name: ").setAttributes(styles.header);
      body.appendParagraph(student.maidenName).setAttributes(styles.text);
    }
    body.appendParagraph("PPS Number: ").setAttributes(styles.header);
    try {
      body.appendParagraph(student.pps).setAttributes(styles.text);
    }
    catch(e){Logger.log(e);}
    body.appendParagraph("Date of Birth: ").setAttributes(styles.header);
    body.appendParagraph(Utilities.formatDate(new Date(student.dob), "GMT", "d MMMM yyyy")).setAttributes(styles.text);
    body.appendParagraph("Gender: ").setAttributes(styles.header);
    body.appendParagraph(student.gender).setAttributes(styles.text);
    body.appendParagraph("Phone: ").setAttributes(styles.header);
    body.appendParagraph(student.phone+"").setAttributes(styles.text);
    body.appendParagraph("Email: ").setAttributes(styles.header);
    body.appendParagraph(student.email).setAttributes(styles.text);
    body.appendParagraph("Home Address: ").setAttributes(styles.header);
    body.appendParagraph(student.address).setAttributes(styles.text);
    body.appendHorizontalRule();
    if (student.social == "Yes"){
      body.appendParagraph("In receipt of Social Welfare Payment: ").setAttributes(styles.header);
      body.appendParagraph(student.social).setAttributes(styles.text);
    }
    if(student.medical){
      body.appendParagraph("Medical Card: ").setAttributes(styles.header);
      body.appendParagraph(student.medical).setAttributes(styles.text);
    }
    if(student.ceBinary == "Yes"){
      body.appendParagraph("Taking part in CE Scheme?:");
      body.appendParagraph("Yes").setAttributes(styles.text);

      body.appendParagraph("CE Supervisor: ").setAttributes(styles.header);
      body.appendParagraph(student.ceName+" ("+student.ceEmail+")").setAttributes(styles.text);
    }
    body.appendPageBreak();
  }
}

function getStyles(){
  //Set styles
  const styles = {
    text: {
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.BOLD]: false,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
      [DocumentApp.Attribute.FONT_FAMILY]: "Google Sans",
      [DocumentApp.Attribute.FOREGROUND_COLOR]: "#000000"
    },
    logo: {
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
    },
    header: {
      [DocumentApp.Attribute.FONT_SIZE]: 12,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT,
      [DocumentApp.Attribute.FONT_FAMILY]: "Google Sans",
      [DocumentApp.Attribute.FOREGROUND_COLOR]: "#000000"
    },
    title: {
      [DocumentApp.Attribute.FONT_SIZE]: 12,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER,
      [DocumentApp.Attribute.FONT_FAMILY]: "Google Sans",
      [DocumentApp.Attribute.FOREGROUND_COLOR]: "#000000"
    }
  };
  return styles;
}
