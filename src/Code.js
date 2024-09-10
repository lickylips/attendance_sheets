//import "google-apps-script";


function buildAttendanceSheet(course) {
  const ssTemplateId = "1cLnKKPwTuMNdA4NnwU0krGjOceXnof6pVeUZo8tS47c";
  const fileTitle = "["+course.tutorName+"] "+course.moduleName+" - "+course.startDate;
  const ssFile = DriveApp.getFileById(ssTemplateId).makeCopy(fileTitle);
  const docId = ssFile.getId();
  const ss = SpreadsheetApp.openById(docId);
  createAttendanceSheet(docId, course);
  createChainOfCustody(docId, course);
  createSignInSheet(docId, course);
  createCertGenerator(docId, course);
  createSummarySheet(docId, course);
  const oldSheet = ss.getSheetByName("Sheet1");
  ss.deleteSheet(oldSheet);
  const destinationFolderId = findDestinationFolder(course);
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  const datedFolder = getOrCreateDatedFolder(destinationFolder, course.startDate);
  const existingFile = findFileByName(datedFolder, fileTitle);
  if (existingFile) {
    // Rename the existing file with "deprecated-Do Not Use" prefix
    const newFileName = "Deprecated-Do Not Use - " + existingFile.getName();
    existingFile.setName(newFileName);
  }
  const datedFolderId = datedFolder.getId();
  const certFolderId = getOrCreateCertsFolder(datedFolderId);
  createSettingsSheet(docId, course, certFolderId);
  const opSheet = DriveApp.getFileById(ss.getId());
  const opSheetUrl = opSheet.getUrl();
  opSheet.moveTo(datedFolder);
  try{
    tutorNotificationEmail(course, opSheetUrl);
  }
  catch(err){
    Logger.log("Error sending tutor notification email");
    Logger.log(err);
  }
  return opSheetUrl;
}

function findFileByName(folder, name) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName() === name) {
      return file;
    }
  }
  return null;
}

function findDestinationFolder(course) {
  const rootFolderId = "1S4OWYJNRCEev0e9IxLuQO1v6DcvQsu2i";
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  const subFolders = rootFolder.getFolders()
  let destinationFolderId;

  if (subFolders !== undefined && subFolders !== null) {
    while (subFolders.hasNext()) {
      let subFolder = subFolders.next();
      if (subFolder.getName() === course.moduleName) {
        destinationFolderId = subFolder.getId();
        break;
      }
    }
  }

  if (destinationFolderId === undefined) {
    let newFolder = rootFolder.createFolder(course.moduleName);
    destinationFolderId = newFolder.getId();
  }
  Logger.log("Destination Folder Id: "+destinationFolderId);
  return destinationFolderId;
}

function extractName(input) {
  // Split the input into lines
  var lines = input.split('\n');

  // Initialize the name variable
  var name = "Name not found";

  // Loop through the lines to find the name
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();

    // Check if the line is not empty and does not start with ###
    if (line && !line.startsWith("###")) {
      name = line;
      break; // Exit the loop after finding the name
    }
  }

  return name;
}

function extractEmail(input) {
  Logger.log("Extracting tutor email from "+input);
  // Split the input into lines
  var lines = input.split('\n');

  // Initialize the email variable
  var email = "Email not found";

  // Loop through the lines to find the email
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();

    // Check if the line contains "@" to identify an email
    if (line.includes("@")) {
      email = line;
      break; // Exit the loop after finding the email
    }
  }

  return email;
}

function getCourseData() {
  Logger.log("Getting course data from spreadsheet/salesforce");
  const dataSs = SpreadsheetApp.openById("1oC8wzfx9ORiB-VEqrylhX6fOZIbL6g6fTqfmp2wd2kA");
  const dataSheet = dataSs.getSheetByName("data");
  return dataSheet.getDataRange().getValues();
}

function testGetCourseData(){
  const courseData = getCourseData();
  Logger.log(courseData[0])
  Logger.log(typeof courseData[0][0])

}

function buildCourse(ssId){
  Logger.log("Building courses");
  // create the course details class
  
  //Open the uploaded file to read the upcoming course info
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName("Main");
  const data = sheet.getDataRange().getValues();

  //find indexes for required fields
  let courseIndex, participantIndex, locationIndex, 
    tutorIndex, firstNameIndex, startDateIndex, 
    sponsorIndex, endIndex, address1Index, 
    address2Index, cityIndex, homePhoneIndexIndex,
    mobilePhoneIndexIndex, bookingNumberIndex;
  for(i in data[0]){
    if(data[0][i].includes("Course")){courseIndex = Number(i);}
    if(data[0][i].includes("Participants (details)")){participantIndex = Number(i);}
    if(data[0][i].includes("Location")){locationIndex = Number(i);}
    if(data[0][i].includes("Tutor")){tutorIndex = Number(i);}
    if(data[0][i].includes("First name (participant)")){firstNameIndex = Number(i);}
    if(data[0][i].includes("Last name (participant)")){lastNameIndex = Number(i);}
    if(data[0][i].includes("Email address (participant)")){emailIndex = Number(i);}
    if(data[0][i].includes("Start")){startDateIndex = Number(i);}
    if(data[0][i].includes("Email address (customer)")){sponsorIndex = Number(i);}
    if(data[0][i].includes("End")){endIndex = Number(i);}
    if(data[0][i].includes("Participant - Address 1")){address1Index = Number(i);}
    if(data[0][i].includes("Participant - Address 2")){address2Index = Number(i);}
    if(data[0][i].includes("Participant - City")){cityIndex = Number(i);}
    if(data[0][i].includes("Participant - Telephone (home)")){homePhoneIndexIndex = Number(i);}
    if(data[0][i].includes("Participant - Telephone (mobile)")){mobilePhoneIndexIndex = Number(i);}
    if(data[0][i].includes("Booking number")){bookingNumberIndex = Number(i);}
  }
  data.shift(); //drop header row
  //find all courses on this date
  let courseData = getCourseData();
  const courseKeys = [];
  const courses = [];
  for(row of data){
    Logger.log("Row: "+row[courseIndex]+" "+row[tutorIndex]);
    const courseKey = row[courseIndex] + " - " + row[tutorIndex]; // Composite key
    if(courseKeys.indexOf(courseKey) == -1){ //if course not already added
      Logger.log("New course "+row[courseIndex]+" Being Created")
      student = new StudentDetails (
        row[firstNameIndex],
        row[lastNameIndex],
        row[emailIndex],
        row[sponsorIndex],
        row[address1Index]+"\n"+row[address2Index]+"\n"+row[cityIndex],
        "mobile: "+row[mobilePhoneIndexIndex]+" home: "+row[homePhoneIndexIndex],
        row[bookingNumberIndex],
      );
      courseKeys.push(courseKey);
      let course = new CourseDetails(
        row[courseIndex],
        row[locationIndex],
        row[tutorIndex],
        [student],
        courseData,
        row[startDateIndex],
        row[endIndex],
      );
      courses.push(course);
    }
    else{//if course already added
      let course;
      for(line of courses){
        if(line.moduleName == row[courseIndex]){
          course = line;
          student = new StudentDetails (
            row[firstNameIndex],
            row[lastNameIndex],
            row[emailIndex],
            row[sponsorIndex],
            row[address1Index]+"\n"+row[address2Index]+"\n"+row[cityIndex],
            "mobile: "+row[mobilePhoneIndexIndex]+" home: "+row[homePhoneIndexIndex],
            row[bookingNumberIndex],
          );
          course.studentDetails.push(student);
        }
      }
    }
  }
  return(courses);
}

function convertExcelToGoogleSheets(xlsId) {
  Logger.log("Converting Excel to Google Sheets");
  let file = DriveApp.getFileById(xlsId);
  let blob = file.getBlob();
  let folder = "18nt7cn0m-NZW24DERYbF46bcWu7gRZ7a";
  let config = {
    title: "[Google Sheets] " + file.getName(),
    parents: [{id: folder}],
    mimeType: MimeType.GOOGLE_SHEETS
  };

  // Drive API v3 method to create and convert
  let spreadsheet = Drive.Files.create(config, blob, {"convert": true});  
  let ssId = spreadsheet.id;
  Logger.log(ssId);


  // Open the newly created Spreadsheet and set locale
  let ss = SpreadsheetApp.openById(ssId);
  ss.setSpreadsheetTimeZone("GMT"); 


  return ssId; // Return the ID of the Spreadsheet
}

function processUpload(e){
  Logger.log("Processing Upload");
  const xlsUrl = e.values[2];
  const email = e.values[1];
  const xlsId = xlsUrl.substring(33, xlsUrl.length);
  const ssId = convertExcelToGoogleSheets(xlsId);
  let courses = compileCourses(ssId);
  let opSheets = [];
  for(course of courses){
    try{
      let opSheet = buildAttendanceSheet(course);
      let opCourse = {
        course: course,
        sheet: opSheet
      };
      opSheets.push(opCourse);
    }
    catch(err){
      Logger.log(err)
    } 
  }
  emailAttendanceSheets(email, opSheets);
  publishAttendanceSheets(opSheets);
}

function triggerBuild(){
  const date = new Date();
  let courses = buildBookeoCourses(date);
  let email = "sean.obrien@ncutraining.ie, suzannefoster@ncutraining.ie, louisedunne@ncutraining.ie, jenniferknott@ncutraining.ie";
  let opSheets = [];
  for(course of courses){
    let opSheet = buildAttendanceSheet(course);
    let opCourse = {
      course: course,
      sheet: opSheet
    };
    opSheets.push(opCourse);
  }
  emailAttendanceSheets(email, opSheets);
  publishAttendanceSheets(opSheets);
}

function emailAttendanceSheets(email, opSheets){
  Logger.log("Emailing Results to "+ email);
  let template = HtmlService.createTemplateFromFile("emailBody");
  const urls = [];
  const locations = [];
  const tutor = [];
  const courseName = [];
  for(i in opSheets){
    urls.push(opSheets[i].sheet);
    locations.push(opSheets[i].course.deliveryMode);
    tutor.push(opSheets[i].course.tutorName);
    courseName.push(opSheets[i].course.moduleName);
  }
  const messageContent = {
    urls: urls,
    locations: locations,
    tutor: tutor,
    courseName: courseName
  }
  template.messageContent = messageContent;
  const message = template.evaluate().getContent();
  const today = Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy");
  const mail = {
    to: email,
    replyTo: "info@ncultd.ie",
    subject: "Upcoming Course Attendance Sheets for "+today,
    htmlBody: message
  }
  MailApp.sendEmail(mail);
}

function publishAttendanceSheets(opSheets){
  Logger.log("Publishing Attendance Sheets");
  const docId = "1jIZB4ywPC2CDlgSbWx7Muqbsm27Q9DoVBg4uVEvai_0";
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  const date = new Date(opSheets[0].course.startDate.getTime());
  const dateParagraphIndex = findOrInsertDate(body, date);
  for(i in opSheets){
    sheet = opSheets[i];
    const courseLine = sheet.course.moduleName + " - " + sheet.course.deliveryMode;
    const courseParagraph = body.insertParagraph(dateParagraphIndex+1, courseLine);
    courseParagraph.setLinkUrl(sheet.sheet);
  }
}

function findOrInsertDate(body, targetDate) {
  const paragraphs = body.getParagraphs();
  const targetDateObj = new Date(targetDate); // Create a Date object for comparison
  const targetDateString = Utilities.formatDate(targetDateObj, "GMT", "yyyy-MM-dd");

  // Check for existing paragraphs containing the target date
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraphText = paragraphs[i].getText();

    // Validate date format (yyyy-MM-dd)
    if (/^\d{4}-\d{2}-\d{2}$/.test(paragraphText)) {
      const paragraphDateObj = new Date(paragraphText);

      // If dates match, return index immediately
      if (paragraphDateObj.getTime() === targetDateObj.getTime()) {
        return i; 
      }
    } else {
      // Handle invalid date format
    }
  }

  // If the date was not found, insert a new paragraph in the correct place
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraphText = paragraphs[i].getText();

    if (/^\d{4}-\d{2}-\d{2}$/.test(paragraphText)) {
      const paragraphDateObj = new Date(paragraphText);

      if (paragraphDateObj < targetDateObj) {
        // Insert before this paragraph
        let dateParagraph = body.insertParagraph(i, targetDateString);
        dateParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
        return i;
      }
    }
  }

  // If no valid date paragraphs were found, or all existing dates are older,
  // insert at the beginning
  let dateParagraph = body.insertParagraph(1, targetDateString);
  dateParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  return 1;
}

function testDate(){
  ssid = "1745MOHpoyVZQYJwbH0EAumu2qOGnIwtN7pskaFPsqTM";
  ss = SpreadsheetApp.openById(ssid);
  sheet = ss.getSheetByName("Main");
  data = sheet.getDataRange().getValues();
  Logger.log(data);
}