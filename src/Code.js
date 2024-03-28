//import "google-apps-script";

function compileSheet(course){

}

function buildAttendanceSheet(course) {
  const ssTemplateId = "1cLnKKPwTuMNdA4NnwU0krGjOceXnof6pVeUZo8tS47c";
  const ssFile = DriveApp.getFileById(ssTemplateId).makeCopy(course.moduleName+" - "+course.startDate)
  const docId = ssFile.getId();
  const ss = SpreadsheetApp.openById(docId);
  createAttendanceSheet(docId, course);
  createChainOfCustody(docId, course);
  createSignInSheet(docId, course);
  createCertGenerator(docId, course);
  const oldSheet = ss.getSheetByName("Sheet1");
  ss.deleteSheet(oldSheet);
  const destinationFolderId = findDestinationFolder(course);
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  const datedFolder = getOrCreateDatedFolder(destinationFolder, course.startDate);
  const datedFolderId = datedFolder.getId();
  const certFolderId = getOrCreateCertsFolder(datedFolderId);
  createSettingsSheet(docId, course, certFolderId);
  const opSheet = DriveApp.getFileById(ss.getId());
  const opSheetUrl = opSheet.getUrl();
  opSheet.moveTo(datedFolder);
  return opSheetUrl;
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

function buildCourse(ssId){
  //get sheet with course lengths and codes
  const dataSs = SpreadsheetApp.openById("1oC8wzfx9ORiB-VEqrylhX6fOZIbL6g6fTqfmp2wd2kA");
  const dataSheet = dataSs.getSheetByName("data");
  const courseData = dataSheet.getDataRange().getValues();
  // create the course details class
  class CourseDetails {
  constructor(moduleName, deliveryMode, tutorName, studentDetails,  courseData, startDate, end) {
      this.moduleName = moduleName;
      this.tutorName = tutorName;
      this.studentDetails = studentDetails;
      this.deliveryMode = deliveryMode;
      this.courseData = courseData;
      this.startDate = startDate;
      this.end = end;
    }
    courseId(){
      //get headders
      let courseId = "NA";
      for(i in this.courseData){
        if(this.courseData[i][0].trim().includes(this.moduleName.trim())){
          courseId = this.courseData[i][2];
        }
      }
      return courseId;
    }
    sessions(){
      let sessions = 4;
      for(i in this.courseData){
        if(this.courseData[i][0].trim().includes(this.moduleName.trim())){
          sessions = this.courseData[i][1];
        }
      }
      return sessions;
    }
  }
  //Open the uploaded file to read the upcoming course info
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName("Main");
  const data = sheet.getDataRange().getValues();
  //find indexes for required fields
  let courseIndex, participantIndex, locationIndex, 
    tutorIndex, firstNameIndex, startDateIndex, 
    sponsorIndex, endIndex, address1Index, 
    address2Index, cityIndex, homePhoneIndexIndex,
    mobilePhoneIndexIndex;
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
  }
  data.shift(); //drop header row
  //find all courses on this date
  const courseNames = [];
  const courses = [];
  for(row of data){
    if(courseNames.indexOf(row[courseIndex]) == -1){
      Logger.log("New course "+row[courseIndex]+" Being Created")
      let student = {
        name: row[firstNameIndex]+" "+row[lastNameIndex],
        email: row[emailIndex],
        sponsor: row[sponsorIndex],
        address: row[address1Index]+"\n"+row[address2Index]+"\n"+row[cityIndex],
        phone: "mobile: "+row[mobilePhoneIndexIndex]+" home: "+row[homePhoneIndexIndex],
      };
      Logger.log("Adding "+student.name+" to "+row[courseIndex]);
      courseNames.push(row[courseIndex]);
      let course = new CourseDetails(
        row[courseIndex],
        row[locationIndex],
        row[tutorIndex],
        [student],
        courseData,
        row[startDateIndex],
        row[endIndex],
      );
      Logger.log("Course ID: "+course.courseId());
      Logger.log("Number of Sessions: "+course.sessions())
      Logger.log("end date: "+course.end);
      courses.push(course);
    }
    else{
      let course;
      for(line of courses){
        if(line.moduleName == row[courseIndex]){
          course = line;
          student = {
            name: row[firstNameIndex]+" "+row[lastNameIndex],
            email: row[emailIndex],
            sponsor: row[sponsorIndex],
            address: row[address1Index]+"\n"+row[address2Index]+"\n"+row[cityIndex],
            phone: "mobile: "+row[mobilePhoneIndexIndex]+" home: "+row[homePhoneIndexIndex],
          };
          Logger.log("Adding "+student.name+" to "+line.moduleName)
          course.studentDetails.push(student);
        }
      }
    }
  }
  return(courses);
}

function convertExcelToGoogleSheets(xlsId) {
  let file = DriveApp.getFileById(xlsId);
  let blob = file.getBlob();
  let folder = "18nt7cn0m-NZW24DERYbF46bcWu7gRZ7a";
  let config = {
    title: "[Google Sheets] " + file.getName(),
    parents: [{id: folder}],
    mimeType: MimeType.GOOGLE_SHEETS
  };

  // Drive API v3 method 
  let spreadsheet = Drive.Files.create(config, blob, {"convert": true});  
  return spreadsheet.id;
}

function processUpload(e){
  Logger.log(e.values);
  const xlsUrl = e.values[2];
  const email = e.values[1];
  const xlsId = xlsUrl.substring(33, xlsUrl.length);
  const ssId = convertExcelToGoogleSheets(xlsId);
  let courses = buildCourse(ssId);
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
  Logger.log(urls);
  const messageContent = {
    urls: urls,
    locations: locations,
    tutor: tutor,
    courseName: courseName
  }
  template.messageContent = messageContent;
  const message = template.evaluate().getContent();
  const mail = {
    to: email,
    replyTo: "info@ncultd.ie",
    subject: "Upcoming Course Attendance Sheets",
    htmlBody: message
  }
  MailApp.sendEmail(mail);
}

function getTutorEmail(name){
  // Convert the name to lowercase and remove any spaces or apostrophes
  const formattedName = name.toLowerCase().replace(/(\s+)|'/g, "");
  
  // Assuming email format is firstname.lastname@domain.com
  const email = formattedName + "@ncutraining.ie";
  
  return email;
}

function emailTutor(course){
  const email = getTutorEmail(course.tutorName);
  const template = HtmlService.createTemplateFromFile("tutorEmail");
  template.course = course;
  const message = template.evaluate().getContent();
  const mail = {
    to: email,
    replyTo: "info@ncultd.ie",
    subject: "Upcoming Course Attendance Sheets",
    htmlBody: message
  }
  MailApp.sendEmail(mail);
}
