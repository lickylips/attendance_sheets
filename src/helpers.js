function getSettings(docId){
    let ss;
    if(docId == null){
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    else{
      ss = SpreadsheetApp.openById(docId);
    }
    const settingsSheet = ss.getSheetByName("Settings");
    const settingsArray = settingsSheet.getDataRange().getValues();
    const settings = {};
    for(i in settingsArray){
      settings[settingsArray[i][0]] = settingsArray[i][1];
    }
    return settings;
  }

  function getStudentArray(docId){
    let ss;
    if(docId == null){
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    else{
      ss = SpreadsheetApp.openById(docId);
    }
    //check if there is a sheet called Cert Generator
    let studentsSheet;
    studentsSheet = ss.getSheetByName("Document Generator");
    if(studentsSheet == null){
      //create sheet
      Logger.log("Document Generator Not Present")
      studentsSheet = ss.getSheetByName("Cert Generator");
      studentsSheet.setName("Document Generator");
    }
    return studentsSheet;
  }

  function getDocumentGeneratorSheet(ss){
    if(!ss){
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    let studentsSheet;
    studentsSheet = ss.getSheetByName("Document Generator");
    if(studentsSheet == null){
      //create sheet
      Logger.log("Document Generator Not Present")
      studentsSheet = ss.getSheetByName("Cert Generator");
      studentsSheet.setName("Document Generator");
    }
    return studentsSheet;
  }

  function getOrCreateCertsFolder(parentFolderId) {
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    var existingCerts = parentFolder.getFoldersByName("Certs");
  
    if (existingCerts.hasNext()) {
      // "Certs" folder exists, return its ID
      return existingCerts.next().getId(); 
    } else {
      // "Certs" folder doesn't exist, create it and return the ID
      var newCertsFolder = parentFolder.createFolder("Certs");
      return newCertsFolder.getId();
    }
  }

  function findOrCreateLearnersFolder(parentFolderId, learnerName) {
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    var existingFolders = parentFolder.getFoldersByName(learnerName);
  
    if (existingFolders.hasNext()) {
      // Learner folder exists, return its ID
      return existingFolders.next().getId();
    } else {
      // Learner folder doesn't exist, create it and return the ID
      var newLearnerFolder = parentFolder.createFolder(learnerName);
      newLearnerFolder.createFolder("Skills Demo");
      newLearnerFolder.createFolder("Assignment");
      newLearnerFolder.createFolder("Exam");
      return newLearnerFolder.getId();
    }
  }
  
  /**
   * getOrCreateDatedFolder
   * Funcrtion to find if there is a dated folder for the start
   * date of a given course in the cours named folder
   * @param {object} parentFolder 
   * @param {object} folderDate 
   * @returns {object} folderName
   */
  function getOrCreateDatedFolder(parentFolder, folderDate) {
    // Format the date as YYYY-MM-DD
    const formattedDate = Utilities.formatDate(folderDate, "GMT", "yyyy-MM-dd");
  
    // Get folders within the parent folder matching the date pattern
    const folders = parentFolder.getFoldersByName(formattedDate);
  
    // If a folder already exists, return it
    if (folders.hasNext()) {
      Logger.log("Found Folder")
      return folders.next();
    } else {
      // If the folder doesn't exist, create it and return it
      Logger.log("Creating Folder")
      return parentFolder.createFolder(formattedDate);
    }
  }

  function splitName(fullName) {
    Logger.log("Splitting Name "+fullName);
    if (!fullName || typeof fullName !== 'string') {
      return ['', '']; // Return empty strings if input is invalid
    }
  
    const nameParts = fullName.trim().split(/\s+/); // Split on one or more spaces
  
    if (nameParts.length === 1) {
      return [nameParts[0], '']; // If only one name, consider it the first name
    }
  
    const firstName = nameParts.slice(0, -1).join(' '); // Join everything except the last part as first name
    const lastName = nameParts.slice(-1)[0]; // Last part is the last name
  
    return [firstName, lastName];
  }

  function testSplitName(){
    const fullName = "Emily May Keegan";
    const nameParts = splitName(fullName);
    Logger.log(nameParts);
  }

  function addressStringBuilder(address1, address2, city){
    let address = "";
    try{
      address = address1+"\n"+address2+"\n"+city;
    }
    catch(e){
      address = "";
    } 
    return address;
  }

  function phoneStringBuilder(numbers){
    try{
      for(number of numbers){
        phone+=number.type+":"+number.number+"\n";
      }
    }
    catch(e) {phone = "";}
    return phone;
  }

  function addressSplitter(addressString) {
    const parts = addressString.split("\n");
  
    return {
      address1: parts[0] || "", 
      address2: parts[1] || "", 
      city: parts[2] || ""
    };
  }

  function buildUniqueKey(key1, key2){
    return key1+"-" + key2;
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
  
  function getScriptTimeZone() {
    const now = new Date();
    const offsetMinutes = now.getTimezoneOffset();
    const offsetHours = offsetMinutes / 60;
    const timeZone = "GMT" + (offsetHours >= 0 ? "+" : "") + offsetHours; // Construct GMT string
    return timeZone;
  }

  function compareTimestampsForSameDate(timestamp1, timestamp2) {
    var format = "yyyy-MM-dd"; // Format to extract only the date part
  
    var date1 = Utilities.formatDate(new Date(timestamp1), Session.getScriptTimeZone(), format);
    var date2 = Utilities.formatDate(new Date(timestamp2), Session.getScriptTimeZone(), format);
  
    return date1 === date2; 
  }

  function tickOrCross(boolian){
    if(boolian){
      return "✅";
    }
    else if (boolian == false){
      return "❌";
    }
    else{
      return "";
    }
  }

  function isSameDay(date1, date2) {
    // Check if the year, month, and day are the same for both dates
    return (
      date1.getFullYear() === date2.getFullYear() &&
      date1.getMonth() === date2.getMonth() &&
      date1.getDate() === date2.getDate()
    );
  }

  function getSpreadsheetFolder() {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  
    // Get the file object associated with the spreadsheet
    var file = DriveApp.getFileById(ss.getId());
  
    // Get the folder containing the file
    var parentFolder = file.getParents().next(); 
  
    // Return the folder ID 
    return parentFolder.getId(); 
  }

  /**
 * findTextBox
 * This function finds a textbox in a slide
 * @param {!object} slide Slide object of the service health slide
 * @param {!object} searchText the text to search for
 * @return {?object} the textbox object
 */
function findTextBox(slide, searchText) {
  const shapes = slide.getShapes();

  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    if (shape.getText().asString().includes(searchText)) {
      return shape; // Return the textbox shape when found
    }
  }

  return null; // Return null if not found
}

function checkEmail(email) {
  // Regular expression for basic email format validation
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; 

  if (emailRegex.test(email)) {
    Logger.log(email + " is a valid email format.");
    return true;
    // Do something if the email is valid
  } else {
    Logger.log(email + " is not a valid email format.");
    return false
    // Do something if the email is invalid
  }
}

function addEmail(email, emailArray) {
  // Regular expression for basic email format validation
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; 

  if (emailRegex.test(email)) { // Check if it's a valid email format
    if (!emailArray.includes(email)) { // Check if it already exists
      emailArray.push(email);
      Logger.log(email + " added to the array.");
    } else {
      Logger.log(email + " already exists in the array.");
    }
  } else {
    Logger.log(email + " is not a valid email format.");
  }
}

function downloadImage(url) {
  const response = UrlFetchApp.fetch(url);
  const blob = response.getBlob();
  return blob;
}

//get the bookeo API and Secret Keys
function getBookeoApiKeys() {
  let keys = {
      apiKey: PropertiesService.getScriptProperties().getProperty('BOOKEO_API_KEY'),
      secretKey: PropertiesService.getScriptProperties().getProperty('BOOKEO_SECRET_KEY')
  };
  return keys;
}

/**
 * Retrieves the URL of the folder containing a Google Sheet.
 *
 * @param {string} spreadsheetId The ID of the Google Sheet.
 * @return {string|null} The URL of the containing folder, or null if an error occurs.
 * @customfunction
 */
function getSpreadsheetFolderUrl(spreadsheetId) {
  try {
    // Get the file metadata for the spreadsheet
    var file = DriveApp.getFileById(spreadsheetId);
    var parents = file.getParents();

    // Check if the spreadsheet has any parents
    if (!parents.hasNext()) {
      Logger.log("Error: Spreadsheet with ID '" + spreadsheetId + "' has no parent folder.");
      return null;
    }

    // Get the parent folder (assuming only one parent)
    var parentFolder = parents.next();

    //Get the webViewLink of the parent folder
    var folderUrl = parentFolder.getUrl();

    return folderUrl;

  } catch (error) {
    Logger.log("An error occurred: " + error);
    return null;
  }
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

/**
 * Finds the row index in a Google Sheet that contains two specified values anywhere in the row.
 *
 * @param {Spreadsheet} ss The Google Sheet object.
 * @param {any} value1 The first value to search for.
 * @param {any} value2 The second value to search for.
 * @return {number} The row index (0-based) of the row containing both values, or -1 if not found. Returns -2 if the sheet is empty.
 * @customfunction
 */
function findRowWithTwoValues(sheet, value1, value2) {
  try {
    const data = sheet.getDataRange().getValues();

    if (data.length === 0) {
      Logger.log("Sheet is empty.");
      return -2; // Indicate empty sheet
    }

    for (let i = 0; i < data.length; i++) {
      if (data[i].includes(Number(value1)) && data[i].includes(Number(value2))) {
        return i; // Return the row index (0-based)
      }
    }

    return -1; // Return -1 if not found
  } catch (error) {
    Logger.log("Error in findRowWithTwoValues: " + error);
    return -1; // Return -1 in case of any error
  }
}

function getDocumentGeneratorHeaders(ss){
  let sheet = ss.getSheetByName("Document Generator");
  let data = sheet.getDataRange().getValues();
  let headers = data[0];
  let nameIndex, emailIndex, dateIndex, paidIndex, coursePassedIndex, sentIndex, sponsorIndex , letterIndex, certIndex, bookingIdIndex, personNumberIndex, addressIndex, phoneIndex, resultsIndex;
  for(i in headers){
    if(headers[i].includes("Name")){ nameIndex = Number(i);}
    if(headers[i].includes("Email")){ emailIndex = Number(i);}
    if(headers[i].includes("Date")){ dateIndex = Number(i);}
    if(headers[i].includes("Paid")){ paidIndex = Number(i);}
    if(headers[i].includes("Course Passed")){ coursePassedIndex = Number(i);}
    if(headers[i].includes("Sent")){ sentIndex = Number(i);}
    if(headers[i].includes("Sponsor Contact")){sponsorIndex = Number(i);}
    if(headers[i].includes("Letter")){letterIndex = Number(i);}
    if(headers[i].includes("Cert")){certIndex = Number(i);}
    if(headers[i].includes("Tutor")){tutorIndex = Number(i);}
    if(headers[i].includes("Booking ID")){bookingIdIndex = Number(i);}
    if(headers[i].includes("Person Number")){personNumberIndex = Number(i);}
    if(headers[i].includes("Address")){addressIndex = Number(i);}
    if(headers[i].includes("Phone")){phoneIndex = Number(i);}
    if(headers[i].includes("Results Sent")){resultsIndex = Number(i);}
  }
  let headersObj = {
    nameIndex: nameIndex,
    emailIndex: emailIndex,
    dateIndex: dateIndex,
    paidIndex: paidIndex,
    coursePassedIndex: coursePassedIndex,
    sentIndex: sentIndex,
    sponsorIndex: sponsorIndex,
    letterIndex: letterIndex,
    certIndex: certIndex,
    tutorIndex: tutorIndex,
    bookingIdIndex: bookingIdIndex,
    personNumberIndex: personNumberIndex,
    addressIndex: addressIndex,
    phoneIndex: phoneIndex,
    resultsIndex: resultsIndex
  }
  return headersObj;
}

function getResultsHeaders(ss){
  let sheets = ss.getSheets();
  let resultsSheets = []
  for(let sheet of sheets){
    if(sheet.getName().includes("Results")){
      resultsSheets.push(sheet)
    }
  }
  let sheet = resultsSheets[0]
  let data = sheet.getDataRange().getValues();

  // Find the header row
  let headerRow = -1; // Initialize to an invalid row number
  for (let i = 0; i < data.length; i++) {
      if (data[i][0] === "Number") {
      headerRow = i;
      break; 
      }
  }
  if (headerRow === -1) {
      Logger.log("Header row not found");
      return;
  }
  // Extract the data from the header row
  const headers = data[headerRow];
  //get header indexes
  const nameIndex = headers.indexOf("Name");
  const gradeIndex = headers.indexOf("Grade");
  const resultsHeaders = {
    name: nameIndex,
    grade: gradeIndex
  }
  return resultsHeaders;
}

function getAttendanceHeaders(ss){
  //find attendanceSheet Headers:
  const attendanceSheet = ss.getSheets()[0];
  const attendanceData = attendanceSheet.getDataRange().getValues();
  // find the header indexes 
  const headers = attendanceData[2];
  const headers2 = attendanceData[5];
  const bookingIdIndex = headers2.indexOf("BookingID");
  const personNumberIndex = headers2.indexOf("Person Number");
  const nameIndex = headers.indexOf("Learner Name");
  const emailIndex = headers.indexOf("Learner Email");
  const assignmentSubmittedIndex = headers.indexOf("Assignment Submitted");
  const courseCompletedIndex = headers.indexOf("Course Completed");
  const lateSubmissionIndex = headers.indexOf("Late Submission");
  let sessionStart;
  if(lateSubmissionIndex == -1){
    sessionStart = courseCompletedIndex;
  } else {
    sessionStart = lateSubmissionIndex;
  }
  const sessionHeaders = [];
  for(let i=sessionStart; i<bookingIdIndex; i++){
    if(headers[i].includes("Session")){
      let sessionHeader = {
        name: headers[i],
        number: headers[i].match(/\d+/)[0], // "1"
        presentIndex: i,
        noteIndex: i+1
      };
      sessionHeaders.push(sessionHeader);
    }
  }
  let attendanceHeaders = {
    bookingIdIndex: bookingIdIndex,
    personNumberIndex: personNumberIndex,
    nameIndex: nameIndex,
    emailIndex: emailIndex,
    assignmentSubmittedIndex: assignmentSubmittedIndex,
    courseCompletedIndex: courseCompletedIndex,
    lateSubmissionIndex: lateSubmissionIndex,
    sessionHeaders: sessionHeaders
  }
  return attendanceHeaders;
}

function getAllHeaders(ss){
  const attendanceHeaders = getAttendanceHeaders(ss);
  const resultsHeaders = getResultsHeaders(ss);
  const documentGeneratorHeaders = getDocumentGeneratorHeaders(ss);
  //Build full sheet headers object for returning
  let sheetHeaders = {
    attendanceHeaders: attendanceHeaders,
    resultsHeaders: resultsHeaders,
    documentGeneratorHeaders: documentGeneratorHeaders
  }
  return sheetHeaders;
}

function testGetHeaders(){
  const ss = SpreadsheetApp.openById("1taDg6z7Dekk7AjPyouvLZZwCIhbBeJ5GN6J2LJ5xnu4");
  const headers = getHeaders(ss);
  Logger.log(headers);
}

function findHighestRowNumbers(learnerArray, ss) {
  let highestRows = { attendance: -1, documentGenerator: -1 };

  for (const learner of learnerArray) {
    const rows = learner.getRows(ss); // Assuming `ss` is the spreadsheet object
    if (rows.attendanceSheetRow > highestRows.attendance) {
      highestRows.attendance = rows.attendanceSheetRow;
    }
    if (rows.documentGeneratorRow > highestRows.documentGenerator) {
      highestRows.documentGenerator = rows.documentGeneratorRow;
    }
  }
  if(highestRows.attendance == -1){
    let sheet = ss.getSheets()[0];
    highestRows.attendance = sheet.getLastRow()-3;
  }

  return highestRows;
}