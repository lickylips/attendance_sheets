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
  if (!emailArray.includes(email)) {
    emailArray.push(email);
    Logger.log(email + " added to the array.");
  } else {
    Logger.log(email + " already exists in the array.");
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

