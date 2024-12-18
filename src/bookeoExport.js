function updateBookeoBookings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Remove and store the header row

  const bookingIdIndex = headers.indexOf("Booking ID");
  const nameIndex = headers.indexOf("Name");
  const emailIndex = headers.indexOf("Email");

  const apiKey = "AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11"; // Replace with your Bookeo API Key
  const secretKey = "5ajggnHkopp3KCWXnHN5BDJRYjK3oweX"; // Replace with your Bookeo Secret Key
  const apiUrl = "https://api-bookings.bookeo.com/v2";

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const bookingId = row[bookingIdIndex];
    const newName = row[nameIndex];
    const newEmail = row[emailIndex];

    // Update only if both name and email are present
    if (newName && newEmail) {
      const payload = {
        customer: {
          name: newName,
          email: newEmail
        }
      };

      const options = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
          'X-Bookeo-apiKey': apiKey,
          'X-Bookeo-secretKey': secretKey,
        },
        'payload': JSON.stringify(payload)
      };

      const response = UrlFetchApp.fetch(apiUrl + "/bookings/" + bookingId + "/change", options);

      // Error handling and logging (customize as needed)
      if (response.getResponseCode() === 200) {
        Logger.log("Booking updated successfully: " + bookingId);
      } else {
        Logger.log("Error updating booking: " + bookingId);
        Logger.log(response.getContentText());
      }
    }
  }
}

function updateBooking(bookingId, updatedData) {
  // Bookeo API Endpoint and Authentication
  const apiEndpoint = 'https://api.bookeo.com/v2/bookings/'; 
  const apiKey = 'AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11';
  const secretKey = '5ajggnHkopp3KCWXnHN5BDJRYjK3oweX';

  // Prepare the Request Payload
  const payload = updatedData;

  // Make the API Request
  const options = {
    'muteHttpExceptions': true, // Add this option
    'method': 'put', // Use the PUT method for updates
    'contentType': 'application/json',
    'headers': {
      'X-Bookeo-apiKey': apiKey,
      'X-Bookeo-secretKey': secretKey,
    },
    'payload': JSON.stringify(payload),
  };

  try {
    const response = UrlFetchApp.fetch(apiEndpoint + bookingId, options);
    // Check for any HTTP error
    if (response.getResponseCode() >= 400) {
      Logger.log('HTTP Error: ' + response.getResponseCode());
      Logger.log('Full response: ' + response.getContentText()); // Log the full response
      throw new Error('Booking update failed');
    }
    if (response.getResponseCode() === 200) {
      const updatedBooking = JSON.parse(response.getContentText());
      Logger.log('Booking updated successfully:', updatedBooking);
      return updatedBooking; // Return the updated booking data
    } else {
      Logger.log('Error updating booking:', response.getContentText());
      throw new Error('Booking update failed');
    }
  } catch (error) {
    Logger.log('Error updating booking:', error);
    throw error;
  }
}

function testUpdateBooking(){
  const bookingId = "22405213108336";
  const updatedData = {
    productId: "2226YRYY3163AB95546C",
    eventId: "2226YRYY3163AB95546C_222C4X34W18F9A7EC341_2024-06-01",
    customer: {
      id: "22243EEYY18F9A81A311",
      firstName: "Seán",
      lastName: "O'Brien",
    },
    participants: {
      "numbers": [
        {
            "peopleCategoryId": "Cadults",
            "number": 11
        }
      ],
      details: [ // Array for adding participants
        {
          personId: "PNEW",
          peopleCategoryId: "Cadults",
          personDetails: {
            firstName: "Janette",
            lastName: "Smyth",
          },
          categoryIndex: 1,
        },
        {
          personId: "PNEW",
          personDetails: {
            firstName: "Seán",
            lastName: "O'Brien",
            email: "sean.obrien@ncutraining.ie",
          },
          peopleCategoryId: "Cadults",
          categoryIndex: 2,
        }
      ]
    }
  };
  const updatedBooking = updateBooking(bookingId, updatedData);
  Logger.log(updatedBooking);
}

function getDocumentGeneratorData(ss){
  Logger.log("Getting Document Generator Data");
  //Check if Spreadsheet is passed in
  if(!ss){
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  //Find Document Generator Sheet
  sheet = getDocumentGeneratorSheet(ss);
  
}

function updateFromBookeo(){
  Logger.log("Updating Sheet from Bookeo");
  //Get Sheet info
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = ss.getSheets()[0];
  //Get student Array
  const sheet = getStudentArray();
  const studentArray = sheet.getDataRange().getValues();
  Logger.log("Finding Header Indexes");
  let headers = studentArray.shift();
  let bookings = [];
  let nameIndex, emailIndex, dateIndex, paidIndex, 
    coursePassedIndex, sentIndex, sponsorIndex, 
    letterIndex, certIndex, bookingIdIndex, tutorIndex, 
    addressIndex, phoneIndex, personNumberIndex;
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
    if(headers[i].includes("Person Number")){personNumberIndex = Number(i);}
    if(headers[i].includes("Booking ID")){bookingIdIndex = Number(i);}
    if(headers[i].includes("Address")){addressIndex = Number(i);}
    if(headers[i].includes("Phone")){phoneIndex = Number(i);}
  }

  let keys = getBookeoApiKeys();
  for(let i=0; i<studentArray.length; i++){
    Logger.log("Checking Booking for: "+studentArray[i][nameIndex]);
    let learner = studentArray[i];
    let bookeoLearner;
    let bookingDetails = bookeoLibrary.getBookingById(learner[bookingIdIndex], keys.apiKey, keys.secretKey);
    const personDetails = bookingDetails.participants.details;
    for(person of personDetails){
      if(person.categoryIndex == studentArray[i][personNumberIndex]){
        bookeoLearner = person;
        break;
      }
    }
    //Process Booking Details for name, phone and address
    let name = bookeoLearner.personDetails.firstName + " " + bookeoLearner.personDetails.lastName;
    let address = bookeoLearner.personDetails.streetAddress.address1+"\n"+bookeoLearner.personDetails.streetAddress.address2;
    let phone="";
    try{
      for(numbers of bookeoLearner.personDetails.phoneNumbers){
        phone+=numbers.type+":"+numbers.number+"\n";
      }
    }
    catch(e){
      Logger.log("No phone number found");
    }
    Logger.log("Found Bookeo Learner details for "+name);

    //Check if name is the same as the sheet name
    let row = i+2;
    let sheetName = sheet.getRange(row, nameIndex+1).getValue();
    Logger.log("Checking sheet name: "+sheetName+" against bookeo name: "+name);
    //handle if a customer has canceled
    Logger.log("Checking if learner has cancelled booking");
    Logger.log(bookingDetails.canceled);
    if(bookingDetails.canceled == true){
      Logger.log("Learner has cancelled booking");
      sheet.deleteRow(row);
    }
    if(sheetName.trim() != name.trim()){
      //Names Don't Match, Update Names
      Logger.log("Changes required for "+name);
      sheet.getRange(row, nameIndex+1).setValue(name);
      updateMainSheet(learner[bookingIdIndex], studentArray[i][personNumberIndex], name, ss);
      replaceTextInSheet("Results Summary Sheet", sheetName, name);
      replaceTextInSheet("Chain of Custody", sheetName, name);
      replaceTextInSheet("Printable Sign In Sheet", sheetName, name);
    } else {
      Logger.log("No changes required for "+name);
    }
  }
  //check if there are new learners

}

function updateMainSheet(bookingId, personNumber, name, ss){
  Logger.log("Updating Main Sheet");
  const attendanceSheet = ss.getSheets()[0];
  let headers = attendanceSheet.getRange(6,1,1,attendanceSheet.getLastColumn()).getValues();
  let bookingIdIndex = headers[0].indexOf("BookingID");
  let personNumberIndex = headers[0].indexOf("Person Number");
  let learnerData = attendanceSheet.getRange(7,1,attendanceSheet.getLastRow(),attendanceSheet.getLastColumn()).getValues();
  for(let i=0; i<learnerData.length; i++){
    Logger.log("checking bookingID: "+learnerData[i][bookingIdIndex]+" against bookingID: "+bookingId);
    if(learnerData[i][bookingIdIndex] == bookingId && learnerData[i][personNumberIndex] == personNumber){
      Logger.log("Updating Learner Name");
      attendanceSheet.getRange(i+7, 1).setValue(name);
      break;
    }
  }
}

function getCustomerDetails(bookingId){
  const apiKey = 'AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11';
  const secretKey = '5ajggnHkopp3KCWXnHN5BDJRYjK3oweX';
  const apiUrlBase = 'https://api.bookeo.com/v2/';
  const booking = bookeoLibrary.getBookingById(bookingId);
  const customer = booking.customer;
  Logger.log(customer)
  return customer;
}

function replaceTextInSheet(sheetName, searchText, replaceText) {
  // Get the spreadsheet and sheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  // Find the cell containing the search text.
  var foundRange = sheet.createTextFinder(searchText).findNext();

  // If the text is found, replace it with the new text.
  if (foundRange) {
    foundRange.setValue(replaceText);
  } else {
    Logger.log("Text not found in sheet.");
  }
}

function updateBookeoNameFromSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const headers = sheet.getRange(6,1,1,sheet.getLastColumn()).getValues();
  const bookingIdIndex = headers[0].indexOf("BookingID");
  const personNumberIndex = headers[0].indexOf("Person Number");
  const learnerData = sheet.getRange(7,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  for(let i=0 ; i<learnerData.length; i++){
    if(learnerData[i][bookingIdIndex] == null || learnerData[i][bookingIdIndex] == ""){
      break;
    }
    const bookingId = learnerData[i][bookingIdIndex];
    const personNumber = learnerData[i][personNumberIndex];
    const booking = bookeoLibrary.getBookingById(bookingId);
    const customerId = booking.customer.id;
    let participant = booking.participants.details.find(p => p.categoryIndex === personNumber);
    const participantId = participant.personId;
    const bookeoFullName = participant.personDetails.firstName + " " + participant.personDetails.lastName;
    const fullName = splitName(learnerData[i][0]);
    if(fullName[0] + " " + fullName[1] == bookeoFullName){
      Logger.log("No changes required for booking: " + bookingId);
      continue;
    }
    else{
      Logger.log("Updating Booking: " + bookingId + " with new name: " + fullName[0] + " " + fullName[1]);
      const payload = createBookeoPayload(booking, personNumber, fullName[0], fullName[1]);
      updateParticipant(customerId, participantId, payload);
      replaceTextInSheet("Results Summary Sheet", bookeoFullName, fullName[0] + " " + fullName[1]);
      replaceTextInSheet("Chain of Custody", bookeoFullName, fullName[0] + " " + fullName[1]);
      replaceTextInSheet("Printable Sign In Sheet", bookeoFullName, fullName[0] + " " + fullName[1]);
      replaceTextInSheet("Document Generator", bookeoFullName, fullName[0] + " " + fullName[1]);
    }
    
  }

}

function createBookeoPayload(booking, participantIndex, newFirstName, newLastName) {
  // Find the participant with the given categoryIndex
  let participant = booking.participants.details.find(p => p.categoryIndex === participantIndex);

  if (participant) {
    // Create the payload object with the required fields
    let payload = {
      "firstName": newFirstName,
      "lastName": newLastName,
      "emailAddress": participant.personDetails.emailAddress,
      "phoneNumbers": participant.personDetails.phoneNumbers,
      "streetAddress": participant.personDetails.streetAddress,
    };
    return payload;
  } else {
    Logger.log("Participant not found.");
    return null;
  }
}

function updateParticipant(customerId, participantId, payload) {
  // Bookeo API Endpoint and Authentication
  const apiEndpoint = 'https://api.bookeo.com/v2/customers/'; 
  const apiKey = 'AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11';
  const secretKey = '5ajggnHkopp3KCWXnHN5BDJRYjK3oweX';

  // Make the API Request
  const options = {
    'muteHttpExceptions': true, // Add this option
    'method': 'put', // Use the PUT method for updates
    'contentType': 'application/json',
    'headers': {
      'X-Bookeo-apiKey': apiKey,
      'X-Bookeo-secretKey': secretKey,
    },
    'payload': JSON.stringify(payload),
  };

  try {
    const response = UrlFetchApp.fetch(apiEndpoint + customerId+"/linkedpeople/"+participantId, options);
    Logger.log('Full response: ' + response.getContentText()); // Log the full response 
    // Check for any HTTP error
    if (response.getResponseCode() >= 400) {
      Logger.log('HTTP Error: ' + response.getResponseCode());
      Logger.log('Full response: ' + response.getContentText()); // Log the full response
      throw new Error('Booking update failed');
    }
    if (response.getResponseCode() === 200) {
      //const updatedBooking = JSON.parse(response.getContentText());
      Logger.log('Booking updated successfully');
      return "updatedBooking"; // Return the updated booking data
    } else {
      Logger.log('Error updating booking:', response.getContentText());
      throw new Error('Booking update failed');
    }
  } catch (error) {
    Logger.log('Error updating booking:', error);
    throw error;
  }
}

function findBookingsForModule(startDate, moduleName){
  const keys = getBookeoApiKeys();
  const date = new Date(startDate);
  Logger.log(date);
  let bookings = bookeoLibrary.getBookeoBookingsForDate(date, keys.apiKey, keys.secretKey);
  const moduleBookings = [];
  for (let booking of bookings.data) {
    if (booking.productName.includes(moduleName)) {
      Logger.log("Found booking("+booking.bookingNumber+") for module: " + booking.productName);
      moduleBookings.push(booking);
    }
  }
  return moduleBookings;
}

function testFindBookingsForModule(){
  const startDate = new Date("2024-09-17");
  const moduleName = "Test Course";
  const bookings = findBookingsForModule(startDate, moduleName);
  Logger.log(bookings);
}

function updateSpreadsheetFromBookeo() {
  Logger.log("Updating Spreadsheet from Bookeo");
  // Build module from sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Building module from sheet");
  const module = buildModuleFromSheet(ss);

  // Get learners existing in the sheet
  Logger.log("Getting learners from sheet");
  const sheetLearners = getLearnerArray(ss);

  // Get learners from the module object
  Logger.log("Getting learners from module");
  const moduleLearners = module.getLearners();

  // Track removed learners for logging
  const removedLearners = [];

  // Loop through sheet learners and check for removal
for (const sheetLearner of sheetLearners) {
  Logger.log(sheetLearner);
  const found = moduleLearners.some(moduleLearner => {
    // Convert both to numbers for strict comparison
    const sheetBookingId = Number(sheetLearner.bookingId);
    const sheetPersonNumber = Number(sheetLearner.personNumber);
    const moduleBookingId = Number(moduleLearner.bookingId);
    const modulePersonNumber = Number(moduleLearner.personNumber);
    return moduleBookingId === sheetBookingId && modulePersonNumber === sheetPersonNumber;
  });

  if (!found) {
    // Learner not found in module, remove the row
    const row = sheetLearner.getRows(ss);
    Logger.log(row);
    ss.getSheets()[0].deleteRow(row.attendanceSheetRow+1);
    ss.getSheetByName("Document Generator").deleteRow(row.documentGeneratorRow+1);
  }
}

  // Update remaining learners in the sheet
  const headers = getAllHeaders(ss);
  const lastRows = findHighestRowNumbers(moduleLearners, ss);
  for (const learner of moduleLearners) {
    Logger.log("Reviewing " + learner.name);
    const rows = learner.getRows(ss);
    if (rows.attendanceSheetRow === -1) {
      addAttendanceSheetRow(ss, learner, Number(lastRows.attendance), headers.attendanceHeaders);
    } else {
      updateAttendanceSheetRow(ss, learner, rows.attendanceSheetRow, headers.attendanceHeaders);
    }
    if(rows.documentGeneratorRow === -1){
      addDocumentGeneratorRow(ss, learner, Number(lastRows.documentGenerator), headers.documentGeneratorHeaders, module);
    } else {
      updateDocumentGeneratorRow(ss, learner, rows.documentGeneratorRow, headers.documentGeneratorHeaders, module);
    }
  }

  // Log information about removed learners (if any)
  if (removedLearners.length > 0) {
    Logger.log(`Removed learners: ${removedLearners.join(', ')}`);
  }
}

function updateAttendanceSheetRow(ss, learner, row, headers){
  let studentRow = row+1;
  let sheet = ss.getSheets()[0];
  sheet.getRange(studentRow, headers.nameIndex+1).setValue(learner.getName());
  sheet.getRange(studentRow, headers.emailIndex+1).setValue(learner.email);
  sheet.getRange(studentRow, headers.bookingIdIndex+1).setValue(learner.bookingId);
  sheet.getRange(studentRow, headers.personNumberIndex+1).setValue(learner.personNumber);
}

function addAttendanceSheetRow(ss, learner, row, headers){
  let sheet = ss.getSheets()[0];
  sheet.insertRowAfter(row+1);
  let studentRow = row+2;
  sheet.getRange(studentRow, headers.nameIndex+1).setValue(learner.getName());
  sheet.getRange(studentRow, headers.emailIndex+1).setValue(learner.email);
  sheet.getRange(studentRow, headers.assignmentSubmittedIndex+1).insertCheckboxes();
  sheet.getRange(studentRow, headers.courseCompletedIndex+1).insertCheckboxes();
  sheet.getRange(studentRow, headers.lateSubmissionIndex+1).insertCheckboxes();
  for(i=6; i<sheet.getLastColumn(); i++){
    let test = sheet.getRange(5, i).getValues();
    if(test[0][0].toString().includes("Present")){
      sheet.getRange(studentRow, i).insertCheckboxes();
    }
  }
  sheet.getRange(studentRow, headers.bookingIdIndex+1).setValue(learner.bookingId);
  sheet.getRange(studentRow, headers.personNumberIndex+1).setValue(learner.personNumber);
}

function updateDocumentGeneratorRow(ss, learner, row, headers, module){
  let sheet = ss.getSheetByName("Document Generator");
  let learnerRow = row+1;
  sheet.getRange(learnerRow, headers.nameIndex+1).setValue(learner.getName());
  sheet.getRange(learnerRow, headers.emailIndex+1).setValue(learner.email);
  sheet.getRange(learnerRow, headers.sponsorIndex+1).setValue(learner.sponsor);
  sheet.getRange(learnerRow, headers.tutorIndex+1).setValue(module.tutorName);
  sheet.getRange(learnerRow, headers.dateIndex+1).setValue(module.startDate);
  sheet.getRange(learnerRow, headers.addressIndex+1).setValue(learner.address);
  sheet.getRange(learnerRow, headers.phoneIndex+1).setValue(learner.phone);
  sheet.getRange(learnerRow, headers.bookingIdIndex+1).setValue(learner.bookingId);
  sheet.getRange(learnerRow, headers.personNumberIndex+1).setValue(learner.personNumber);
  sheet.getRange(learnerRow, headers.paidIndex+1).insertCheckboxes();
  sheet.getRange(learnerRow, headers.paidIndex+1).setValue(learner.isBookingPaid()).insertCheckboxes();
  sheet.getRange(learnerRow, headers.sentIndex+1).insertCheckboxes();
  sheet.getRange(learnerRow, headers.resultsIndex+1).insertCheckboxes();

}

function addDocumentGeneratorRow(ss, learner, row, headers, module){
  let sheet = ss.getSheetByName("Document Generator");
  sheet.insertRowAfter(row+1);
  let learnerRow = row+1;
  updateDocumentGeneratorRow(ss, learner, learnerRow, headers, module);
}