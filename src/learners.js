function buildLearnerObject(name, ss){
    //1 Open the Spreadsheet and prepare the search
    if(ss == null){
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    const sheet = ss.getSheetByName("Document Generator");
    const data = sheet.getDataRange().getValues();

    //2. Search for the name in the spreadsheet
    const headers = data[0];
    const nameIndex = headers.indexOf("Name");
    const bookingIdIndex = headers.indexOf("Booking ID");
    const personNumberIndex = headers.indexOf("Person Number");
    const certIndex = headers.indexOf("Cert");
    const letterIndex = headers.indexOf("Letter");
    const sendIndex = headers.indexOf("Sent");
    let resultsIndex = headers.indexOf("Results Sent");
    //handle missing results index
    if(resultsIndex == -1){
      resultsIndex = addMissingResultsSentField(ss);
    }
    data.shift(); // Remove the header row
    const row = data.find(row => row[nameIndex] === name);
    if (!row) {
      Logger.log(`Name "${name}" not found in the spreadsheet.`);
    } else {
        //3. Create the Learner object
        const bookingId = row[bookingIdIndex];
        const personNumber = row[personNumberIndex];
        const keys = getBookeoApiKeys();
        const booking = bookeoLibrary.getBookingById(bookingId, keys.apiKey, keys.secretKey);
        const participantDetails = booking.participants.details.find(detail => detail.categoryIndex === personNumber);
        if (participantDetails) {
            Logger.log(participantDetails.personDetails); // Log the person details
            // Build the address string using the existing function
            let address = "";
            if(participantDetails.streetAddress){
              try{
                address = addressStringBuilder(
                  participantDetails.streetAddress.address1,
                  participantDetails.streetAddress.address2 || "", // Handle potential missing address2
                  participantDetails.streetAddress.city || "" // Handle potential missing city
                );
              }
              catch(err){
                Logger.log(err);
              }
            }
            // Build the phone string using the existing function
            const phone = phoneStringBuilder(participantDetails.phoneNumbers);
            
            // Create the Learner object
            const learner = new LearnerDetails(
                participantDetails.personDetails.firstName,
                participantDetails.personDetails.lastName,
                participantDetails.personDetails.emailAddress,
                booking.customer.emailAddress, // Sponsor email is now retrieved from the booking
                address, 
                phone, 
                bookingId,
                personNumber,
                booking.price.totalPaid.amount === booking.price.totalGross.amount,
                false,
                row[sendIndex],
                row[certIndex],
                row[letterIndex]
            );
            learner.resultsSent = row[resultsIndex];
            const settings = getSettings(ss.getId());
            learner.settings = settings;
            return learner;
          } else {
            Logger.log(`Person with ID "${personNumber}" not found in the booking.`);
            //create learner from sheet stuff
            let name = splitName(row[nameIndex]);
            const learner = new LearnerDetails(
              name[0],
              name[1],
              "deleted",
              "deleted",
              "",
              "",
              row[bookingIdIndex],
              row[personNumberIndex],
              "",
              "",
              false,
              false,
              false,
            );
            return learner;
          }
    }
}

function testBuildLearnerObject(){
  const name = "William Bob";
  const ss = SpreadsheetApp.openById("1ZOequ0t3RKz45BpM7NqY9MRmS97NGHt3R2gTXvjtUgU")
  const learner = buildLearnerObject(name, ss);
  const attendance = learner.getAttendanceRecords(ss);
  Logger.log("learner: "+learner);
  for(session of attendance.sessions){
    Logger.log(session.name+" - "+session.present);
  }
}

function getLearnerArray(ss){
  const sheet = ss.getSheetByName("Document Generator");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameIndex = headers.indexOf("Name");
  data.shift(); // Remove the header row
  let learners = [];
  for (const row of data) {
    const name = row[nameIndex];
    const learner = buildLearnerObject(name, ss);
    learners.push(learner);
  }
  return learners;
}

function testGetLearnerArray(){
  const ss = SpreadsheetApp.openById("1ZOequ0t3RKz45BpM7NqY9MRmS97NGHt3R2gTXvjtUgU");
  const learners = getLearnerArray(ss);
  for(learner of learners){
    Logger.log(learner.name);
  }
}

function getBookingCounts(learners) {
  const bookingCounts = {};
  for (const learner of learners) {
    const bookingId = learner.bookingId;
    if (bookingCounts[bookingId]) {
      bookingCounts[bookingId]++; // Increment count for existing booking
    } else {
      bookingCounts[bookingId] = 1; // Initialize count for new booking
    }
  }
  return bookingCounts;
}

function addLearnerRowToAttendanceSheet(learner, sheet, studentRow){
  //paste first student row
  sheet.getRange(studentRow, 1).setValue(learner.getName());
  sheet.getRange(studentRow, 2).setValue(learner.email);
  sheet.getRange(studentRow, 3).insertCheckboxes();
  sheet.getRange(studentRow, 4).insertCheckboxes();
  sheet.getRange(studentRow, 5).insertCheckboxes();
  for(i=6; i<sheet.getLastColumn(); i++){
    let test = sheet.getRange(5, i).getValues();
    if(test[0][0].toString().includes("Present")){
      sheet.getRange(studentRow, i).insertCheckboxes();
    }
  }
  sheet.getRange(studentRow, startCol).setValue(learner.bookingId);
  sheet.getRange(studentRow, startCol+1).setValue(learner.personNumber);
}

function buildLearnerObjectByBooking(booking, personNumber){
  const keys = getBookeoApiKeys();
  //const booking = bookeoLibrary.getBookingById(booking.id, keys.apiKey, keys.secretKey);
  const participantDetails = booking.participants.details.find(detail => detail.categoryIndex === personNumber);
  let phone, address;
  if (participantDetails) {
    // Build the address string using the existing function
    address = "";
    if(participantDetails.streetAddress){
      try{
        address = addressStringBuilder(
          participantDetails.streetAddress.address1,
          participantDetails.streetAddress.address2 || "", // Handle potential missing address2
          participantDetails.streetAddress.city || "" // Handle potential missing city
        );
      }
      catch(err){
        Logger.log(err);
      }
    }
    // build phone string using existing function
    phone = phoneStringBuilder(participantDetails.phoneNumbers);
  }
  // Create the Learner object
  const learner = new LearnerDetails(
      participantDetails.personDetails.firstName,
      participantDetails.personDetails.lastName,
      participantDetails.personDetails.emailAddress,
      booking.customer.emailAddress, // Sponsor email is now retrieved from the booking
      address,
      phone,
      booking.id,
      personNumber);
  return learner;    
}