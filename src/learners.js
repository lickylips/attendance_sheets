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
        const booking = bookeoLibrary.getBookingById(bookingId);
        const participantDetails = booking.participants.details.find(detail => detail.categoryIndex === personNumber);

        if (participantDetails) {
            Logger.log(participantDetails.personDetails); // Log the person details
            // Build the address string using the existing function
            let address = "";
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
            return learner;
          } else {
            Logger.log(`Person with ID "${personNumber}" not found in the booking.`);
          }
    }
}

function testBuildLearnerObject(){
  const name = "William Bob";
  const ss = SpreadsheetApp.openById("1RPGTUZiJ4BUr0mpZLga_udpAa5izSnkn4pBXy5h__Dc")
  const learner = buildLearnerObject(name, ss);
  Logger.log(learner);
}