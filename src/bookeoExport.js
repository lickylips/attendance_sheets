//get the bookeo API and Secret Keys
function getBookeoApiKeys() {
  let keys = {
      apiKey: PropertiesService.getScriptProperties().getProperty('BOOKEO_API_KEY'),
      secretKey: PropertiesService.getScriptProperties().getProperty('BOOKEO_SECRET_KEY')
  };
  return keys;
}

function getBookeoBookingsForDate(date, productId) {
  if(!date) {
    //date = new Date();
    date = new Date("2024-05-03"); // For testing purposes
  }
  const apiKey = 'AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11';
  const secretKey = '5ajggnHkopp3KCWXnHN5BDJRYjK3oweX';
  const apiUrlBase = 'https://api.bookeo.com/v2/'; 

  // Calculate start and end dates for the given day (UTC timezone)
  const nowUtc = date.toISOString(); // Get current time in ISO 8601 (UTC)
  const startTime = nowUtc.slice(0, 10) + 'T00:00:00Z'; // Start of today UTC in RFC 3339
  const endTime = new Date(Date.parse(nowUtc))
    .toISOString().slice(0, 10) + 'T23:59:59Z'; // Start of next week UTC in RFC 3339

  // Construct API request URL
  let url = `${apiUrlBase}bookings?apiKey=${encodeURIComponent(apiKey)}&secretKey=${encodeURIComponent(secretKey)}&startTime=${encodeURIComponent(startTime)}&endTime=${encodeURIComponent(endTime)}`;
  url+= "&expandCustomer=true";
  url+= "&expandParticipants=true";
  if(productId) {
    url+= "&productId="+productId;
  }
  // Fetch data from Bookeo API
  const response = UrlFetchApp.fetch(url);
  const bookingsData = JSON.parse(response.getContentText());


  Logger.log("Bookings for Date URL: "+url);
  return bookingsData;
}
  

function testGetBookeoBookingsForDate(){
  const date = new Date("2024-09-17"); // For testing purposes
  const productId = "2229XUEWF191DB0F6226";
  const bookings = getBookeoBookingsForDate(date, productId);
  Logger.log(bookings);
}

function getCoursesForDate(date) {
  if(!date) {
    //date = new Date();
    date = new Date("2024-05-03"); // For testing purposes
  }
  const apiKey = 'AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11';
  const secretKey = '5ajggnHkopp3KCWXnHN5BDJRYjK3oweX';
  const apiUrlBase = 'https://api.bookeo.com/v2/';

  // Calculate start and end dates for the given day (UTC timezone)
  const nowUtc = date.toISOString(); // Get current time in ISO 8601 (UTC)
  const startTime = nowUtc.slice(0, 10) + 'T00:00:00Z'; // Start of today UTC in RFC 3339
  const endTime = new Date(Date.parse(nowUtc))
    .toISOString().slice(0, 10) + 'T23:59:59Z'; // Start of next week UTC in RFC 3339

  // Construct API request URL
  let url = `${apiUrlBase}availability/slots?apiKey=${encodeURIComponent(apiKey)}&secretKey=${encodeURIComponent(secretKey)}&startTime=${encodeURIComponent(startTime)}&endTime=${encodeURIComponent(endTime)}`;
  
  // Fetch data from Bookeo API
  const response = UrlFetchApp.fetch(url);
  const courseData = JSON.parse(response.getContentText());
  Logger.log("Courses for Date API URL: "+url);

  return courseData;
}

function testGetBookeoBookingsForDate(){
  const date = new Date("2024-09-17");
  const bookings = getBookeoBookingsForDate(date);
  Logger.log(bookings);
}

function buildBookeoCourses(date){
  if(!date) {
    //date = new Date();
    date = new Date("2024-09-17"); // For testing purposes
  }
  Logger.log("Getting course data for date: " + date);
  const courseData = getCoursesForDate(date);
  const courses = courseData.data;
  const courseObjects = [];
  Logger.log("Found " + courses.length + " courses");
  for(let course of courses){
    Logger.log("Course: " + course.productId);
    const sessions = course.courseSchedule.events;
    const startDate = new Date(course.startTime);
    const endDate = new Date(sessions[sessions.length-1].endTime);
    const deliveryMode = course.courseSchedule.title;
    const tutorName = course.resources[0].name;
    const bookings = getBookeoBookingsForDate(startDate, course.productId).data;
    Logger.log("Found " + bookings.length + " bookings");
    const moduleName = bookings[0].productName;
    let learners = [];
    for(let booking of bookings){
      const sponsor = booking.customer.emailAddress;
      for(let learner of booking.participants.details){
        let email, firstName, lastName, address;
        let phone = "";
        try{
          email = learner.personDetails.emailAddress;
        } 
        catch(e) {
          email = "";
        }
        try{firstName = learner.personDetails.firstName;}
        catch(e) {firstName = "";}
        try{lastName = learner.personDetails.lastName;}
        catch(e) {lastName = "";}
        try{address = learner.personDetails.streetAddress.address1+"\n"+learner.personDetails.streetAddress.address2;}
        catch(e) {address = "";}
        try{
          for(numbers of learner.personDetails.phoneNumbers){
            phone+=numbers.type+":"+numbers.number+"\n";
          }
        }
        catch(e) {phone = "";}
        const studentDetails = new StudentDetails(firstName, lastName, email, sponsor, address, phone);
        learners.push(studentDetails);
      }
      Logger.log("Booking ID " + booking.bookingNumber + " has " + learners.length + " learners");
    }
    const courseObject = new CourseDetails(moduleName, deliveryMode, tutorName, learners, sessions, startDate, endDate);
    courseObjects.push(courseObject);
  }
  Logger.log("Compiled " + courseObjects.length + " courses");
  return courseObjects;
}

function buildBookeoCourses2(date){
  if(!date) {
    //date = new Date();
    date = new Date("2024-09-17"); // For testing purposes
  }
  const bookings = getBookeoBookingsForDate(date);
  const productNames = [];
  for(data of bookings.data){
    if(productNames.includes(data.productName)){
      continue;
    } else {
      productNames.push(data.productName);
    }
  }
  for(let productName of productNames){
    Logger.log("Creating "+ productName);
    Logger.log(bookings);
  }
  Logger.log(productNames);
}

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

function getBookingById(bookingId){
  const apiKey = 'AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11';
  const secretKey = '5ajggnHkopp3KCWXnHN5BDJRYjK3oweX';
  const apiUrlBase = 'https://api.bookeo.com/v2/'; 

  // Construct API request URL
  let url = `${apiUrlBase}bookings/${bookingId}?apiKey=${encodeURIComponent(apiKey)}&secretKey=${encodeURIComponent(secretKey)}`;
  url+= "&expandCustomer=true";
  url+= "&expandParticipants=true";
  Logger.log("Get booking by ID URL: "+url);

  // Fetch data from Bookeo API
  const response = UrlFetchApp.fetch(url);
  const bookingsData = JSON.parse(response.getContentText());

  return bookingsData;
}

function testGetBookingById(){
  const bookingId = "22405213108336";
  const bookings = getBookingById(bookingId);
  Logger.log(bookings);
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

function getCourseSettings(productId){
  const apiKey = 'AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11';
  const secretKey = '5ajggnHkopp3KCWXnHN5BDJRYjK3oweX';
  const apiUrlBase = 'https://api.bookeo.com/v2/';

  // Construct API request URL
  let url = `${apiUrlBase}settings/products?apiKey=${encodeURIComponent(apiKey)}&secretKey=${encodeURIComponent(secretKey)}`;
  Logger.log(url);

  // Fetch data from Bookeo API 
  const response = UrlFetchApp.fetch(url);
  const products = JSON.parse(response.getContentText());
  const product = products.data.find(product => product.productId === productId);

  return product;
}

function testgetCourseSettings(){
  const productId = "224XW643R149D19C779C";
  const courseSettings = getCourseSettings(productId);
  Logger.log(courseSettings);
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

  for(let i=0; i<studentArray.length; i++){
    Logger.log("Checking Booking for: "+studentArray[i][nameIndex]);
    let learner = studentArray[i];
    let bookeoLearner;
    let bookingDetails = getBookingById(learner[bookingIdIndex]);
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

function getLearnerDetails(bookingId, personNumber){
  const apiKey = 'AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11';
  const secretKey = '5ajggnHkopp3KCWXnHN5BDJRYjK3oweX';
  const apiUrlBase = 'https://api.bookeo.com/v2/';
  const booking = getBookingById(bookingId);
  Logger.log(booking);
  const personDetails = booking.participants.details;
  for(person of personDetails){
    if(person.categoryIndex === personNumber){
      return person;
    }
  }
  Logger.log(person)
  return person;
}

function getCustomerDetails(bookingId){
  const apiKey = 'AXYXHY6PRA3XP7XHU6FNE224NR4XX3148FA63EA11';
  const secretKey = '5ajggnHkopp3KCWXnHN5BDJRYjK3oweX';
  const apiUrlBase = 'https://api.bookeo.com/v2/';
  const booking = getBookingById(bookingId);
  const customer = booking.customer;
  Logger.log(customer)
  return customer;
}

function testgetLearnerDetails(){
  const bookingId = "22402237741173";
  const personNumber = 1;
  const learnerDetails = getLearnerDetails(bookingId, personNumber);
  Logger.log(learnerDetails);
}

function testBookeoCompileCourses(){
  const date = new Date("2024-09-17");
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
  //emailAttendanceSheets(email, opSheets);
  //publishAttendanceSheets(opSheets);
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
    const booking = getBookingById(bookingId);
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