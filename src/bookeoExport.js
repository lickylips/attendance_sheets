function getBookeoBookingsForDate(date, productId) {
  if(!date) {
    //date = new Date();
    date = new Date("2024-05-03"); // For testing purposes
  }
  const apiKey = 'AKEJWP7UJA3XP7XHU6FNE224NR4XX3148FA63EA11';
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
  Logger.log(url);

  // Fetch data from Bookeo API
  const response = UrlFetchApp.fetch(url);
  const bookingsData = JSON.parse(response.getContentText());


  Logger.log(url);
  return bookingsData;
}
  
function getCoursesForDate(date) {
  if(!date) {
    //date = new Date();
    date = new Date("2024-05-03"); // For testing purposes
  }
  const apiKey = 'AKEJWP7UJA3XP7XHU6FNE224NR4XX3148FA63EA11';
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
  Logger.log(url);

  return courseData;
}

function buildBookeoCourses(date){
  if(!date) {
    //date = new Date();
    date = new Date("2024-05-03"); // For testing purposes
  }
  const courseData = getCoursesForDate(date);
  const courses = courseData.data;
  const courseObjects = [];
  for(let course of courses){
    const sessions = course.courseSchedule.events;
    const startDate = new Date(course.startTime);
    const endDate = new Date(sessions[sessions.length-1].endTime);
    const deliveryMode = course.courseSchedule.title;
    const tutorName = course.resources[0].name;
    const bookings = getBookeoBookingsForDate(startDate, course.productId).data;
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
    }
    const courseObject = new CourseDetails(moduleName, deliveryMode, tutorName, learners, sessions, startDate, endDate);
    Logger.log(sessions[0][0])
    courseObjects.push(courseObject);
  }
  return courseObjects;
}

function buildBookeoCourses2(date){
  if(!date) {
    //date = new Date();
    date = new Date("2024-05-03"); // For testing purposes
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