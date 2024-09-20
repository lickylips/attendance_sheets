function compileCourses(ssId){
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
      const millisecondsInEightHours = 8 * 60 * 60 * 1000;
      let gmtStartDate = new Date(Date.parse(row[startDateIndex]));
      let gmtEndDate = new Date(Date.parse(row[endIndex]));
      const courseKey = row[courseIndex] + " - " + row[tutorIndex]; // Composite key


      if(courseKeys.indexOf(courseKey) == -1){ //if course not already added
        Logger.log("New course "+row[courseIndex]+" Being Created")
        courseKeys.push(courseKey);
        let course = new ModuleDetails(
          row[courseIndex],
          row[locationIndex],
          row[tutorIndex],
          [row[bookingNumberIndex]],
          courseData,
          gmtStartDate,
          gmtEndDate
        );
        courses.push(course);
      }
      else{//if course already added
        let course;
        for(line of courses){
          if(line.moduleName == row[courseIndex]){
            course = line;
            course.bookingIds.push(row[bookingNumberIndex]);
          }
        }
      }
    }
    return(courses);
}

function testCompileCourses(){
  const ssId = "1RYs9q9wnoTyL_EdrdpGIchOlNgOn1pmh6ekwyygaZlw"; //4 sessions
  //const ssId = "1ccGcxkl_GwYk3WIGcxwsQSRR0VdAkYWTUNN3gcwh3_0"; //One Session
  const courses = compileCourses(ssId);
  for(course of courses){
    buildAttendanceSheet(course);
  }
}

function extractCourseFromSheet(){
  const studentSheet = getStudentArray();
  const studentArray = studentSheet.getDataRange().getValues();
  const settings = getSettings();
  //get headding Indexs
  let headers = studentArray.shift()
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
  const learners = [];
  for(row of studentArray){
    let nameParts = splitName(row[nameIndex]);
    const learner = new Learner(
      nameParts[0],
      nameParts[1],
      row[emailIndex],
      row[sponsorIndex],
      row[addressIndex],
      row[phoneIndex],
      row[bookingIdIndex],
      row[personNumberIndex],
      row[paidIndex],
      row[coursePassedIndex],
      row[sentIndex],
      row[letterIndex],
      row[certIndex],
    );
    learners.push(learner);
  }
  return(learners);
}