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
      mobilePhoneIndexIndex, bookingNumberIndex, productCodeIndex;
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
      if(data[0][i].includes("Product code")){productCodeIndex = Number(i);}
    }
    data.shift(); //drop header row
    //find all courses on this date
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
          row[productCodeIndex],
          gmtStartDate,
          gmtEndDate,
          null,
          [row[bookingNumberIndex]]
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
  const ssId = "1WlwJS3d1H5JwzzRw3snBxTRpMvz2rSk8x4_Q8q_Ahpg"; //1 session with multiple bookings
  //const ssId = "1ccGcxkl_GwYk3WIGcxwsQSRR0VdAkYWTUNN3gcwh3_0"; //One Session
  const courses = compileCourses(ssId);
  for(course of courses){
    buildAttendanceSheet(course);
  }
}

function buildModuleFromSheet(ss){
  const settings = getSettings(ss.getId());
  //find booking ids for this course
  const moduleDetails = new ModuleDetails(
    settings.courseName,
    settings.location,
    settings.tutor,
    settings.courseId,
    settings.startDate,
    settings.endDate,
    settings
  );
  return moduleDetails;
}

function testBuildModuleFromSheet(){
  const ss = SpreadsheetApp.openById("1taDg6z7Dekk7AjPyouvLZZwCIhbBeJ5GN6J2LJ5xnu4");
  const moduleDetails = buildModuleFromSheet(ss);
  for(learner of moduleDetails.getLearners()){
    Logger.log(learner.name+" on row "+learner.getRows(ss).attendanceSheetRow);
  }
}