class CourseDetails {
  constructor(moduleName, deliveryMode, tutorName, studentDetails, events, startDate, end) {
      this.moduleName = moduleName;
      this.tutorName = tutorName;
      this.studentDetails = studentDetails;
      this.deliveryMode = deliveryMode;
      this.events = events;
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
    let courseData = this.events;
    let sessions = 4;
    for(i in courseData){
        if(courseData[i][0].trim().includes(this.moduleName.trim())){
            sessions = courseData[i][1];
        }
    }
    if(sessions == 0){
      return 4;
    }
    return sessions;
  }
  getEnd(){
    if(this.end){
      return this.end;
    } else {
      let endDate = new Date(this.startDate);
      Date.setDate(date.getDate()+(7*sessions()));
      return endDate;
    }
  }
}

class StudentDetails {

    constructor(firstName, lastName, email, sponsor, address, phone, bookingId) {
      this.firstName = firstName;
      this.lastName = lastName;
      this.email = email;
      this.sponsor = sponsor;
      this.address = address; 
      this.phone = phone;
      this.bookingId = bookingId;
    }
  
    get name() {
      return `${this.firstName} ${this.lastName}`;
    }
  
    getName() {
      return this.name; 
    }
  
  }

  class LearnerDetails {
    constructor(firstName, lastName, email, sponsor, address, phone, bookingId, personNumber, paid = false, passed = false, sent = false, letterUrl = "", certUrl = "") {
      this.firstName = firstName;
      this.lastName = lastName;
      this.email = email;
      this.sponsor = sponsor;
      this.address = address;
      this.phone = phone;
      this.bookingId = bookingId;
      this.personNumber = personNumber;

      // Optional Properties
      this.paid = paid;
      this.passed = passed;
      this.sent = sent;
      this.letterUrl = letterUrl;
      this.certUrl = certUrl;
    }

    get name() {
      return `${this.firstName} ${this.lastName}`;
    }
  
    getName() {
      return this.name; 
    }
    getUniqueKey() {
      return `${this.bookingId}-${this.personNumber}`;
    }
    isBookingPaid() { 
      let keys = getBookeoApiKeys();
      let booking = bookeoLibrary.getBookingById(this.bookingId, keys.apiKey, keys.apiSecret);
      // Assuming the booking object has a property like "price.totalPaid.amount"
      return booking.price.totalPaid.amount === booking.price.totalGross.amount; 
    }
    getAttendanceRecords(ss) {
      // get the attendance sheet
      const sheet = ss.getSheets()[0];
      const data = sheet.getDataRange().getValues();
      // find the header indexes 
      const headers = data[2];
      const headers2 = data[5];
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
      Logger.log("Headers: " + headers);
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
      //find the booking id
      let bookingId = this.bookingId;
      let personNumber = this.personNumber;
      for(let i=0; i<data.length; i++){
        if(data[i][bookingIdIndex] == bookingId && data[i][personNumberIndex] == personNumber){
          Logger.log("Found attendance for " + this.name);
          let learner = {
            name: data[i][nameIndex],
            email: data[i][emailIndex],
            bookingId: data[i][bookingIdIndex],
            personNumber: data[i][personNumberIndex],
            assignmentSubmitted: data[i][assignmentSubmittedIndex],
            courseCompleted: data[i][courseCompletedIndex],
            lateSubmission: data[i][lateSubmissionIndex],
            sessions: []
          };
          for(let j=0; j<sessionHeaders.length; j++){
            let session = sessionHeaders[j];
            let present = data[i][sessionHeaders[j].presentIndex];
            let note = data[i][sessionHeaders[j].noteIndex];
            learner.sessions.push({
              name: session.name,
              number: session.number,
              present: present,
              note: note
            });
          }
          return learner;
        }
      }
    }
  }

  class ModuleDetails {
    constructor(moduleName, deliveryMode, tutorName, bookingIds, courseData, startDate, endDate, settings, events=[]) {
      this.moduleName = moduleName;
      this.deliveryMode = deliveryMode;
      this.tutorName = tutorName;
      this.bookingIds = bookingIds;
      this.courseData = courseData;
      this.startDate = startDate;
      this.endDate = endDate;
      this.settings = settings;
      this.events = events;
    }
    courseId(){
      let courseData = this.courseData;
      let courseDataArray = courseData.split("-");
      if(courseDataArray.length > 1){
        return courseDataArray[0];
      }else {
        return courseDataArray;
      }
    }
    sessions(){
      let courseData = this.courseData;
      let courseDataArray = courseData.split("-");
      if(courseDataArray.length > 1){
        return courseDataArray[2];
      }else if(isSameDay(this.startDate, this.endDate)){
        return 1;
      } else {
        return 4;
      }
    }
    getDeliveryMethod(){
      let courseData = this.courseData;
      let courseDataArray = courseData.split("-");
      let deliveryMethod;
      if(courseDataArray.length > 1){
        deliveryMethod = courseDataArray[1];
      } else {
        deliveryMethod = "Not Available";
      }

      if(deliveryMethod == "OL"){return  "Online";}
      else if(deliveryMethod == "IC"){return "In Class";}
      else if(deliveryMethod == "BL"){return "Blended";}
      else {return "Not Available";}
    }
    getEnd(){
      if(this.endDate){
        return this.endDate;
      } else {
        let endDate = new Date(this.startDate);
        endDate.setDate(endDate.getDate()+(7*this.sessions()));
        return endDate;
      }
    }
    getLearners(){
      let learnerArray = [];
      const uniqueBookingIds= []
      for(let bookingId of this.bookingIds){
        if(uniqueBookingIds.includes(bookingId)){
          continue;
        } else {
          uniqueBookingIds.push(bookingId);
        }
      }
      for(let bookingId of uniqueBookingIds){
        let booking = getBookingById(bookingId);
        let sponsor = booking.customer.emailAddress;
        let learners = booking.participants.details;
        for(let learner of learners){
          let email, firstName, lastName, address, phone;
          if(learner.personId == "PUNKNOWN"){
            email = "";
            firstName = "";
            lastName = "";
            address = "";
            phone = "";
          } else {
            firstName = learner.personDetails.firstName;
            lastName = learner.personDetails.lastName;
            try{email = learner.personDetails.emailAddress;}
            catch(e) {email = "";}
            try{address = learner.personDetails.streetAddress.address1+"\n"+learner.personDetails.streetAddress.address2;}
            catch(e) {address = "";}
            try{
              for(numbers of learner.personDetails.phoneNumbers){
                phone+=numbers.type+":"+numbers.number+"\n";
              }
            }
            catch(e) {phone = "";}
          }
          let personNumber = learner.categoryIndex;
          let learnerDetails = new LearnerDetails(firstName, lastName, email, sponsor, address, phone, bookingId, personNumber);
          learnerDetails.productId = booking.productId;
          learnerArray.push(learnerDetails);
        }
      }
      return learnerArray;
    }
    issuedOn(){
      let issuedOnDate = new Date(this.getEnd());
      return issuedOnDate;
    }
    renewsOn(){
      let issuedOnDate = new Date(this.getEnd());
      let renewsOnDate = new Date(issuedOnDate.setFullYear(issuedOnDate.getFullYear()+settings.renewalDuration))
      return renewsOnDate;
    }
    getClassId(){
      let courseCode = this.courseId();
      let tutor = this.tutorName;
      let startDate = this.startDate;
      return courseCode+"-"+tutor+"-"+startDate;
    }
}