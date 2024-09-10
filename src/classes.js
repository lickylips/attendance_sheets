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
  }

  class ModuleDetails {
    constructor(moduleName, deliveryMode, tutorName, bookingIds, courseData, startDate, endDate, settings) {
      this.moduleName = moduleName;
      this.deliveryMode = deliveryMode;
      this.tutorName = tutorName;
      this.bookingIds = bookingIds;
      this.courseData = courseData;
      this.startDate = startDate;
      this.endDate = endDate;
      this.settings = settings;
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
}