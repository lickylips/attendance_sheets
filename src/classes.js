class CourseDetails {
    constructor(moduleName, deliveryMode, tutorName, studentDetails,  events, startDate, end) {
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
        return this.events.length
    }
}

class StudentDetails {

    constructor(firstName, lastName, email, sponsor, address, phone) {
      this.firstName = firstName;
      this.lastName = lastName;
      this.email = email;
      this.sponsor = sponsor;
      this.address = address; 
      this.phone = phone;
    }
  
    get name() {
      return `${this.firstName} ${this.lastName}`;
    }
  
    getName() {
      return this.name; 
    }
  
  }