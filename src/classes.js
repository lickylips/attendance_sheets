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
    end(){
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