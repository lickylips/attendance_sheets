function extractEmail(input) {
    Logger.log("Extracting tutor email from "+input);
    // Split the input into lines
    var lines = input.split('\n');
  
    // Initialize the email variable
    var email = "Email not found";
  
    // Loop through the lines to find the email
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
  
      // Check if the line contains "@" to identify an email
      if (line.includes("@")) {
        email = line;
        break; // Exit the loop after finding the email
      }
    }
  
    return email;
  }

  function extractName(input) {
    // Split the input into lines
    var lines = input.split('\n');
  
    // Initialize the name variable
    var name = "Name not found";
  
    // Loop through the lines to find the name
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
  
      // Check if the line is not empty and does not start with ###
      if (line && !line.startsWith("###")) {
        name = line;
        break; // Exit the loop after finding the name
      }
    }
  
    return name;
  }

  function publishAttendanceSheets(opSheets){
    Logger.log("Publishing Attendance Sheets");
    const docId = "1jIZB4ywPC2CDlgSbWx7Muqbsm27Q9DoVBg4uVEvai_0";
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const date = new Date(opSheets[0].course.startDate.getTime());
    const dateParagraphIndex = findOrInsertDate(body, date);
    for(i in opSheets){
      sheet = opSheets[i];
      const courseLine = sheet.course.moduleName + " - " + sheet.course.deliveryMode;
      const courseParagraph = body.insertParagraph(dateParagraphIndex+1, courseLine);
      courseParagraph.setLinkUrl(sheet.sheet);
    }
  }

  function emailAttendanceSheets(email, opSheets){
    Logger.log("Emailing Results to "+ email);
    let template = HtmlService.createTemplateFromFile("emailBody");
    const urls = [];
    const locations = [];
    const tutor = [];
    const courseName = [];
    for(i in opSheets){
      urls.push(opSheets[i].sheet);
      locations.push(opSheets[i].course.deliveryMode);
      tutor.push(opSheets[i].course.tutorName);
      courseName.push(opSheets[i].course.moduleName);
    }
    const messageContent = {
      urls: urls,
      locations: locations,
      tutor: tutor,
      courseName: courseName
    }
    template.messageContent = messageContent;
    const message = template.evaluate().getContent();
    const mail = {
      to: email,
      replyTo: "info@ncultd.ie",
      subject: "Upcoming Course Attendance Sheets",
      htmlBody: message
    }
    MailApp.sendEmail(mail);
  }

  function tutorNotificationEmail(course, url){
    Logger.log("Sending tutor notification email");
    const tutorEmail = extractEmail(course.tutor);
    let template = HtmlService.createTemplateFromFile("tutorEmail");
    const messageContent = {
      url: url,
      course: course
    }
    template.messageContent = messageContent;
    const message = template.evaluate().getContent();
    const mail = {
      to: course.tutorEmail(),
      replyTo: "info@ncultd.ie",
      subject: "Upcoming GNC Course Details",
      htmlBody: message
    }
      
  }