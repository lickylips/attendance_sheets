function addToTracker(course, attendanceSheet){
    const tracker = readTrackerSheet();
    const ss = tracker.trackerSheet;
    const sheet = ss.getSheetByName("Tracker");
    const attendanceSheets = tracker.course;
    const headers = tracker.headersObject;
    //check if course exists
    const classId = course.getClassId();
    let exists = false;
    let row = 2;
    let created = new Date();
    let lastReCreated = "";
    Logger.log(attendanceSheets);
    for(let course of attendanceSheets){
      if(course.classId == classId){
        Logger.log("Course already exists in tracker - updating");
        exists = true;
        Logger.log(course.row)
        row = course.row;
        created = course.created;
        lastReCreated = new Date();
      }
    }
    if(!exists){
      Logger.log("Adding new course to tracker");
      sheet.insertRowBefore(2);
    }
    let hyperlinkFormula = `=HYPERLINK("${attendanceSheet.getUrl()}", "${course.moduleName}")`;


    //add details to row
    Logger.log(headers.startIndex);
    sheet.getRange(row, headers.startIndex+1).setValue(course.startDate); 
    sheet.getRange(row, headers.courseNameIndex+1).setFormula(hyperlinkFormula); // Column 2
    sheet.getRange(row, headers.tutorIndex+1).setValue(course.tutorName); // Column 3
    sheet.getRange(row, headers.sessionsIndex+1).setValue(0); // Column 4
    sheet.getRange(row, headers.resultsIndex+1).setValue(false)
      .insertCheckboxes(); // Column 5
    sheet.getRange(row, headers.eaSubmittedIndex+1).setValue(false)
      .insertCheckboxes(); // Column 6
    sheet.getRange(row, headers.classIdIndex+1).setValue(course.getClassId()); // Column 7
    sheet.getRange(row, headers.createdIndex+1).setValue(created); // Column 8
    sheet.getRange(row, headers.lastReCreatedIndex+1).setValue(lastReCreated); // Column 9
    sheet.getRange(row, headers.folderIndex+1).setValue(attendanceSheet.getUrl()); // Column 10
    sheet.getRange(row, headers.sheetIndex+1).setValue(getSpreadsheetFolderUrl(attendanceSheet.getId())); // Column 11
}

function readTrackerSheet(){
    const trackerId = "1Jxs9tD058w24xHoFjZFJOP9Lt2-HCthZ1ctT4hsgXL8";
    const ss = SpreadsheetApp.openById(trackerId);
    const trackerSheet = ss.getSheetByName("Tracker");
    const trackerArray = trackerSheet.getDataRange().getValues();
    const attendanceSheets = [];
    //get headers
    const headers = trackerArray[0];
    const startIndex = headers.indexOf("Start Date");
    const courseNameIndex = headers.indexOf("Course Name");
    const tutorIndex = headers.indexOf("Tutor");
    const sessionsIndex = headers.indexOf("Sessions Complete");
    const resultsindex = headers.indexOf("Results Sent");
    const eaSubmittedIndex = headers.indexOf("EA Submitted");
    const classIdIndex = headers.indexOf("Class ID");
    const createdIndex = headers.indexOf("Created");
    const lastReCreatedIndex = headers.indexOf("Last Recreated");
    const folderIndex = headers.indexOf("Folder");
    const sheetIndex = headers.indexOf("Sheet");
    //create headersObject
    const headersObject = {
        startIndex: startIndex,
        courseNameIndex: courseNameIndex,
        tutorIndex: tutorIndex,
        sessionsIndex: sessionsIndex,
        resultsIndex: resultsindex,
        eaSubmittedIndex: eaSubmittedIndex,
        classIdIndex: classIdIndex,
        createdIndex: createdIndex,
        lastReCreatedIndex: lastReCreatedIndex,
        folderIndex: folderIndex,
        sheetIndex: sheetIndex
    };
    trackerArray.shift(); //remove headers
    //loop through tracker array
    for(let i = 0; i < trackerArray.length; i++){
        const row = trackerArray[i];
        let attendanceSheet = {
            row: i+2,
            startDate: row[startIndex],
            courseName: row[courseNameIndex],
            tutor: row[tutorIndex],
            sessions: row[sessionsIndex],
            resultsSent: row[resultsindex],
            eaSubmitted: row[eaSubmittedIndex],
            classId: row[classIdIndex],
            created: row[createdIndex],
            lastReCreated: row[lastReCreatedIndex],
            folder: row[folderIndex],
            sheet: row[sheetIndex],
        };
        attendanceSheets.push(attendanceSheet);
    }
    let attendanceSheetsObject = {
        course: attendanceSheets,
        trackerSheet: ss,
        headersObject: headersObject
    };
    return attendanceSheetsObject;
}