function createSummarySheet(docId, course){
    Logger.log("creating summary sheet");
    const ss = SpreadsheetApp.openById(docId);
    const summarySheet = ss.insertSheet("Results Summary Sheet");
    
    //Add Document Headder
    let startRow = 1;
    summarySheet.getRange(startRow,1).setValue("Course Results Summary")
    .setFontSize(20)
    .setHorizontalAlignment("center");
    summarySheet.getRange(startRow,1,1,10).merge();
    startRow++;

    //Logo
  const logoUrl = "https://lickylip.net/wp-content/uploads/2023/09/21-small.png";
  const image = SpreadsheetApp.newCellImage()
                              .setSourceUrl(logoUrl)
                              .build();
  summarySheet.getRange(startRow,1).setValue(image)
                               .setHorizontalAlignment("center");
  summarySheet.setRowHeight(startRow, 100);
  summarySheet.getRange(startRow,1, 1, 10).merge();
  startRow++;

  //Marks Level Tables
  let marks = [
    ["Marks", "Grade"],
    ["0-49", "F"],
    ["50-64", "P"],
    ["65-79", "M"],
    ["80-100", "D"],
  ]
  const marksTable = summarySheet.getRange(startRow,2, 5, 2).setValues(marks);
  let borderStyle = SpreadsheetApp.BorderStyle.SOLID_THICK; // Choose your style
  let borderColor = "#4B3A71"; // GNC color
  marksTable.setHorizontalAlignment("center");
  marksTable.setBorder(true, true, true, true, true, true, borderColor, borderStyle);

  //Course Details
  let courseDetails = [
    ["Start Date", course.startDate],
    ["Venue", ""],
    ["Group Name", ""],
    ["Module Title", course.moduleName],
    ["Module Code", course.courseId()],
  ]
  const courseDetailsTable = summarySheet.getRange(startRow,6, 5, 2).setValues(courseDetails);
  courseDetailsTable.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
  courseDetailsTable.setHorizontalAlignment("center");
  let detailsRow = startRow;
  for(i=0; i<5; i++){
    summarySheet.getRange(detailsRow, 7, 1, 2).merge();
    detailsRow++;
  }
  startRow += 7;

  //Student Results Table
  //Header
  summarySheet.getRange(startRow,1).setValue("Number");
  summarySheet.getRange(startRow,2).setValue("Name");
  summarySheet.getRange(startRow,3).setValue("");
  summarySheet.getRange(startRow,4).setValue("Project\nXX%");
  summarySheet.getRange(startRow,5).setValue("Exam\nXX%");
  summarySheet.getRange(startRow,6).setValue("Skills Demo\nXX%");
  summarySheet.getRange(startRow,7).setValue("Total");
  summarySheet.getRange(startRow,8).setValue("Grade");
  const headerRow = summarySheet.getRange(startRow,1, 1, 8);
  headerRow.setFontWeight("bold");
  headerRow.setHorizontalAlignment("center");
  headerRow.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
  summarySheet.getRange(startRow,2, 1, 2).merge();

  startRow++;

  //Student rows
  let number = 1;
  //figure out if we're using the old studentDetails or the new getLearners()
  if(course.studentDetails == null){
    course.studentDetails = course.getLearners();
  }
  for(student of course.studentDetails){
    summarySheet.getRange(startRow,1).setValue(number);
    summarySheet.getRange(startRow,2).setValue(student.getName());
    let studentRange = summarySheet.getRange(startRow,1, 1, 8);
    studentRange.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
    // Check if the row number is even (0-indexed)
    if (startRow % 2 === 0) {
        studentRange.setBackground("#DDE5ED"); // Set background for the entire row
    }
    summarySheet.getRange(startRow,2, 1, 2).merge();
    startRow++;
    number++;
  }
  summarySheet.getRange(startRow,1).setValue("Tutor:");
  summarySheet.getRange(startRow,2).setValue(course.tutorName);
  let tutorRange = summarySheet.getRange(startRow,1, 1, 8);
  tutorRange.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
  tutorRange.setFontWeight("bold");
  summarySheet.getRange(startRow,2, 1, 2).merge();

}