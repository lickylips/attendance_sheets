/**
 * createSignInSheet
 * Function to build a printable sign in sheet for use in class
 * @param {string} docId 
 * @param {object} course 
 */
function createSignInSheet(docId, course){
    const ss = SpreadsheetApp.openById(docId);
    Logger.log("Building Printable Sign In Sheet");
    const siSheet = ss.insertSheet("Printable Sign In Sheet");
    //Add Document Header
    
    let startRow = 2;
    //First Row / Title
    siSheet.getRange(startRow,1).setValue("Sign In Sheet")
                          .setFontSize(20)
                          .setHorizontalAlignment("center");
    siSheet.getRange(startRow,1,1,10).merge();
    startRow++;
  
    //Logo
    const logoUrl = "https://lickylip.net/wp-content/uploads/2023/09/21-small.png";
    const image = SpreadsheetApp.newCellImage()
                                .setSourceUrl(logoUrl)
                                .build();
    siSheet.getRange(startRow,1).setValue(image)
                                 .setHorizontalAlignment("center");
    siSheet.setRowHeight(startRow, 100);
    siSheet.getRange(startRow,1, 1, 10).merge();
    startRow++
  
    //Title, Date, QQI Code
    siSheet.getRange(startRow,1).setValue("Course Title: ");
    siSheet.getRange(startRow,2).setValue(course.moduleName);
    siSheet.getRange(startRow,2,1,3).merge();
    siSheet.getRange(startRow,6).setValue("QQI Code: ");
    siSheet.getRange(startRow,7).setValue(course.courseId());
    siSheet.getRange(startRow,7,1,3).merge();
  
    startRow++;
    //Start date and tutor name & signature
    siSheet.getRange(startRow,1).setValue("Start Date:");
    let startDate = new Date(course.startDate);
    let formattedDate = Utilities.formatDate(startDate, "GMT", "dd/MM/yyyy");
    siSheet.getRange(startRow,2).setValue(formattedDate)
           .setHorizontalAlignment("left");
    siSheet.getRange(startRow,2,1,3).merge();
    siSheet.getRange(startRow,6).setValue("Tutor Signature:")
    siSheet.getRange(startRow,7).setValue("___________________");
    siSheet.getRange(startRow,7,1,3).merge();
    startRow++;
    
    siSheet.getRange(startRow,7).setValue(course.moduleName);
    siSheet.getRange(startRow,7,1,3).merge();
    startRow++;
  
    //Instruction on block caps
    siSheet.getRange(startRow,1).setValue("PLEASE WRITE FULL NAME IN BLOCK CAPITALS")
            .setHorizontalAlignment("center");
    siSheet.getRange(startRow,1,1,10).merge();
    startRow++;
  
    //Clear boarders first
    siSheet.getRange(1,1,siSheet.getLastRow(),siSheet.getLastColumn())
            .setBorder(false, false, false, false, false, false);
  
    //Column Headers
    const borderStyle = SpreadsheetApp.BorderStyle.SOLID;
    const borderColor = "#4B3A71";
    const borderWidth = 1;
    siSheet.getRange(startRow,1).setValue("Number")
                                 .setHorizontalAlignment("center")
                                 .setWrap(true)
                                 .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    siSheet.getRange(startRow,2).setValue("Student")
                                 .setHorizontalAlignment("center")
                                 .setWrap(true)
                                 .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
    Logger.log("Adding headders for "+course.sessions()+" sessions");
    let sessionNo = 1;
    let col = 3;
    for (let i=0; i<Number(course.sessions()); i++){
      Logger.log("adding session "+sessionNo);
      siSheet.getRange(startRow, col).setValue("Session "+sessionNo)
                                    .setHorizontalAlignment("center")
                                    .setWrap(true)
                                    .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
      col+=1;
      sessionNo+=1;
    }
  
    startRow++;
    //Student rows
    //figure out if we're using the old studentDetails or the new getLearners()
    if(course.studentDetails == null){
        course.studentDetails = course.getLearners();
      }
    for(let i=0; i<course.studentDetails.length; i++){
      siSheet.getRange(startRow, 1).setValue(i+1)
                                    .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
      siSheet.getRange(startRow, 2).setValue(course.studentDetails[i].name)
                                    .setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID);
      siSheet.getRange(startRow,3,1,course.sessions()).setBorder(true, true, true, true, true, true, "#4B3A71", SpreadsheetApp.BorderStyle.SOLID)
      startRow++;
    }
    siSheet.autoResizeColumn(2);
    siSheet.getRange(1,1,siSheet.getLastRow(),siSheet.getLastColumn())
            .setFontColor("#4B3A71");
    siSheet.setHiddenGridlines(true);
  }
  