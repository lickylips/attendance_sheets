function generateLabels(){
    const lableTemplateId = "1KNplm1GPPowrGvEvASLC7os_fBT8rQjrvLF-ORsP-LM";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = ss.getSheetByName("Document Generator");
    const studentData = studentSheet.getDataRange().getValues();
    const students = getStudentObjects(studentData);
    const studentNumbers = students.length;
    const lableTemplate = DriveApp.getFileById(lableTemplateId);
    const newlableDoc = lableTemplate.makeCopy();
    const folder = DriveApp.getFileById(ss.getId()).getParents().next();
    newlableDoc.setName("Labels");
    newlableDoc.moveTo(folder);
    const newlableDocId = newlableDoc.getId();
    const lableDoc = DocumentApp.openById(newlableDocId);
    const body = lableDoc.getBody();
    const tables = body.getTables();
    let studentNumber = 0;
    // Iterate through each table found in the document
    for (let i = 0; i < tables.length; i++) {
        const table = tables[i];

        // Iterate through each row of the table
        for (let j = 0; j < table.getNumRows(); j++) {
            const row = table.getRow(j);

            // Iterate through each cell in the row
            for (let k = 0; k < row.getNumCells(); k++) {
                if (studentNumber < studentNumbers) {
                    let student = students[studentNumber];
                    const cell = row.getCell(k);
                    let cellText = cell.getText();

                    // Your placeholder and replacement logic:
                    if (cellText.includes("{{NAME}}")) {
                    cellText = cellText.replace("{{NAME}}", student.name);
                    studentNumber++; 
                    }
                    if (cellText.includes("{{ADDRESS}}")) {
                    cellText = cellText.replace("{{ADDRESS}}", student.address); 
                    }

                    // Update the cell's text
                    const textElement = cell.editAsText(); // Get text element for modifications
                    textElement.setText(cellText);        
                    const paragraphs = cell.getNumChildren();
                    for(let p = 0; p < paragraphs; p++){
                        const paragraph = cell.getChild(p).asParagraph();
                        paragraph.setFontFamily('Poppins');
                        paragraph.setFontSize(12);
                        paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
                    }
                } else {
                    // Delete the cell if there are no more students
                    let cell = row.getCell(k);
                    cell.setText("");
                }
            }
        }
    }
}

function getStudentObjects(data){
    //get header indexes
    for(i=0; i<data[0].length; i++){
        if(data[0][i].includes("Name")){nameIndex = Number(i);}
        if(data[0][i].includes("Email")){emailIndex = Number(i);}
        if(data[0][i].includes("Sponsor")){sponsorIndex = Number(i);}
        if(data[0][i].includes("Tutor")){tutorIndex = Number(i);}
        if(data[0][i].includes("Date")){dateIndex = Number(i);}
        if(data[0][i].includes("Paid")){paidIndex = Number(i);}
        if(data[0][i].includes("Course Passed")){passedIndex = Number(i);}
        if(data[0][i].includes("Sent")){sentIndex = Number(i);}
        if(data[0][i].includes("Cert")){certIndex = Number(i);}
        if(data[0][i].includes("Letter")){letterIndex = Number(i);}
        if(data[0][i].includes("Address")){addressIndex = Number(i);}
        if(data[0][i].includes("Phone")){phoneIndex = Number(i);}
    }
    data.shift()//remove header row
    const studentArray = [];
    for(row of data){
        student = {
            name: row[nameIndex],
            email: row[emailIndex],
            sponsor: row[sponsorIndex],
            tutor: row[tutorIndex],
            date: row[dateIndex],
            paid: row[paidIndex],
            passed: row[passedIndex],
            sent: row[sentIndex],
            cert: row[certIndex],
            letter: row[letterIndex],
            address: row[addressIndex],
            phone: row[phoneIndex]
        }
        studentArray.push(student);
    }
    return studentArray;
}