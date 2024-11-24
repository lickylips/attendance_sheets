function getSheetsInFolder(folderId) {
    // Create an array to store the sheet files
    const sheetFiles = [];
  
    // Get the folder by ID
    const folder = DriveApp.getFolderById(folderId);
  
    // Get all files in the folder
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
    // Iterate over the files and add them to the array if they meet the criteria
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      if (fileName.indexOf("Deprecated") === -1) {
        sheetFiles.push(file);
      }
    }
  
    // Get all subfolders in the folder
    const subFolders = folder.getFolders();
  
    // Iterate over the subfolders and recursively call the function
    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      const subFolderName = subFolder.getName();
      // Exclude the "Automation Code" folder
      if (subFolderName !== "Automation Code") {
        sheetFiles.push(...getSheetsInFolder(subFolder.getId()));
      }
    }
  
    // Return the array of sheet files
    return sheetFiles;
  }

  function testGetSheetsInFolder() {
    const folderId = "1S4OWYJNRCEev0e9IxLuQO1v6DcvQsu2i";
    const sheetFiles = getSheetsInFolder(folderId);
    Logger.log(sheetFiles);
  }