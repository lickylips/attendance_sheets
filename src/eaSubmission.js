function createLinkedFolder() {
    // Get the parent folder ID (replace with your actual folder ID)
    const parentFolderId = '1_fNaRPPCYb3ZBnEP_AG8934Q3okQHMbc'; 
    const parentFolder = DriveApp.getFolderById(parentFolderId);
  
    // Get a list of all subfolders within the parent folder
    const subfolders = parentFolder.getFolders();
  
    // Create an array to store subfolder names
    const folderNames = [];
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      folderNames.push(folder.getName());
    }
  
    // Display a prompt to the user to select a subfolder
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Select a Folder',
      'Please enter the number of the folder you want to link to:\n' +
      folderNames.map((name, index) => `${index + 1}. ${name}`).join('\n'),
      ui.ButtonSet.OK_CANCEL
    );
  
    // Check if the user clicked OK
    if (response.getSelectedButton() == ui.Button.OK) {
      // Parse the user's input to get the selected folder index
      const selectedIndex = parseInt(response.getResponseText()) - 1;
  
      // Validate the user's input
      if (selectedIndex >= 0 && selectedIndex < folderNames.length) {
        // Get the selected folder
        const selectedFolder = parentFolder.getFoldersByName(folderNames[selectedIndex]).next();
  
        // Get the ID of the folder you want to link to 
        let ss = SpreadsheetApp.getActiveSpreadsheet();
        let ssId = ss.getId()
        let settings = getSettings(ssId);
        let file = DriveApp.getFileById(ssId);
        let attendanceFolder = file.getParents().next();
  
        // **Check if a folder named settings.courseName exists within selectedFolder**
        let courseFolder = selectedFolder.getFoldersByName(settings.courseName);
        if (!courseFolder.hasNext()) {
          // **If it doesn't exist, create it**
          courseFolder = selectedFolder.createFolder(settings.courseName);
        } else {
          // **If it exists, get the folder**
          courseFolder = courseFolder.next();
        }
  
        // **Create the link in the courseFolder**
        courseFolder.createShortcut(attendanceFolder.getId());
  
        Logger.log('Link created successfully!');
      } else {
        ui.alert('Invalid folder number.');
      }
    }
  }