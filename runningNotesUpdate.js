// Purpose: 
// To automate the process of extracting specific information from Google Docs and organizing it in tables within another Google Document. The script has two primary functions:
// 1. getMonthly(): This function retrieves information from monthly meeting notes files stored in a secondary Google Drive folder. It searches for a specific target phrase within the files, extracts relevant data such as dates and text, and then updates a target Google Document with this information in a table format.
// 2. getTitles(): This function retrieves the titles and URLs of project folders from a primary Google Drive folder. It looks for a specific target phrase in the folder names, extracts the date from the folder titles, and then updates a target Google Document with this data, sorting the information in a table based on the dates.
// The script helps streamline the process of aggregating and organizing important information from various files and folders into a centralized and structured document, making it easier for users to access and review the required data.

// To note: 
// This script is developed as a Google Apps Script library script: i.e. a script that is not bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This standalone script project contains reusable code and functions and can be shared and included in multiple other scripts, allowing developers to easily reuse code across different projects.

//To use:
// How to use a library script: 
// In the script editor, click on "File" > "Project properties."
// In the "Script properties" tab, you will find the "Script ID." Copy this ID.
// To use the library in another script project, open the script editor of that project.
// Click on "Resources" > "Libraries."
// In the "Add a library" section, paste the Script ID and click "Add."
// Choose the version of the library you want to use (usually, you'll want to use the latest version).
// Set the identifier, which is the name you will use to reference the library functions in your main script project.
// Click "Save."
// After adding the library, you can use its functions in your main script project by calling them with the specified identifier. This way, you can take advantage of the shared code and easily maintain and update common functionalities across multiple projects.

/////////////////////////////////////////////

// function to create custom menu with buttons
function customMenu() {
  DocumentApp.getUi() // 
      .createMenu('Update Info')
      .addItem('Update Monthly and Project Meetings','runBothActionItems')      
      .addToUi();
}

//Global variables
var PRIMARY_FOLDER_ID ='13zl2CvMNtDMFKcNZetAA00e5tkh3Eo_M';// SNWG MO Meeting Notes folder
var SECONDARY_FOLDER_ID ='1l8gVfZqse5AWTfbVxVCwRcL84hf2XBc7'; // SNWG MO Monthly Project Status Updates folder

// helper funtion to return the ID of the active document
function getActiveDocumentId(){
  var getActiveDocumentId = DocumentApp.getActiveDocument().getId();
  return getActiveDocumentId;
}

// helper funtion to access a folder by its ID
function accessFolder(folderId){
  var folder = DriveApp.getFolderById(folderId);
  return folder;
}

// helper funtion to retrieve all files in the folder and its subfolders
function getFilesInFolder(folder) {
  var files = [];
  var fileIterator = folder.getFiles();
  var folderIterator = folder.getFolders();

  while (fileIterator.hasNext()) {
    var file = fileIterator.next();
    files.push(file);
  }

  while (folderIterator.hasNext()) {
    var subfolder = folderIterator.next();
    var subfolderFiles = getFilesInFolder(subfolder);
    files = files.concat(subfolderFiles);  
  }
  return files;
}

// helper funtion to extract date from document title 
function extractDateFromTitle(title) {
  var datePattern = /^\d{4}-\d{2}-\d{2}/;
  var match = title.match(datePattern);

  if (match) {
    return match[0];
  } else {
    return null;
  }
}

// helper funtion to find target phrase in the table in the monthly notes and return the text in cell 2 if found
function findInTable(document, TARGET_PHRASE){
  var body = document.getBody();
  var tables = body.getTables();

  if (tables.length > 0){
    var table = tables[0];
    var numRows = table.getNumRows();

    for (var i = 0; i < numRows; i++) {
      var row = table.getRow(i);
      var cell1 = row.getCell(0);
      var cell1Text = cell1.getText();

      if (cell1Text.includes(TARGET_PHRASE)) {
        var cell2 = row.getCell(1);
        var cell2Text = cell2.getText();
        return cell2Text;
      }
    }
  }
  return null;
}

// helper function to update the target document for the getMonthly function with the provided data (date, title, URL)
function updateTargetDocumentForGetMonthly(TARGET_PHRASE, date, title, url) {
  var TARGET_DOCUMENT_ID = getActiveDocumentId(); 
  var targetDocument = DocumentApp.openById(TARGET_DOCUMENT_ID);
  var body = targetDocument.getBody();
  var tables = body.getTables();

  if (tables.length > 0) {
    var table = tables[0];
    var numRows = table.getNumRows();

    for (var i = 0; i < numRows; i++) {
      var row = table.getRow(i);
      var cell1 = row.getCell(0);
      var cell1Url = cell1.getLinkUrl();
      var cell2 = row.getCell(1);
      var cell2Url = cell2.getLinkUrl();

      // If cell 1 or cell 2 contains the same URL as the source document, a matching entry is found, so stop and return.
      if (cell1Url === url || cell2Url === url) {
        return;
      }
    }

    var newRow = table.insertTableRow(0);

    var newCell1 = newRow.appendTableCell();
    newCell1.setText(date).setLinkUrl(url);  // Apply hyperlink to cell 1
    formatCell(newCell1);

    var newCell2 = newRow.appendTableCell();
    newCell2.setText(title);
    formatCell(newCell2);
  }
}

// helper function to update the target document for the getTitles function with the provided data (date, title, URL)
function updateTargetDocumentForGetTitles(TARGET_PHRASE, date, title, url) {
  var TARGET_DOCUMENT_ID = getActiveDocumentId(); 
  var targetDocument = DocumentApp.openById(TARGET_DOCUMENT_ID);
  var body = targetDocument.getBody();
  var tables = body.getTables();

  if (tables.length > 0) {
    var table = tables[0];
    var numRows = table.getNumRows();

    for (var i = 0; i < numRows; i++) {
      var row = table.getRow(i);
      var cell1 = row.getCell(0);
      var cell1Url = cell1.getLinkUrl();
      var cell2 = row.getCell(1);
      var cell2Url = cell2.getLinkUrl();

      // If cell 1 or cell 2 contains the same URL as the source document, a matching entry is found, so stop and return.
      if (cell1Url === url || cell2Url === url) {
        return;
      }
    }

    var newRow = table.insertTableRow(0);

    var newCell1 = newRow.appendTableCell();
    newCell1.setText(date);
    formatCell(newCell1);

    var newCell2 = newRow.appendTableCell();
    newCell2.setText(title).setLinkUrl(url); // Apply hyperlink to cell 2
    formatCell(newCell2);
  }
}

// helper function to format new cells in Running Notes table
function formatCell(cell) {
  var fontSize = 11;
  var fontFamily = 'Raleway';
  var backGroundColor = '#FFFFFF';
  
  cell.setFontSize(fontSize);
  cell.setFontFamily(fontFamily);
  cell.setBackgroundColor(backGroundColor);
}

// helper function to sort newly created rows in descending order
function sortTable(document, newRowCount) {
  var body = document.getBody();
  var tables = body.getTables();

  if (tables.length > 0) {
    var table = tables[0];
    var data = [];

    for (var i = 0; i < newRowCount; i++) {
      var row = table.getRow(i);
      var cell1 = row.getCell(0);
      var cell2 = row.getCell(1);
      data.push([cell1.getText(), cell2.getText(), cell1.getLinkUrl()]);

      table.removeRow(i);
    }

    data.sort(function(a, b) {
      return b[0].localeCompare(a[0]);
    });

    for (var i = 0; i < newRowCount; i++) {
      var newRow = table.insertTableRow(i);
      var newCell1 = newRow.appendTableCell(data[i][0]).setLinkUrl(data[i][2]);
      formatCell(newCell1);  // Apply formatting to cell 1
      var newCell2 = newRow.appendTableCell(data[i][1]);
      formatCell(newCell2);  // Apply formatting to cell 2
    }
  }
}

// Primary function to get paragraph updates from monthly meeting notes
function getMonthly() {
  var TARGET_DOCUMENT_ID = getActiveDocumentId();
  var secondaryFolder = accessFolder(SECONDARY_FOLDER_ID);
  var files = getFilesInFolder(secondaryFolder);

  var twoYearsAgo = new Date();
  twoYearsAgo.setFullYear(twoYearsAgo.getFullYear() - 2);
  
  var targetDocument = DocumentApp.openById(TARGET_DOCUMENT_ID);
  var tables = targetDocument.getBody().getTables();
  var initialRowCount = tables[0].getNumRows();
  
  files.forEach(function(file) {
    // Skip "Template" files and move to the next.
    if (file.getName().includes("Template")){
      return;
    }
    // Search only files updated in last two years
    if (file.getLastUpdated() >= twoYearsAgo) {
      if (file.getMimeType() === 'application/vnd.google-apps.document') {
        var doc = DocumentApp.openById(file.getId());
        var text = findInTable(doc, TARGET_PHRASE);
        if (text !== null) {
          var documentTitle = file.getName();
          var documentUrl = "https://docs.google.com/document/d/" + file.getId();
          updateTargetDocumentForGetMonthly(TARGET_PHRASE, documentTitle, text, documentUrl);           
        }
      }
    }
  });
  
  var finalRowCount = tables[0].getNumRows();
  var newRowCount = finalRowCount - initialRowCount;
  sortTable(targetDocument, newRowCount);
}

// Primary function to get agenda title links from project folders
function getTitles() {
  // var TARGET_DOCUMENT_ID = getActiveDocumentId(); 

  var primaryFolder = accessFolder(PRIMARY_FOLDER_ID);
  var files = getFilesInFolder(primaryFolder);

  // Create an array to store new rows
  var newRows = [];

  files.forEach(function(file) {

    // Skip "Template" files and move to the next.
    if (file.getName().includes("Template")){
      return;
    }

    if (file.getName().includes(TARGET_PHRASE)) {
      var date = extractDateFromTitle(file.getName());
      if (date === null) {
        date = "Date Not Found";
      }

      // Instead of adding a new row to the document here, add the row data to newRows
      newRows.push({
        date: date,
        title: file.getName(),
        url: "https://docs.google.com/document/d/" + file.getId(),
      });
    }
  });

  // Sort newRows by date in descending order (newest first)
  newRows.sort(function(a, b) {
    return a.date.localeCompare(b.date);
  });

  // Add the sorted rows to the document
  newRows.forEach(function(row) {
    updateTargetDocumentForGetTitles(TARGET_PHRASE, row.date, row.title, row.url);
  });
}

//Secondary function to run both primary functions
function runBothActionItems() {
  getMonthly();
  getTitles();
}
