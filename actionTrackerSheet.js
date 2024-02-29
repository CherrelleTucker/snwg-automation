/* Purpose: 
1. Search agendas for action items to be completed and populate in the Action Tracking Google Sheet.
2. Push status updates from from the Action Tracking Google sheet to the Action source agendas as changed.

Future development: 
Preserve links in task items from the source documents. Will first require the same development in the inDocActionItems script, as the links are lost in that action first. 

To note: 
This script is developed as a Google Apps Script container script: i.e. a script that is bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This container script acts as the file's custom script, allowing users to extend the functionality of the file by adding custom functions, triggers, and menu items to enhance its behavior and automation.

Instructions for Using this Script in your container file:
1. Open a new or existing Google Sheets file where you want to use the script.
2. Click on "Extensions" in the top menu and select "Apps Script" from the dropdown menu. This will open the Google Apps Script editor in a new tab.
3. Copy and paste the provided script into the Apps Script editor, replacing the existing code (if any).
4. In the TablePullPopulate function, replace the placeholder values for folderId and spreadsheetId with the actual IDs of your Google Drive folder containing the agendas and the Google Sheets spreadsheet where you want to track the actions, respectively.
5. In the updateStatus function, replace the placeholder value for spreadsheetId with the actual ID of your Google Sheets spreadsheet.
6. Save the script by clicking on the floppy disk icon or by pressing "Ctrl + S" (Windows) or "Cmd + S" (Mac).
7. Go back to your Google Sheets file, refresh the page, and you'll see a new custom menu labeled "Action Items" in the top menu.
8. Click on "Action Items" in the top menu to access the custom menu. You'll find two options:
  a. "Get Actions from Agendas": This option will pull actions from the specified agendas and populate them in the "Table Pull" sheet in your Google Sheets file.
  b. "Update Status in Source Document": This option will push status updates from the "Table Pull" SEPack to the corresponding action items in the source agendas.
9. Whenever you want to get actions from the agendas or update status in the source documents, simply click on the corresponding option from the "Action Items" menu.
Note: Make sure to properly set up the correct folder structure in Google Drive and name your agendas and sheets according to the script's logic for pulling and updating actions. This documentation assumes you have some basic familiarity with Google Apps Script and how to run container-bound scripts within a Google Sheets file. */ 


///////////////////Custom Menu//////////////////////////////////////////////

// Add a custom menu to the spreadsheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Action Items')
    .addItem('Pull actions for ALL folders', 'pullActionsForAllFolders')
    .addItem('Pull actions for MO', 'pullActionsForFolderMO')
    .addItem('Pull actions for SEP', 'pullActionsForFolderSEP')
    .addItem('Pull actions for DevSeed', 'pullActionsForFolderDevSeed')
    .addItem('Pull actions for AssessmentHQ','pullActionsForFolderAssessmentHQ')
    .addSeparator()
    .addItem('Push actions from ALL tabs', 'pushActionsFromAllTabs') 
    .addItem('Push actions from MO', 'pushActionsFromTabMO') 
    .addItem('Push actions from SEP', 'pushActionsFromTabSEP') 
    .addItem('Push actions from DevSeed', 'pushActionsFromTabDevSeed') 
    .addItem('Push actions from Assessment HQ', 'pushActionsFromTabAssessmentHQ')
    .addToUi();
}

///////////////////// Global variables//////////////////////////////////

var spreadsheetId = 'xxxxxxxxxxxxxxxxxxxxx'; // Action Tracking Sheet
var folderIds = {
MO: 'xxxxxxxxxxxxxxxxxxxxx', // SNWG Management Office Weekly Internal Planning Meeting folder 
SEP: 'xxxxxxxxxxxxxxxxxxxxx', // Stakeholder Engagement Program Weekly Tag Up folder
DevSeed: 'xxxxxxxxxxxxxxxxxxxxx',// Assessment/DevSeed Weekly Tag up Folder
AssessmentHQ: 'xxxxxxxxxxxxxxxxxxxxx' // Assessment HQ CY2024 Folder Id  
};
var originalFolderId = 'xxxxxxxxxxxxxxxxxxxxx'; // AssessmentHQ folder that contains .docx files (Katrina prefers to work offline in MS Word, then upload to Drive)
var conversionFolderId = 'xxxxxxxxxxxxxxxxxxxxx-XJhQ10'; // Testing Folder ID for file conversion

///////////////////Pull///////////////////////////////////////////////

// Function to pull actions for all folders
function pullActionsForAllFolders() {
logExecutionTime(function() {
  pullActionsForFolder('MO');
  pullActionsForFolder('SEP');
  pullActionsForFolder('DevSeed');
  pullActionsForFolder('AssessmentHQ');
}, 'pullActionsForAllFolders');
}

// Function to pull actions for MO folder
function pullActionsForFolderMO() {
logExecutionTime(function() {
  pullActionsForFolder('MO');
}, 'pullActionsForFolderMO');
}

// Function to pull actions for SEP folder
function pullActionsForFolderSEP() {
logExecutionTime(function() {
  pullActionsForFolder('SEP');
}, 'pullActionsForFolderSEP');
}

// Function to pull actions for DevSeed folder
function pullActionsForFolderDevSeed() {
logExecutionTime(function() {
  pullActionsForFolder('DevSeed');
}, 'pullActionsForFolderDevSeed');
}

// Function to pull actions for AssessmentHQ folder
function pullActionsForFolderAssessmentHQ() {
logExecutionTime(function() {
  pullActionsForFolder('AssessmentHQ');
}, 'pullActionsForFolderAssessmentHQ'); 
}

////////////////////////////////////////////////////////

// Pull function: convert a Microsoft Word uploaded to Google Drive to a Google Doc file for ineraction with Google Apps Scripts. Due to Google Drive permission shenanigans that make sense to neither me nor ChatGPT, the original docx file is copied to a folder for conversion and then converted to a Google Doc. Both the original and duplicated .docx files are deleted.  
function convertDocxToGoogleDoc(originalFileId, conversionFolderId, originalFolderId) {
var originalFile = DriveApp.getFileById(originalFileId); 
var conversionFolder = DriveApp.getFolderById(conversionFolderId); 
var originalFolder = DriveApp.getFolderById(originalFolderId); 
// Make a copy of the .docx file in the conversion folder
var copiedFile = originalFile.makeCopy(originalFile.getName(), conversionFolder);

// Set up options for the API call
var options = {
  method: "POST",
  contentType: "application/json",
  payload: JSON.stringify({
    title: copiedFile.getName(),
    mimeType: MimeType.GOOGLE_DOCS
  }),
  headers: {
    Authorization: "Bearer " + ScriptApp.getOAuthToken()
  },
  muteHttpExceptions: true
};

try {
  // Make the API call to convert the copied file
  var response = UrlFetchApp.fetch("https://www.googleapis.com/drive/v2/files/" + copiedFile.getId() + "/copy", options);
  var jsonResponse = JSON.parse(response.getContentText());
  var googleDocId = jsonResponse.id;

  // Move the converted Google Doc back to the original folder
  var convertedDoc = DriveApp.getFileById(googleDocId);
  moveFileToFolder(convertedDoc, originalFolder); 

  // Delete the original .docx file in the destination folder
  originalFile.setTrashed(true);

  // Also delete the copied .docx file in the conversion folder
  copiedFile.setTrashed(true);

  return { id: googleDocId, url: jsonResponse.alternateLink };
} catch (e) {
  Logger.log("Error during conversion: " + e.toString());
  return null;
}
}

// pull function: move the newly created Google Doc file back to the destination folder 
function moveFileToFolder(file, destinationFolder) {
var parents = file.getParents();
while (parents.hasNext()) {
  var parent = parents.next();
  parent.removeFile(file);
}
destinationFolder.addFile(file);
}

// Pull function: pull actions from a specific meeting agendas Google Drive folder
function pullActionsForFolder(folderName) {
var folderId = folderIds[folderName];
var tablePullSheetName = folderName;

Logger.log('Step 1: Pulling actions from documents...');
var actions = pullActionsFromDocuments(folderId);
Logger.log('Step 1: Actions retrieved:', actions);

Logger.log('Step 2: Populating sheet with actions...');
populateSheetWithActions(spreadsheetId, tablePullSheetName, actions);
Logger.log('Step 2: Sheet populated with actions.');
}

// Pull function: Get table headers
function getTableHeaders(table) {
var headers = [];
var row1 = table.getRow(0);

for (var col = 0; col < row1.getNumCells(); col++) {
  headers.push(row1.getCell(col).getText().trim());
}

return headers;
}

// Pull function: Check if the table has the required headers
function hasRequiredHeaders(headers) {
// Check if the table has "Status" "Owner" and "Action" or "Who" "What" "Status"
var hasStatus = headers.includes('Status') || headers.includes('What'); // Column A header in document table
var hasOwner = headers.includes('Owner') || headers.includes('Who'); // Column B header in document table
var hasAction = headers.includes('Action') || headers.includes('Status'); // Column C header in document table

return hasStatus && hasOwner && hasAction;
}

// Pull function: Filter table data based on headers
function filterTableData(tableData, headers) {
var filteredData = [];

for (var i = 0; i < tableData.length; i++) {
  var rowData = tableData[i];
  var filteredRow = ['', '', '']; // Initialize with empty values

  // Check if the row is completely empty
  if (rowData.join('').trim() === '') {
    continue; // Skip empty rows
  }

  for (var j = 0; j < headers.length; j++) {
    var header = headers[j];
    var columnIndex = headers.indexOf(header);

    if (columnIndex !== -1) {
      // Set values in corresponding columns
      switch (header.toLowerCase()) {
        case 'status':
          filteredRow[0] = rowData[columnIndex];
          break;
        case 'owner':
        case 'who':
          filteredRow[1] = rowData[columnIndex];
          break;
        case 'action':
        case 'what':
          filteredRow[2] = rowData[columnIndex];
          break;
      }
    }
  }

  Logger.log('Filtered Row:', filteredRow); // Add this line for debugging
  filteredData.push(filteredRow);
}

return filteredData;
}

// Pull function: Pull actions from documents
function pullActionsFromDocuments(folderId) {
var folder = DriveApp.getFolderById(folderId);
var files = folder.getFiles();
var allActions = [];

while (files.hasNext()) {
  var file = files.next();
  var fileName = file.getName();
  var fileId = file.getId();
  var docId, documentLink;

  // Log the file name and ID for debugging
  Logger.log("Processing file: " + fileName + " with ID: " + fileId);

  // Determine the document ID and link based on file type
  if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
    docId = file.getId();
    documentLink = '=HYPERLINK("' + file.getUrl() + '", "' + file.getName() + '")';

  } else if (file.getMimeType() === MimeType.MICROSOFT_WORD || file.getName().toLowerCase().endsWith(".docx")) {
    var conversionResult = convertDocxToGoogleDoc(fileId, conversionFolderId,folderId);
    // Convert .docx file to Google Doc and wait for the process to complete
    
    if (conversionResult && conversionResult.id) {
    docId = conversionResult.id;
    documentLink = '=HYPERLINK("' + conversionResult.url + '", "' + file.getName() + '")';
  } else {
    // Handle the case where conversion fails
    Logger.log("Failed to convert .docx to Google Doc for file: " + file.getName());
    continue; // Skip unsupported file types
  }
}  

  try {
    var doc = DocumentApp.openById(docId);
    var tables = doc.getBody().getTables();

    for (var i = 0; i < tables.length; i++) {
      var table = tables[i];
      var headers = getTableHeaders(table);

      if (hasRequiredHeaders(headers)) {
        var tableData = tableTo2DArray(table);
        var filteredTableData = filterTableData(tableData, headers);

        for (var j = 0; j < filteredTableData.length; j++) {
          var row = [documentLink].concat(filteredTableData[j]);
          allActions.push(row);
        }
      }
    }
  } catch (error) {
    console.error("Error processing document (ID: " + docId + "): " + error.message);
    // Skip this document and continue with the next
    continue;
  }
}
return allActions;
}

// Pull function: Find the second table with qualifying headers
function findSecondTable(tables) {
for (var i = 0; i < tables.length; i++) {
  var table = tables[i];
  var row1 = table.getRow(0);

  var statusCell = row1.getCell(0);
  var ownerCell = row1.getCell(1);
  var actionCell = row1.getCell(2);

  if (
    statusCell.getText().toLowerCase() === 'status' &&
    ownerCell.getText().toLowerCase() === 'owner' &&
    actionCell.getText().toLowerCase() === 'action'
  ) {
    return table;
  }
}

return null;
}

// Pull function: Convert table to 2D array for convenience and flexibility in data processing
function tableTo2DArray(table) {
var numRows = table.getNumRows();
var numCols = table.getRow(0).getNumCells();
var data = [];

for (var i = 1; i < numRows; i++) {
  var rowData = [];
  for (var j = 0; j < numCols; j++) {
    var cellValue = table.getCell(i, j).getText();
    rowData.push(cellValue);
  }
  data.push(rowData);
}

return data;
}

// Pull function: Populate the Sheet with actions
function populateSheetWithActions(spreadsheetId, sheetName, actions) {
var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
var sheet = spreadsheet.getSheetByName(sheetName);

if (!sheet) {
  sheet = spreadsheet.insertSheet(sheetName);
} else {
  // Clear only the contents of columns A to E
  sheet.getRange(1, 1, sheet.getMaxRows(), 5).clearContent();
}

var headerRow = ['Action Source', 'Status', 'Assigned to', 'Task'];
actions.unshift(headerRow);

var numRows = actions.length;
var numCols = headerRow.length;

if (numRows > 0 && numCols > 0) {
  // Set values starting from cell A1
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(actions);
}
}

////////////////// Push ////////////////////////////////////////////////////////

// Function to call the primary function on all named tabs
function pushActionsFromAllTabs() {
logExecutionTime(function() {
  updateStatusOnTab("MO");
  updateStatusOnTab("SEP");
  updateStatusOnTab("DevSeed");
  updateStatusOnTab("AssessmentHQ");
}, 'pushActionsFromAllTabs');
}

// Function to call the primary function on the "MO" tab
function pushActionsFromTabMO() {
logExecutionTime(function() {
  updateStatusOnTab("MO");
}, 'pushActionsFromTabMO');
}

// Function to call the primary function on the "SEP" tab
function pushActionsFromTabSEP() {
logExecutionTime(function() {
  updateStatusOnTab("SEP");
}, 'pushActionsFromTabSEP');
}

// Function to call the primary function on the "DevSeed" tab
function pushActionsFromTabDevSeed() {
logExecutionTime(function() {
  updateStatusOnTab("DevSeed");
}, 'pushActionsFromTabDevSeed');
}

// Function to call the primary function on the "AssessmentHQ" tab
function pushActionsFromTabAssessmentHQ() {
logExecutionTime(function() {
  updateStatusOnTab("AssessmentHQ");
}, 'pushActionsFromTabAssessmentHQ');
}

/////////////////////////////////////////

// Push function to update status in the source document based on data in a spreadsheet
function updateStatusOnTab(tabName) {
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
var range = sheet.getRange("A:G");
var values = range.getValues();

for (var i = 1; i < values.length; i++) {
  var actionSourceUrl = values[i][6];
  var status = values[i][1];
  var task = values[i][3];

  if (actionSourceUrl || status || task) {
    console.log("Processing task: " + task);
    updateStatusInSourceDoc(actionSourceUrl, status, task);
  }
}
}

// Push function to update the status in the source document
function updateStatusInSourceDoc(actionSource, status, task) {
var actionSourceUrl = actionSource;

if (!actionSourceUrl) {
  console.error("Invalid URL in Column G:", actionSource);
  return;
}

try {
  var sourceDoc = DocumentApp.openByUrl(actionSourceUrl);
} catch (error) {
  console.error("Error opening document by URL:", error);
  return;
}

  var body = sourceDoc.getBody();
var table = findTableByHeadings(body, [["Who", "What", "Status"], ["Status", "Owner", "Action"]]);

if (!table) {
  console.error("Table not found with specified headings.");
  return;
}

var taskColumnIndex = findColumnIndexByHeader(table, "What"); // or "Action"
var statusColumnIndex = findColumnIndexByHeader(table, "Status");

if (taskColumnIndex === -1 || statusColumnIndex === -1) {
  console.error("Required columns not found in the table.");
  return;
}

var rowIndex = findRowIndexByColumnValue(table, taskColumnIndex, task);

if (rowIndex !== -1) {
  table.getCell(rowIndex, statusColumnIndex).setText(status); // Update the status in the Status column
  console.log("Status Updated Successfully!");
} else {
  console.error("Row with task not found in the table.");
}
}

// Push function to get the values of a specific row in the table
function getTableRowValues(table, rowIndex) {
var row = table.getRow(rowIndex);
var numCells = row.getNumCells();
var rowValues = [];

for (var i = 0; i < numCells; i++) {
  rowValues.push(row.getCell(i).getText());
}

return rowValues;
}

// Push Function to find the column index in the table by header
function findColumnIndexByHeader(table, header) {
var headerRow = table.getRow(0);
var numCells = headerRow.getNumCells();

for (var i = 0; i < numCells; i++) {
  var cellText = headerRow.getCell(i).getText();
  if (cellText === header || (header === "What" && cellText === "Action")) {
    return i;
  }
}

console.error("Header not found: " + header);
return -1;
}

// Push function to find the table in the document with specified headings
function findTableByHeadings(body) {
var tables = body.getTables();
var possibleHeadingsSets = [
  ["Who", "What", "Status"],
  ["Status", "Owner", "Action"]
];

for (var i = 0; i < tables.length; i++) {
  var table = tables[i];
  var headerRow = table.getRow(0);
  var headers = [];

  // Get headers from the table
  for (var j = 0; j < headerRow.getNumCells(); j++) {
    headers.push(headerRow.getCell(j).getText().trim());
  }

  // Check if headers match any of the possible sets
  var isMatch = possibleHeadingsSets.some(set => 
    set.every((heading, index) => 
      headers[index] === heading
    )
  );

  if (isMatch) {
    return table;
  }
}

console.error("Table not found with specified headings.");
return null;
}

// Push function to find the row index in the table with a specific column value
function findRowIndexByColumnValue(table, columnName, value) {
var columnIndex = getColumnIndex(table, columnName);

console.log("Searching for task: " + value);

if (columnIndex !== -1) {
  var numRows = table.getNumRows();

  for (var i = 1; i < numRows; i++) {
    var cellValue = table.getCell(i, columnIndex).getText();

    // Log each cell value for investigation
    console.log("Row " + i + " Task found: " + cellValue);
    
    // Log the comparison result
    console.log("Comparison Result:", cellValue === value);

    if (cellValue === value) {
      console.log("Task matched at Row " + i);
      return i;
    }
  }
}

console.log("Task not found in the table.");
return -1;
}

// Function to get the index of a column in the table
function getColumnIndex(table, columnName) {
var headerRow = table.getRow(0);
var numCells = headerRow.getNumCells();

for (var i = 0; i < numCells; i++) {
  var cellText = headerRow.getCell(i).getText();
  if (cellText === columnName || cellText === "What" || cellText === "Action" || cellText === "Task") {
    console.log("Column found: " + cellText + " at index " + i)
    return i;
  }
}

console.log("Column " + columnName + " not found.");
return -1;
}

///////////////Combine Open Actions///////////////////////////////////////////////

/* Preferred method: 
Copy this formula to cell A1 of a new sheet to combine all open actions across the named tabs into one sheet. 
"={FILTER(MO!A2:D, MO!B2:B <> "done", MO!B2:B <> "Done", LEN(MO!A2:A) > 1, LEN(MO!A2:A) > 1); 
FILTER(DevSeed!A2:D, DevSeed!B2:B <> "done", DevSeed!B2:B <> "Done", LEN(DevSeed!A2:A) > 0, LEN(DevSeed!A2:A) > 0); 
FILTER(SEP!A2:D, SEP!B2:B <> "done", SEP!B2:B <> "Done", LEN(SEP!A2:A) > 0, LEN(SEP!A2:A) > 0)}"
*/

/* Script function to copy all rows of each sheet that do not contain "Done" in the "Status Column". Maintain column D formatting from source sheet to track length of time the action has been open. Use script function only if data becomes too large for in-sheet formula to work.
function copyDataToAllOpenSheet() {
// Get a reference to the currently open spreadsheet
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sourceSheetNames = ['MO', 'SEP'];
var targetSheetName = 'All Open (Sort Only - Do not Edit)';

var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
var targetSheet = spreadsheet.getSheetByName(targetSheetName);

// Create "All Open (Sort Only - Do not Edit)" sheet if it doesn't exist
if (!targetSheet) {
  targetSheet = spreadsheet.insertSheet(targetSheetName);
}

// Clear existing data and formatting if the sheet already exists
targetSheet.clear();

// Loop through source sheets and copy data to "All Open (Sort Only - Do not Edit)" sheet
sourceSheetNames.forEach(function (sourceSheetName) {
  var sourceSheet = spreadsheet.getSheetByName(sourceSheetName);

  // Copy data and formatting from source sheet to "All Open (Sort Only - Do not Edit)" sheet
  var sourceRange = sourceSheet.getDataRange();
  var numRows = sourceRange.getNumRows();
  var numCols = sourceRange.getNumColumns();

  var data = sourceRange.getValues();
  var richTextValues = sourceRange.getRichTextValues();
  var backgrounds = sourceRange.getBackgrounds();
  var fonts = sourceRange.getFontWeights();

  // Filter data based on the condition (Column B not containing 'done')
  var filteredData = data.filter(row => row[1].toLowerCase().indexOf('done') === -1);

  // Append filtered data to "All Open (Sort Only - Do not Edit)" sheet along with hyperlinks and formatting
  for (var rowIndex = 0; rowIndex < filteredData.length; rowIndex++) {
    var row = filteredData[rowIndex];
    var hyperlink = richTextValues[rowIndex][0].getLinkUrl();
    if (hyperlink) {
      row[0] = '=HYPERLINK("' + hyperlink + '","' + row[0] + '")';
    }

    // Append the row to "All Open (Sort Only - Do not Edit)" sheet
    targetSheet.appendRow(row);

    // Copy formatting from source sheet to "All Open (Sort Only - Do not Edit)" sheet for column D
    var targetRange = targetSheet.getRange(targetSheet.getLastRow(), 4, 1, 1);
    targetRange.setRichTextValues([[richTextValues[rowIndex][3]]]);
    targetRange.setBackgrounds([[backgrounds[rowIndex][3]]]);
    targetRange.setFontWeights([[fonts[rowIndex][3]]]);
  }
});
}*/

////////////////////////////Duplicate Action Update//////////////////////////////////

/*
Triggered when a cell in a Google Sheet is edited. This function updates the status of all tasks in the same sheet that match the task of the edited status cell.

Structure assumptions:
- Column A: Action Source
- Column B: Status
- Column C: Assigned To
- Column D: Task

@param {object} e The event object that contains information about the cell that was edited.
 */
function onEdit(e) {
  // Extract necessary details from the event object
  var range = e.range; // The range that was edited
  var sheet = range.getSheet(); // The sheet where the edit happened
  var editedRow = range.getRow(); // The row number of the edited cell
  var editedCol = range.getColumn(); // The column number of the edited cell

  var statusColumn = 2; // The column number where statuses are stored (B)
  var taskColumn = 4;  // The column number where tasks are stored (D)

  // Check if the edit was made in the status column
  if (editedCol === statusColumn) {
    var updatedStatus = e.value; // Get the new status value from the edited cell
    var task = sheet.getRange(editedRow, taskColumn).getValue(); // Retrieve the task associated with the edited status

    // Get all data in the sheet to search for matching tasks
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // Iterate through each row in the sheet
    values.forEach(function(row, index) {
      // Check if the task in the current row matches the edited task and is not the same row that was edited
      if (row[taskColumn - 1] === task && index !== (editedRow - 1)) {
        // Update the status of the matching task to the new status
        sheet.getRange(index + 1, statusColumn).setValue(updatedStatus);
      }
    });
  }
}



/////////////////////Testing Functions////////////////////////////////

// Testing function to verify Google Sheet ID, Google Sheet Name, Sheet tab names for all tabs
function getSpreadsheetInfoTest() {
// Get the active spreadsheet
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// Get the spreadsheet name
var spreadsheetName = spreadsheet.getName();

// Get the spreadsheet ID
var spreadsheetId = spreadsheet.getId();

// Get an array of sheet names
var sheetNames = spreadsheet.getSheets().map(function(sheet) {
  return sheet.getName();
});

// Log the information
Logger.log("Spreadsheet Name: " + spreadsheetName);
Logger.log("Spreadsheet ID: " + spreadsheetId);
Logger.log("Sheet Names: " + sheetNames.join(", "));
}

// Testing function for docx conversion
function testDocxConversion() {
var originalFolderId = 'xxxxxxxxxxxxxxxxxxxxx'; // AssessmentHQ Weekly FY 2024 folder
var conversionFolderId = 'xxxxxxxxxxxxxxxxxxxxx-xxxxxxx'; // Testing Folder

var originalFolder = DriveApp.getFolderById(originalFolderId);
var conversionFolder = DriveApp.getFolderById(conversionFolderId);

// Assuming you're testing with a specific .docx file in the original folder
var files = originalFolder.getFilesByType(MimeType.MICROSOFT_WORD);

if (files.hasNext()) {
  var file = files.next();
  var fileName = file.getName();

  // Make a copy of the .docx file in the conversion folder
  var copiedFile = file.makeCopy(fileName, conversionFolder);

  // Convert the copied .docx file to Google Doc
  var conversionResult = convertDocxToGoogleDoc(copiedFile.getId());

  if (conversionResult && conversionResult.id) {
    // Move the converted Google Doc back to the original folder
    var convertedDoc = DriveApp.getFileById(conversionResult.id);
    var parents = convertedDoc.getParents();

    while (parents.hasNext()) {
      var parent = parents.next();
      parent.removeFile(convertedDoc);
    }

    originalFolder.addFile(convertedDoc);
    Logger.log('Conversion and move successful. Google Doc ID: ' + conversionResult.id);
  } else {
    Logger.log('Conversion failed for file: ' + fileName);
  }

  // Optional: Delete the copied .docx file from the conversion folder
  // copiedFile.setTrashed(true);
} else {
  Logger.log('No .docx files found in the original folder.');
}
}

// Helper function to log execution time in minutes, seconds, and milliseconds
function logExecutionTime(func, functionName) {
var startTime = new Date();

func();

var endTime = new Date();
var executionTime = endTime - startTime;

var minutes = Math.floor(executionTime / 60000);
var seconds = ((executionTime % 60000) / 1000).toFixed(3);

Logger.log(functionName + ' execution time: ' + minutes + ' minutes, ' + seconds + ' seconds');
}
