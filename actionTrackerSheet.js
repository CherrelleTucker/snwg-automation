/* Purpose: 
1. Search agendas for action items to be completed and populate in the Action Tracking Google Sheet.
2. Push status updates from from the Action Tracking Google sheet to the Action source agendas as changed.
/* Purpose: 
1. Search agendas for action items to be completed and populate in the Action Tracking Google Sheet.
2. Push status updates from from the Action Tracking Google sheet to the Action source agendas as changed.

Future development: 
Preserve links in task items from the source documents. Will first require the same development in the inDocActionItems script, as the links are lost in that action first. 
Future development: 
Preserve links in task items from the source documents. Will first require the same development in the inDocActionItems script, as the links are lost in that action first. 

To note: 
This script is developed as a Google Apps Script container script: i.e. a script that is bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This container script acts as the file's custom script, allowing users to extend the functionality of the file by adding custom functions, triggers, and menu items to enhance its behavior and automation.
To note: 
This script is developed as a Google Apps Script container script: i.e. a script that is bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This container script acts as the file's custom script, allowing users to extend the functionality of the file by adding custom functions, triggers, and menu items to enhance its behavior and automation.

<<<<<<< HEAD
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

//////////////////////////////////////////////////

// Testing function to verify Google Sheet ID, Google Sheet Name, Sheet tab names for all tabs
function getSpreadsheetInfoTest() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
=======
// Instructions for Using this Script in your container file:
// 1. Open a new or existing Google Sheets file where you want to use the script.
// 2. Click on "Extensions" in the top menu and select "Apps Script" from the dropdown menu. This will open the Google Apps Script editor in a new tab.
// 3. Copy and paste the provided script into the Apps Script editor, replacing the existing code (if any).
// 4. In the MOPopulate function, replace the placeholder values for folderId and spreadsheetId with the actual IDs of your Google Drive folder containing the agendas and the Google Sheets spreadsheet where you want to track the actions, respectively.
// 5. In the updateStatus function, replace the placeholder value for spreadsheetId with the actual ID of your Google Sheets spreadsheet.
// 6. Save the script by clicking on the floppy disk icon or by pressing "Ctrl + S" (Windows) or "Cmd + S" (Mac).
// 7. Go back to your Google Sheets file, refresh the page, and you'll see a new custom menu labeled "Action Items" in the top menu.
// 8. Click on "Action Items" in the top menu to access the custom menu. You'll find two options:
//    a. "Get Actions from Agendas": This option will pull actions from the specified agendas and populate them in the "MO" sheet in your Google Sheets file.
//    b. "Update Status in Source Document": This option will push status updates from the "MO" sheet back to the corresponding action items in the source agendas.
// 9. Whenever you want to get actions from the agendas or update status in the source documents, simply click on the corresponding option from the "Action Items" menu.
// Note: Make sure to properly set up the correct folder structure in Google Drive and name your agendas and sheets according to the script's logic for pulling and updating actions. This documentation assumes you have some basic familiarity with Google Apps Script and how to run container-bound scripts within a Google Sheets file.

//////////////////////////////////////////////////

// Global Variables: Replace 'folderId' and 'spreadsheetId' with your actual Google Drive folder ID and Google Sheets spreadsheet ID, respectively.
// var folderId = '1WKYw4jnP6ejRkOLAIPoPvbEYClaLE4eR'; // SNWG MO Weekly Internal Planning > FY 23 Google Drive folder
var folderId = '1SRIUs7CUEdGUw0r1PI52e0OJpfXYN0z8'; // SNWG MO Weekly Internal Planning > FY 24 Google Drive folder
var spreadsheetId = '1uYgX660tpizNbIy44ddQogrRphfwZqn1D0Oa2RlSYKg'; // SNWG MO Action Tracking Spreadsheet
>>>>>>> b746d3712728e7b81086aac526d2a6c48e28b83d

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

///////////////////Custom Menu//////////////////////////////////////////////

// Add a custom menu to the spreadsheet
function onOpen() {
<<<<<<< HEAD
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Action Items')
      .addItem('Pull actions for ALL folders', 'pullActionsForAllFolders')
      .addItem('Pull actions for MO', 'pullActionsForFolderMO')
      .addItem('Pull actions for SEP', 'pullActionsForFolderSEP')
      .addItem('Pull actions for DevSeed', 'pullActionsForFolderDevSeed')
      .addSeparator()
      .addItem('Push actions from ALL tabs', 'pushActionsFromAllTabs') 
      .addItem('Push actions from MO', 'pushActionsFromTabMO') 
      .addItem('Push actions from SEP', 'pushActionsFromTabSEP') 
      .addItem('Push actions from DevSeed', 'pushActionsFromTabDevSeed') 
=======
  SpreadsheetApp.getUi()
      .createMenu('Action Items')
      .addItem('Get Actions from Agendas','MOPopulate')
      .addItem('Push Status Updates to Source Document','updateStatus')
>>>>>>> b746d3712728e7b81086aac526d2a6c48e28b83d
      .addToUi();
}

// Global variables
var spreadsheetId = '13xgmbfP8X8lu9tlD_cHVCF9RmQeKvToqtbiDSfp_Nfg'; // Testing ALL Action Tracking Sheet
var folderIds = {
  MO: '1SRIUs7CUEdGUw0r1PI52e0OJpfXYN0z8',
  SEP: '1Cw_sdH_IleGbtW1mVoWnJ0yqoyzr4Oe0',
  DevSeed: '1Bvj1b1u2LGwjW5fStQaLRpZX5pmtjHSj'
};

///////////////////Pull///////////////////////////////////////////////

// Function to pull actions for all folders
function pullActionsForAllFolders() {
  logExecutionTime(function() {
    pullActionsForFolder('MO');
    pullActionsForFolder('SEP');
    pullActionsForFolder('DevSeed');
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

// Helper function to pull actions from a specific folder
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

// Global variables
var spreadsheetId = '13xgmbfP8X8lu9tlD_cHVCF9RmQeKvToqtbiDSfp_Nfg'; // Testing ALL Action Tracking Sheet
var folderIds = {
  MO: '1SRIUs7CUEdGUw0r1PI52e0OJpfXYN0z8',
  SEP: '1Cw_sdH_IleGbtW1mVoWnJ0yqoyzr4Oe0',
  DevSeed: '1Bvj1b1u2LGwjW5fStQaLRpZX5pmtjHSj'
};

///////////////////Pull///////////////////////////////////////////////

// Function to pull actions for all folders
function pullActionsForAllFolders() {
  logExecutionTime(function() {
    pullActionsForFolder('MO');
    pullActionsForFolder('SEP');
    pullActionsForFolder('DevSeed');
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

// Helper function to pull actions from a specific folder
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

// Pull helper function: Pull actions from documents
function pullActionsFromDocuments(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_DOCS);

  var allActions = [];

  while (files.hasNext()) {
    var file = files.next();
    var docId = file.getId();
    var doc = DocumentApp.openById(docId);
    var tables = doc.getBody().getTables();

    for (var i = 0; i < tables.length; i++) {
      var table = tables[i];
      var numCols = table.getRow(0).getNumCells();

      if (numCols !== 3) { // Skip tables that do not have 3 columns
        continue;
      }

      var documentName = file.getName();
      var documentLink = '=HYPERLINK("' + file.getUrl() + '", "' + documentName + '")';

      if (documentName.toLowerCase().indexOf('template') !== -1) { // Skip files with "Template" in the title
        continue;
      }

      var tableData = tableTo2DArray(table);
      var modifiedTableData = tableData.map(function (row) {
        return [documentLink].concat(row);
      });

      for (var j = 0; j < modifiedTableData.length; j++) {
        var row = modifiedTableData[j];

        if (row.some(function (cell) { return cell === ''; })) {  // Skip rows with empty cells
          continue;
        }

        allActions.push(row);
      }
    }
  }

  return allActions;
}

// Pull helper function: Find the second table with qualifying headers
// Pull helper function: Find the second table with qualifying headers
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

// Pull helper function: Convert table to 2D array for convenience and flexibility in data processing
// Pull helper function: Convert table to 2D array for convenience and flexibility in data processing
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

// Pull helper function: Populate the Sheet with actions
// Pull helper function: Populate the Sheet with actions
function populateSheetWithActions(spreadsheetId, sheetName, actions) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    // Clear only the contents of columns A to E
    sheet.getRange(1, 1, sheet.getMaxRows(), 5).clearContent();
    // Clear only the contents of columns A to E
    sheet.getRange(1, 1, sheet.getMaxRows(), 5).clearContent();
  }

  var headerRow = ['Action Source', 'Status', 'Assigned to', 'Task'];
  actions.unshift(headerRow);

  var numRows = actions.length;
  var numCols = headerRow.length;

  if (numRows > 0 && numCols > 0) {
    // Set values starting from cell A1
    // Set values starting from cell A1
    var range = sheet.getRange(1, 1, numRows, numCols);
    range.setValues(actions);
  }
}

<<<<<<< HEAD
// Primary Pull function: Pull action items from meeting notes to populate action tracking workbook
=======
// Primary function: Pull action items from meeting notes to populate action tracking workbook
>>>>>>> b746d3712728e7b81086aac526d2a6c48e28b83d
function MOPopulate() {
  var tablePullSheetName = 'MO';

  Logger.log('Step 1: Pulling actions from documents...');
  var actions = pullActionsFromDocuments(folderIds.MO);
  var actions = pullActionsFromDocuments(folderIds.MO);
  Logger.log('Step 1: Actions retrieved:', actions);

  Logger.log('Step 2: Populating sheet with actions...');
  populateSheetWithActions(spreadsheetId, tablePullSheetName, actions);
  Logger.log('Step 2: Sheet populated with actions.');
}

////////////////// Pull Testing Log //////////////////////////////////////////

/*
pullActionsForFolderMO: success 2024-01-11 10:43:24 AM	pullActionsForFolderMO execution time: 0 minutes, 12.737 seconds
pullActionsForFolderSEP: success 2024-01-10 10:46:01 AM	pullActionsForFolderSEP execution time: 0 minutes, 10.042 seconds
pullActionsForFolderDevSeed: success 10:59:36 AM	pullActionsForFolderDevSeed execution time: 0 minutes, 9.670 seconds
pullActionsForAllFolders: success 11:02:56 AM	pullActionsForAllFolders execution time: 0 minutes, 37.867 seconds
*/

////////////////// Push ////////////////////////////////////////////////////////
////////////////// Pull Testing Log //////////////////////////////////////////

/*
pullActionsForFolderMO: success 2024-01-11 10:43:24 AM	pullActionsForFolderMO execution time: 0 minutes, 12.737 seconds
pullActionsForFolderSEP: success 2024-01-10 10:46:01 AM	pullActionsForFolderSEP execution time: 0 minutes, 10.042 seconds
pullActionsForFolderDevSeed: success 10:59:36 AM	pullActionsForFolderDevSeed execution time: 0 minutes, 9.670 seconds
pullActionsForAllFolders: success 11:02:56 AM	pullActionsForAllFolders execution time: 0 minutes, 37.867 seconds
*/

////////////////// Push ////////////////////////////////////////////////////////

// Function to call the primary function on all named tabs
function pushActionsFromAllTabs() {
  logExecutionTime(function() {
    updateStatusOnTab("MO");
    updateStatusOnTab("SEP");
    updateStatusOnTab("DevSeed");
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

function updateStatusOnTab(tabName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var range = sheet.getRange("A:G"); // Assuming the URL is in Column G
  var values = range.getValues();

  // Loop through each row of the data
  for (var i = 1; i < values.length; i++) {
    var actionSourceUrl = values[i][6]; // Column G (assuming the URL is in Column G)
    var status = values[i][1]; // Column B
    var assignedTo = values[i][2]; // Column C
    var task = values[i][3]; // Column D

    // Check if the row is empty or if it's the header row
    if (actionSourceUrl || status || assignedTo || task) {
      // Log actionSourceUrl for investigation
      console.log("Row:", i + 1, "actionSourceUrl:", actionSourceUrl);

      // Call the function to update status in the source document
      updateStatusInSourceDoc(actionSourceUrl, status, task);
    }
  }
}

function updateStatusInSourceDoc(actionSource, status, task) {
  // Column G now contains the extracted URL
  var actionSourceUrl = actionSource; // No need for extractUrlFromHyperlink

  if (!actionSourceUrl) {
    console.error("Invalid URL in Column A:", actionSource);
    return;
  }

    try {
      console.log("Attempting to open document with URL:", actionSourceUrl);
      var sourceDoc = DocumentApp.openByUrl(actionSourceUrl);
      } catch (error) {
        console.error("Error opening document by URL:", error);
        return;
      }
  // Get the body of the document
  var body = sourceDoc.getBody();

  // Find the table with the specified headings
  var table = findTableByHeadings(body, ["Status", "Action"]);

  if (table) {
    // Find the row with the matching task in the "Action" column
    var rowIndex = findRowIndexByColumnValue(table, "Action", task);

    if (rowIndex !== -1) {
      // Update the "Status" column with the status from the sheet
      table.getCell(rowIndex, getColumnIndex(table, "Status")).setText(status);
    }
  }
}

function findTableByHeadings(body, headings) {
  // Function to find a table in the document with specified headings
  var tables = body.getTables();
  
  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    var headerRow = table.getRow(0);
    
    // Check if the table has the specified headings
    if (headings.every(function (heading) {
      return headerRow.getText().indexOf(heading) !== -1;
    })) {
      return table;
// Function to call the primary function on all named tabs
function pushActionsFromAllTabs() {
  logExecutionTime(function() {
    updateStatusOnTab("MO");
    updateStatusOnTab("SEP");
    updateStatusOnTab("DevSeed");
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

function updateStatusOnTab(tabName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var range = sheet.getRange("A:G"); // Assuming the URL is in Column G
  var values = range.getValues();

  // Loop through each row of the data
  for (var i = 1; i < values.length; i++) {
    var actionSourceUrl = values[i][6]; // Column G (assuming the URL is in Column G)
    var status = values[i][1]; // Column B
    var assignedTo = values[i][2]; // Column C
    var task = values[i][3]; // Column D

    // Check if the row is empty or if it's the header row
    if (actionSourceUrl || status || assignedTo || task) {
      // Log actionSourceUrl for investigation
      console.log("Row:", i + 1, "actionSourceUrl:", actionSourceUrl);

      // Call the function to update status in the source document
      updateStatusInSourceDoc(actionSourceUrl, status, task);
    }
  }
}

function updateStatusInSourceDoc(actionSource, status, task) {
  // Column G now contains the extracted URL
  var actionSourceUrl = actionSource; // No need for extractUrlFromHyperlink

  if (!actionSourceUrl) {
    console.error("Invalid URL in Column A:", actionSource);
    return;
  }

    try {
      console.log("Attempting to open document with URL:", actionSourceUrl);
      var sourceDoc = DocumentApp.openByUrl(actionSourceUrl);
      } catch (error) {
        console.error("Error opening document by URL:", error);
        return;
      }
  // Get the body of the document
  var body = sourceDoc.getBody();

  // Find the table with the specified headings
  var table = findTableByHeadings(body, ["Status", "Action"]);

  if (table) {
    // Find the row with the matching task in the "Action" column
    var rowIndex = findRowIndexByColumnValue(table, "Action", task);

    if (rowIndex !== -1) {
      // Update the "Status" column with the status from the sheet
      table.getCell(rowIndex, getColumnIndex(table, "Status")).setText(status);
    }
  }
}

function findTableByHeadings(body, headings) {
  // Function to find a table in the document with specified headings
  var tables = body.getTables();
  
  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    var headerRow = table.getRow(0);
    
    // Check if the table has the specified headings
    if (headings.every(function (heading) {
      return headerRow.getText().indexOf(heading) !== -1;
    })) {
      return table;
    }
  }
  
  
  return null;
}

function findRowIndexByColumnValue(table, columnName, value) {
  // Function to find the row index in the table with a specific column value
  var columnIndex = getColumnIndex(table, columnName);
  
  if (columnIndex !== -1) {
    var numRows = table.getNumRows();
    
    for (var i = 1; i < numRows; i++) {
      if (table.getCell(i, columnIndex).getText() === value) {
        return i;
      }
    }
  }
  
  return -1;
}

function getColumnIndex(table, columnName) {
  // Function to get the index of a column in the table
  var headerRow = table.getRow(0);
  var numCells = headerRow.getNumCells();
  
  for (var i = 0; i < numCells; i++) {
    if (headerRow.getCell(i).getText() === columnName) {
      return i;
    }
  }
  
  return -1;
}

<<<<<<< HEAD
////////////////// Push Testing Log //////////////////////////////////////////
=======
// Primary function: push status updates back to action tracking tables in meeting notes
function updateStatus() {
  var sheetName = 'MO'; // <--Replace with the name of your sheet
>>>>>>> b746d3712728e7b81086aac526d2a6c48e28b83d

/* 
'pushActionsFromTabMO':  success 3:59:02 PM	execution time: 0 minutes, 13.421 seconds
'pushActionsFromTabSEP': success 4:02:31 PM	execution time: 0 minutes, 15.707 seconds
'pushActionsFromTabDevSeed': success 4:12:57 PM execution time: 0 minutes, 11.042 seconds
'pushActionsFromAllTabs': 4:16:43 PM execution time: 0 minutes, 31.064 seconds
*/

<<<<<<< HEAD
///////////////Combine Open Actions///////////////////////////////////////////////

=======
  Logger.log('Step 2: Syncing status to source documents...');
  syncStatusToSource(actions);
  Logger.log('Step 2: Status synced to source documents.');
}

//////////////////////////////////////////////////////////////
>>>>>>> b746d3712728e7b81086aac526d2a6c48e28b83d
// function to copy all rows of each sheet that do not contain "Done" in the "Status Column". Maintain column D formatting from source sheet to track length of time the action has been open.
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
<<<<<<< HEAD
}
=======
}


>>>>>>> b746d3712728e7b81086aac526d2a6c48e28b83d
