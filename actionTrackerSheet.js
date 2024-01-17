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

//////////////////////////////////////////////////

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
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Action Items')
      .addItem('Pull actions for ALL folders', 'pullActionsForAllFolders')
      .addItem('Pull actions for MO', 'pullActionsForFolderMO')
      .addItem('Pull actions for SEP', 'pullActionsForFolderSEP')
      .addItem('Pull actions for DevSeed', 'pullActionsForFolderDevSeed')
      .addItem('Pull actions for Assessment HQ', 'pullActionsForFolderAssessmentHQ')
      .addSeparator()
      .addItem('Push actions from ALL tabs', 'pushActionsFromAllTabs') 
      .addItem('Push actions from MO', 'pushActionsFromTabMO') 
      .addItem('Push actions from SEP', 'pushActionsFromTabSEP') 
      .addItem('Push actions from DevSeed', 'pushActionsFromTabDevSeed') 
      .addItem('Push actions from AssessmentHQ', 'pushActionsFromTabAssessmentHQ')
      .addToUi();
}

// Global variables
var spreadsheetId = '13xgmbfP8X8lu9tlD_cHVCF9RmQeKvToqtbiDSfp_Nfg'; // Testing ALL Action Tracking Sheet
var folderIds = {
  MO: '1SRIUs7CUEdGUw0r1PI52e0OJpfXYN0z8',
  SEP: '1Cw_sdH_IleGbtW1mVoWnJ0yqoyzr4Oe0',
  DevSeed: '1Bvj1b1u2LGwjW5fStQaLRpZX5pmtjHSj',
  AssessmentHQ: '1V40h1Df4TMuuGzTMiLHxyBRPC-XJhQ10'// Testing Folder
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

// Function to pull actions for AssessmentHQ folder
function pullActionsForFolderAssessmentHQ() {
  logExecutionTime(function() {
    pullActionsForFolder('AssessmentHQ');
  }, 'pullActionsForFolderAssessmentHQ');
}

// Helper function to convert .docx file to Google Doc
function convertDocxToGoogleDoc(fileId) {
  var file = DriveApp.getFileById(fileId);
  var docxBlob = file.getBlob();
  
  // Convert .docx to Google Doc
  var options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify({
      title: file.getName(),
      mimeType: MimeType.GOOGLE_DOCS
    }),
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch("https://www.googleapis.com/drive/v2/files/" + fileId + "/copy", options);
  var googleDocId = JSON.parse(response.getContentText()).id;
  
  // Delete the original .docx file
  DriveApp.getFileById(fileId).setTrashed(true);

  return googleDocId;
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

// Modified pullActionsFromDocuments function
function pullActionsFromDocuments(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var allActions = [];

  while (files.hasNext()) {
    var file = files.next();
    var docId;

    if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
      docId = file.getId();
    } else if (file.getMimeType() === MimeType.MICROSOFT_WORD || file.getName().toLowerCase().endsWith(".docx")) {
      // Convert .docx file to Google Doc
      docId = convertDocxToGoogleDoc(file.getId());
    } else {
      // Skip unsupported file types
      continue;
    }

    var doc = DocumentApp.openById(docId);
    var tables = doc.getBody().getTables();

    for (var i = 0; i < tables.length; i++) {
      var table = tables[i];
      var headers = getTableHeaders(table);

      // Check if the table has the required headers
      if (hasRequiredHeaders(headers)) {
        var documentName = file.getName();
        var documentLink = '=HYPERLINK("' + file.getUrl() + '", "' + documentName + '")';

        var tableData = tableTo2DArray(table);
        var filteredTableData = filterTableData(tableData, headers);

        for (var j = 0; j < filteredTableData.length; j++) {
          var row = [documentLink].concat(filteredTableData[j]);
          allActions.push(row);
        }
      }
    }
  }

  return allActions;
}

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

// Primary Pull function: Pull action items from meeting notes to populate action tracking workbook
function MOPopulate() {
  var tablePullSheetName = 'MO';

  Logger.log('Step 1: Pulling actions from documents...');
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

// Push function to call the primary function on all named tabs
function pushActionsFromAllTabs() {
  logExecutionTime(function() {
    updateStatusOnTab("MO");
    updateStatusOnTab("SEP");
    updateStatusOnTab("DevSeed");
    updateStatusOnTab("AssessmentHQ");
  }, 'pushActionsFromAllTabs');
}

// Push function to call the primary function on the "MO" tab
function pushActionsFromTabMO() {
  logExecutionTime(function() {
    updateStatusOnTab("MO");
  }, 'pushActionsFromTabMO');
}

// Push function to call the primary function on the "SEP" tab
function pushActionsFromTabSEP() {
  logExecutionTime(function() {
    updateStatusOnTab("SEP");
  }, 'pushActionsFromTabSEP');
}

// Push function to call the primary function on the "DevSeed" tab
function pushActionsFromTabDevSeed() {
  logExecutionTime(function() {
    updateStatusOnTab("DevSeed");
  }, 'pushActionsFromTabDevSeed');
}

// Push function to call the primary function on the "AssessmentHQ" tab
function pushActionsFromTabAssessmentHQ() {
  logExecutionTime(function() {
    updateStatusOnTab("AssessmentHQ");
  }, 'pushActionsFromTabAssessmentHQ');
}

// push function:
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

//push function:
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
  var table = findTableByHeadings(body);

  if (table) {
    // Log the table content for investigation
    console.log("Table Content:", table.getText());

    // Check the headers dynamically and find the row with the matching task
    var columnIndex;
    var expectedHeaders;
    var lowerCaseHeaders = table.getRow(0).getText().toLowerCase().trim(); // Trim leading/trailing spaces
    if (lowerCaseHeaders.includes("who")) {
      // Headers are "Who What Status"
      columnIndex = getColumnIndex(table, "What");
      expectedHeaders = "Who What Status";
    } else if (lowerCaseHeaders.includes("status") && 
               lowerCaseHeaders.includes("owner") &&
               lowerCaseHeaders.includes("action")) {
      // Headers are "Status Owner Action"
      columnIndex = getColumnIndex(table, "Action");
      expectedHeaders = "Status Owner Action";
    } else {
      // Headers not recognized
      console.error("Headers not recognized. Detected headers:", lowerCaseHeaders);
      expectedHeaders = "Unknown";
    }

    // Find the row with the matching task in the specified column
    var rowIndex = findRowIndexByColumnValue(table, table.getRow(0).getCell(columnIndex).getText(), task);

    if (rowIndex !== -1) {
      // Update the "Status" column with the status from the sheet
      table.getCell(rowIndex, getColumnIndex(table, "Status")).setText(status);
      console.log("Status Updated Successfully!");
    } else {
      console.error("Row with task not found in the table. Expected headers:", expectedHeaders);
    }
  } else {
    console.error("Table not found with specified headings. Expected headers:", expectedHeaders);
  }
}

// Push function: Get table headers
function getTableHeaders(table) {
  var headers = [];
  var row1 = table.getRow(0);

  for (var col = 0; col < row1.getNumCells(); col++) {
    headers.push(row1.getCell(col).getText().trim());
  }

  return headers;
}

// Helper function: Check if the table has the required headers
function hasRequiredHeaders(headers) {
  // Check if the table has "Status" "Owner" "Action" or "Who" "What" "Status"
  var hasStatus = headers.includes('Status') || headers.includes('Who'); // Column one heading
  var hasOwner = headers.includes('Owner') || headers.includes('What'); // Column two heading
  var hasAction = headers.includes('Action') || headers.includes('Status'); // Column three heading

  return hasStatus && hasOwner && hasAction;
}

// Helper function: Filter table data based on headers
function filterTableData(tableData, headers) {
  var filteredData = [];

  for (var i = 0; i < tableData.length; i++) {
    var rowData = tableData[i];
    var filteredRow = ['', '', '']; // Initialize with empty values

    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var columnIndex = headers.indexOf(header);

      if (columnIndex !== -1) {
        // Set values in corresponding columns based on header variations
        switch (header.toLowerCase()) {
          case 'status owner action':
            filteredRow[0] = rowData[columnIndex];
            filteredRow[1] = rowData[columnIndex + 1]; // Assuming 'Owner' is next
            filteredRow[2] = rowData[columnIndex + 2]; // Assuming 'Action' is next
            break;
          case 'who what status':
            filteredRow[0] = rowData[columnIndex + 2]; // Assuming 'Status' is last
            filteredRow[1] = rowData[columnIndex]; // Assuming 'Who' is first
            filteredRow[2] = rowData[columnIndex + 1]; // Assuming 'What' is in the middle
            break;
          // Add more cases for other header variations if needed
        }
      }
    }

    filteredData.push(filteredRow);
  }
  return filteredData;
}

function findTableByHeadings(body) {
  // Push function to find a table in the document with specified headings
  var tables = body.getTables();

  // Define possible header patterns
  var headerPatterns = [
    ["who", "what", "status"],
    ["status", "owner", "action"]
  ];

  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    var headerRow = table.getRow(0);
    var headerText = headerRow.getText().toLowerCase().replace(/\n/g, '');

    // Check if the table headers match any of the predefined patterns
    for (var j = 0; j < headerPatterns.length; j++) {
      var pattern = headerPatterns[j];
      var patternText = pattern.join("");
      if (headerText.includes(patternText)) {
        console.log("Table found with specified headings:", patternText);
        return table;
      }
    }
  }

  console.error("Table not found with specified headings.");
  console.log("Detected headings:", tables.map(table => table.getRow(0).getText().toLowerCase()));
  return null;
}

function findRowIndexByColumnValue(table, columnName, value) {
  // Push function to find the row index in the table with a specific column value
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
  // Push function to get the index of a column in the table
  var headerRow = table.getRow(0);
  var numCells = headerRow.getNumCells();
  
  for (var i = 0; i < numCells; i++) {
    if (headerRow.getCell(i).getText() === columnName) {
      return i;
    }
  }
  
  return -1;
}

////////////////// Push Testing Log //////////////////////////////////////////

/* 
'pushActionsFromTabMO':  success 3:59:02 PM	execution time: 0 minutes, 13.421 seconds
'pushActionsFromTabSEP': success 4:02:31 PM	execution time: 0 minutes, 15.707 seconds
'pushActionsFromTabDevSeed': success 4:12:57 PM execution time: 0 minutes, 11.042 seconds
'pushActionsFromAssessmentHQ': success 11:55:32 PM execution time: 0 minutes, 4.876 seconds
'pushActionsFromAllTabs': success 11:57:00 PM execution time: 0 minutes, 35.094 seconds
*/

///////////////Combine Open Actions///////////////////////////////////////////////

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
}