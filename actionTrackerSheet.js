// Purpose: 
// 1. Search agendas for action items to be completed and populate in the Action Tracking Google Sheet.
// 2. Push status updates from from the Action Tracking Google sheet to the Action source agendas as changed.

// Future development: preserve links in task items from the source documents. Will first require the same development in the inDocActionItems script, as the links are lost in that action first. 

// This script is developed as a Google Apps Script container script: a script that is bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This container script acts as the file's custom script, allowing users to extend the functionality of the file by adding custom functions, triggers, and menu items to enhance its behavior and automation.

// Instructions for Using the Script:
// 1. Open a new or existing Google Sheets file where you want to use the script.
// 2. Click on "Extensions" in the top menu and select "Apps Script" from the dropdown menu. This will open the Google Apps Script editor in a new tab.
// 3. Copy and paste the provided script into the Apps Script editor, replacing the existing code (if any).
// 4. In the TablePullPopulate function, replace the placeholder values for folderId and spreadsheetId with the actual IDs of your Google Drive folder containing the agendas and the Google Sheets spreadsheet where you want to track the actions, respectively.
// 5. In the updateStatus function, replace the placeholder value for spreadsheetId with the actual ID of your Google Sheets spreadsheet.
// 6. Save the script by clicking on the floppy disk icon or by pressing "Ctrl + S" (Windows) or "Cmd + S" (Mac).
// 7. Go back to your Google Sheets file, refresh the page, and you'll see a new custom menu labeled "Action Items" in the top menu.
// 8. Click on "Action Items" in the top menu to access the custom menu. You'll find two options:
//    a. "Get Actions from Agendas": This option will pull actions from the specified agendas and populate them in the "Table Pull" sheet in your Google Sheets file.
//    b. "Update Status in Source Document": This option will push status updates from the "Table Pull" sheet back to the corresponding action items in the source agendas.
// 9. Whenever you want to get actions from the agendas or update status in the source documents, simply click on the corresponding option from the "Action Items" menu.
// Note: Make sure to properly set up the correct folder structure in Google Drive and name your agendas and sheets according to the script's logic for pulling and updating actions. This documentation assumes you have some basic familiarity with Google Apps Script and how to run container-bound scripts within a Google Sheets file.

//////////////////////////////////////////////////

// Global Variables - Replace with your actual IDs
var folderId = '1WKYw4jnP6ejRkOLAIPoPvbEYClaLE4eR';
var spreadsheetId = '1uYgX660tpizNbIy44ddQogrRphfwZqn1D0Oa2RlSYKg';

// Secondary function to create custom menu
function customMenu() {
  SpreadsheetApp.getUi()
      .createMenu('Action Items')
      .addItem('Get Actions from Agendas','TablePullPopulate')
      .addItem('Update Status in Source Document','updateStatus')
      .addToUi();
}

// Primary function: TablePullPopulate
  // INSTRUCTIONS: Replace 'folderId' and 'spreadsheetId' with your actual Google Drive folder ID and Google Sheets spreadsheet ID, respectively.
function TablePullPopulate() {
  var tablePullSheetName = 'Table Pull';

  var actions = pullActionsFromDocuments(folderId);
  Logger.log('Actions:', actions);
  populateSheetWithActions(spreadsheetId, tablePullSheetName, actions);
}

// Helper function: Pull actions from documents
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

// Helper function: Find the second table with qualifying headers
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

// Helper function: Convert table to 2D array
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

// Helper function: Populate the sheet with actions
function populateSheetWithActions(spreadsheetId, sheetName, actions) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  var headerRow = ['Action Source', 'Status', 'Assigned to', 'Task'];
  actions.unshift(headerRow);

  var numRows = actions.length;
  var numCols = headerRow.length;

  if (numRows > 0 && numCols > 0) {
    var range = sheet.getRange(1, 1, numRows, numCols);
    range.setValues(actions);
  }
}
// ////////////////////////////////////////////////////////////////////////////////////////////////////////

// Primary function: updateStatus
function updateStatus() {
  var sheetName = 'Table Pull'; // Replace with the name of your sheet

  var actions = getActionsFromSheet(spreadsheetId, sheetName);
  syncStatusToSource(actions);
}

// Helper function: Get actions from the sheet
function getActionsFromSheet(spreadsheetId, sheetName) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  var data = sheet.getDataRange().getValues();
  var actions = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = row[1];
    var task = row[3];

    if (status.trim().toLowerCase() !== 'not started') {
      actions.push({
        status: status,
        task: task
      });
    }
  }

  return actions;
}

// Helper function: Sync status to source
function syncStatusToSource(actions) {
  for (var i = 0; i < actions.length; i++) {
    var action = actions[i];
    var status = action.status;
    var task = action.task;

    var document = findDocumentWithTask(task);

    if (document) {
      updateStatusInDocument(document, status, task); // Pass task as a parameter
    }
  }
}

// Helper function: Find document with the matching task
function findDocumentWithTask(task) {
  var folderId = '1WKYw4jnP6ejRkOLAIPoPvbEYClaLE4eR'; // Replace with the ID of the folder containing your documents
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_DOCS);

  while (files.hasNext()) {
    var file = files.next();
    var docId = file.getId();
    var doc = DocumentApp.openById(docId);
    var tables = doc.getBody().getTables();

    if (tables.length < 2) {
      // Skip documents with less than 2 tables
      continue;
    }

    for (var i = 0; i < tables.length; i++) {
      var table = tables[i];
      var numCols = table.getRow(0).getNumCells();

      if (numCols !== 3) {
        // Skip tables with less than 3 columns
        continue;
      }

      var numRows = table.getNumRows();

      for (var j = 1; j < numRows; j++) {
        var row = table.getRow(j);
        var action = row.getCell(2).getText();

        if (action === task) {
          return doc;
        }
      }
    }
  }

  return null;
}

// Helper function: Update status in the document
function updateStatusInDocument(document, status, task) { // Add task parameter
  var tables = document.getBody().getTables();

  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    var row1 = table.getRow(0);
    var statusIndex = -1;
    var ownerIndex = -1;
    var actionIndex = -1;

    for (var j = 0; j < row1.getNumCells(); j++) {
      var cellText = row1.getCell(j).getText().toLowerCase();

      if (cellText === 'status') {
        statusIndex = j;
      } else if (cellText === 'owner') {
        ownerIndex = j;
      } else if (cellText === 'action') {
        actionIndex = j;
      }
    }

    if (statusIndex !== -1 && ownerIndex !== -1 && actionIndex !== -1) {
      var numRows = table.getNumRows();

      for (var k = 1; k < numRows; k++) {
        var row = table.getRow(k);
        var action = row.getCell(actionIndex).getText();

        if (action === task) {
          var existingStatus = row.getCell(statusIndex).getText();

          if (existingStatus.trim().toLowerCase() !== status.trim().toLowerCase()) {
            row.getCell(statusIndex).editAsText().setText(status);
          }
        }
      }
    }
  }
}