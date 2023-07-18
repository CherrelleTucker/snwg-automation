// Purpose: 
// 1. Search agendas in SNWG Meeting Notes folder for meeting agenda action items to be completed and populate Source Document, Status, Owner, and Task in SNWG MO Action Tracker Sheet. 
// 2. Push updated statuses to the action tracking table in their agenda.

// Issue: preserve hyperlinks in task items from the source documents (also issue within in-document action tracking script)

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Action Items')
      .addItem('Get Actions','importActionsFromFolder')
      .addItem('Update Document Status','updateStatus')
      .addToUi();
}
// Primary function: TablePullPopulate
function TablePullPopulate() {
  var folderId = '1WKYw4jnP6ejRkOLAIPoPvbEYClaLE4eR';
  var spreadsheetId = '1uYgX660tpizNbIy44ddQogrRphfwZqn1D0Oa2RlSYKg'; // Replace with your actual spreadsheet ID
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

      if (numCols !== 3) {
        // Skip tables that do not have 3 columns
        continue;
      }

      var documentName = file.getName();
      var documentLink = '=HYPERLINK("' + file.getUrl() + '", "' + documentName + '")';

      if (documentName.toLowerCase().indexOf('template') !== -1) {
        // Skip files with "Template" in the title
        continue;
      }

      var tableData = tableTo2DArray(table);
      var modifiedTableData = tableData.map(function (row) {
        return [documentLink].concat(row);
      });

      for (var j = 0; j < modifiedTableData.length; j++) {
        var row = modifiedTableData[j];

        if (row.some(function (cell) { return cell === ''; })) {
          // Skip rows with empty cells
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

// Primary function: push statuses updated in tracking sheet to their source documents
function updateStatus() {
  var spreadsheetId = '1uYgX660tpizNbIy44ddQogrRphfwZqn1D0Oa2RlSYKg'; // Replace with your actual spreadsheet ID
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
