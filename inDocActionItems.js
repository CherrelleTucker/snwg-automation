// Purpose: 
// 1. Search agendas for action items to be completed and populate in the Action Tracking Google Sheet.
// 2. Push status updates from from the Action Tracking Google sheet to the Action source agendas as changed.

// Future development: 
// Preserve links in task items when copying to action table

// To Note:
// This script is developed for use as either as a Google Apps Script container script or as a Google Apps Script library script. 
// Google Apps Script container script: a script that is bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This container script acts as the file's custom script, allowing users to extend the functionality of the file by adding custom functions, triggers, and menu items to enhance its behavior and automation.

// Google Apps Script library script: a self-contained script that contains reusable functions and can be attached to multiple projects or files. By attaching the library script to different projects, developers can access and use its functions across various files, enabling code sharing and improving code maintenance and version control.

// To Use As container script:
// Open your Google Document: Open the Google Document where you want to extract action items and populate them in an action tracking table.
// Open the Script Editor: Click on "Extensions" in the top menu and then select "Apps Script." This will open the Google Apps Script editor in a new tab.
// Copy the Script: Copy the entire script provided above and paste it into the Google Apps Script editor.
// Save the Script: Click on the floppy disk icon or press "Ctrl + S" (Windows) or "Cmd + S" (Mac) to save the script.
// Set Up Trigger (Optional): If you want to automate the process, you can set up a trigger to run the script automatically. To do this, click on the clock icon in the Google Apps Script editor, then add a new trigger, and set it to run the "runBothActionItems" function on the desired schedule (e.g., hourly, daily, etc.).
// Enable Permissions:The script will request permission to access your Google Document. Click "Continue" and grant the necessary permissions.
// Use the "Populate Actions" Button: After saving and granting permissions, you will see a new "Populate Actions" button in the Google Document's toolbar. Whenever you want to extract action items and populate them in the action tracking table, click this button.
// Action Items Population:The script will search for action items in the document's paragraphs and existing action item tables. It will then populate them in the action tracking table under the "Status," "Owner," and "Action" columns. Any new action items found in the document will be added to the table.
// Note that the script mentions "Future development: Preserve links in task items when copying to action table." If you want to implement this feature, you may need to modify the script accordingly.
// Please ensure that you are familiar with the Google Apps Script environment and have appropriate access to edit the Google Document before running the script. Also, review and customize the script as per your specific requirements before using it.

//////////////////////////////////////////////////

// To use this script as a library script:
// 1. Obtain the script ID of the inDocActionItems library script.
//  script ID: 1kBbrOJCXewvSixfq1yR8d-lEtgDG5yzD9-pqPeuC9ugLka7gQULwkBH_ <-- verify current library script id by checking in Project settings (gear icon).
// 2. Open the container document where you want to use this script.
// 3. Click on the "Project settings" gear icon in the script editor.
// 4. In the "Script ID" field, replace the existing script ID with the script ID of the inDocActionItems library script.
// 5. Click "Save" to update the script ID.

///////////////////////////////////////////////////

// Helper function that extracts attendees' names from the document.
function extractAttendees(body) {
  const attendees = [];
  const paragraphs = body.getParagraphs();

  for (let i = 0; i < paragraphs.length; i++) {
    const paragraph = paragraphs[i];
    const text = paragraph.getText();

    if (text.startsWith('Attendees:')) {
      const attendeeText = text.replace('Attendees:', '').trim();
      const attendeeNames = attendeeText.split(',');

      attendeeNames.forEach(name => attendees.push(name.trim().split(' ')[0]));
      break;
    }
  }

  return attendees;
}

// Helper function that creates a new action item table in the document.
function createSecondTable(body) {
  const table = body.appendTable();
  const header = table.appendTableRow();
  header.appendTableCell("Status");
  header.appendTableCell("Owner");
  header.appendTableCell("Action");
  return table;
}

// Helper function that retrieves existing action items from the table in the document.
function getExistingActions(table) {
  const existingActions = [];
  const numRows = table.getNumRows();

  for (let i = 1; i < numRows; i++) { // Start from 1 to skip the header row
    const row = table.getRow(i);
    const ownerName = row.getCell(1).getText();
    const actionContent = row.getCell(2).getText();
    existingActions.push({ name: ownerName, action: actionContent });
  }

  return existingActions;
}

// Helper function that checks if an action item already exists in the table.
function checkIfActionExistsInTable(existingActions, actionItem) {
  for (const existingAction of existingActions) {
    if (existingAction.name === actionItem.name && existingAction.action === actionItem.action) {
      return true;
    }
  }
  return false;
}

// Helper function that populates an action item into the table.
function populateActionInTable(table, actionItem) {
  const numRows = table.getNumRows();

  for (let i = 1; i < numRows; i++) { // Start from 1 to skip the header row
    const row = table.getRow(i);
    const cell1 = row.getCell(0);
    const cell2 = row.getCell(1);
    const cell3 = row.getCell(2);

    if (cell2.getText() === '' && cell3.getText() === '') {
      cell1.setText('Not Started');
      cell2.setText(actionItem.name);
      cell3.setText(actionItem.action);
      return;
    }
  }

  const newRow = table.insertTableRow(1);
  newRow.appendTableCell().setText('Not Started'); // Status cell
  newRow.appendTableCell().setText(actionItem.name); // Owner cell
  newRow.appendTableCell().setText(actionItem.action); // Action cell
}

// Helper function: extractActionsFromTable
function extractActionsFromTable(table) {
  const actionList = [];
  
  if (table) {
    const numRows = table.getNumRows();

    for (let i = 1; i < numRows; i++) { // Start from 1 to skip the header row
      const row = table.getRow(i);
      const numCells = row.getNumCells();

      if (numCells >= 2) {
        const ownerName = row.getCell(1).getText();
        const actionContent = numCells >= 3 ? row.getCell(2).getText() : '';
        actionList.push({ name: ownerName, action: actionContent });
      }
    }
  }

  return actionList;
}

// Helper function that searches for an existing action item table in the document.
function findExistingActionTable(body) {
  const tables = body.getTables();
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    const headerRow = table.getRow(0);

    // Check if the first table has at least 2 columns
    if (headerRow.getNumCells() >= 2) {
      const cell1Text = headerRow.getCell(0).getText();
      const cell2Text = headerRow.getCell(1).getText();

      if (cell1Text === 'Status' && cell2Text === 'Owner') {
        return table;
      }
    }
  }

  return null;
}

// Primary function that extracts action items from an existing table in the document and populates them into the Action Items table.
function actionsFromTable() {
  const document = DocumentApp.getActiveDocument();
  const body = document.getBody();
  const existingTable = findExistingActionTable(body);

  if (existingTable) {
    const actionList = extractActionsFromTable(existingTable);

    if (actionList.length > 0) {
      const existingActions = getExistingActions(existingTable);

      for (const actionItem of actionList) {
        const isDuplicate = checkIfActionExistsInTable(existingActions, actionItem);

        if (!isDuplicate) {
          populateActionInTable(existingTable, actionItem);
        }
      }

      Logger.log('Actions populated from table.');
    } else {
      Logger.log('No action items found in the existing table.');
    }
  } else {
    Logger.log('No existing action item table found.');
  }
}

// Primary function that extracts action items from paragraphs in the document and populates them into an action item table.
function actionsFromParagraphs() {
  const document = DocumentApp.getActiveDocument();
  const body = document.getBody();
  const actionsPhrase = 'Action:';
  const actionList = [];
  const names = extractAttendees(body);

  const paragraphs = body.getParagraphs();
  let isActionParagraph = false;
  paragraphs.forEach(paragraph => {
    const text = paragraph.getText();
    const startIndex = text.indexOf(actionsPhrase);

    if (startIndex !== -1) {
      isActionParagraph = true;
      const sentence = text.substring(startIndex + actionsPhrase.length);
      const words = sentence.split(" ");
      let foundName;

      if (names) {
        foundName = names.find(name => words.join(' ').includes(name));
      }

      if (foundName) {
        const nameIndex = words.indexOf(foundName.split(' ')[0]);
        const action = words.slice(nameIndex + foundName.split(' ').length).join(' ');

        const isDuplicate = actionList.some(item => item.name === foundName && item.action === action);
        if (!isDuplicate) {
          actionList.push({ name: foundName, action });
        }
      }
    }
  });

  if (isActionParagraph) {
    const existingTable = findExistingActionTable(body);
    if (existingTable) {
      const existingActions = getExistingActions(existingTable);
      for (const actionItem of actionList) {
        const isDuplicate = checkIfActionExistsInTable(existingActions, actionItem);
        if (!isDuplicate) {
          populateActionInTable(existingTable, actionItem);
        }
      }
      Logger.log('Actions populated from paragraphs.');
    } else {
      const secondTable = createSecondTable(body);
      for (const actionItem of actionList) {
        populateActionInTable(secondTable, actionItem);
      }
      Logger.log('Actions populated from paragraphs. New action item table created.');
    }
  } else {
    Logger.log('No action items found in paragraphs.');
  }
}

// function to run both actionsFromParagraphs and actionsFromTable functions for the "Populate Actions" button.
function runBothActionItems() {
  // Run actionsFromParagraphs function
  actionsFromParagraphs();

  // Run actionsFromTable function
  actionsFromTable();
}

