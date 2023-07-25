// Purpose: 
// 1. Search agendas for action items to be completed and populate in the Action Tracking Google Sheet.
// Future development: preserve links in task items when copying to action table

// This script is developed for use as either as a Google Apps Script container script or as a Google Apps Script library script. 
// Google Apps Script container script: a script that is bound to a specific file, such as a Google Sheets, Google Docs, or Google Forms file. This container script acts as the file's custom script, allowing users to extend the functionality of the file by adding custom functions, triggers, and menu items to enhance its behavior and automation.

// Google Apps Script library script: a self-contained script that contains reusable functions and can be attached to multiple projects or files. By attaching the library script to different projects, developers can access and use its functions across various files, enabling code sharing and improving code maintenance and version control.

//////////////////////////////////////////////////

// To use this script as a library script, follow the instructions below:
// 1. Obtain the script ID of the inDocActionItems library script.
//  script ID: 1kBbrOJCXewvSixfq1yR8d-lEtgDG5yzD9-pqPeuC9ugLka7gQULwkBH_ <-- verify current library script id by checking in Project settings (gear icon).
// 2. Open the container document where you want to use this script.
// 3. Click on the "Project settings" gear icon in the script editor.
// 4. In the "Script ID" field, replace the existing script ID with the script ID of the inDocActionItems library script.
// 5. Click "Save" to update the script ID.

// function to use in your container to call inDocActionItems as a library script: 
// function libraryCall(){
//  inDocActionItems.runBothActionItems();
//  inDocActionItems.customMenu();
//}

///////////////////////////////////////////////////

// function to create custom menu with buttons on document open
function onOpen() {
    DocumentApp.getUi() // 
        .createMenu('Custom Menu')
        .addItem('Populate Actions','actionsFromParagraphs')      
        .addToUi();
  }
  
  // Primary function that extracts action items from paragraphs in the document and populates them into an action item table.
  function actionsFromParagraphs() {
    const document = DocumentApp.getActiveDocument();
    const body = document.getBody();
    const actionsPhrases = ['Action:', 'Action : ', 'Action-', 'Action -','action:', 'action : ', 'action-', 'action -'];
    const actionList = [];
  
    const paragraphs = body.getParagraphs();
    let isActionParagraph = false;
    paragraphs.forEach(paragraph => {
      const text = paragraph.getText();
      actionsPhrases.forEach(actionsPhrase => {
        const startIndex = text.indexOf(actionsPhrase);
  
        if (startIndex !== -1) {
          isActionParagraph = true;
          const sentence = text.substring(startIndex + actionsPhrase.length);
          const words = sentence.split(" ");
  
          actionList.push({ action: '□ ' + words.join(' ') });
        }
      });
    });
  
    if (isActionParagraph) {
      let existingTable = findExistingActionTable(body);
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
        existingTable = createSecondTable(body);
        for (const actionItem of actionList) {
          populateActionInTable(existingTable, actionItem);
        }
        Logger.log('Actions populated from paragraphs. New action item table created.');
      }
    } else {
      Logger.log('No action items found in paragraphs.');
    }
  }
  
  
  // Helper function that creates a new action item table in the document.
  function createSecondTable(body) {
    const table = body.appendTable();
    return table;
  }
  
  // Helper function that retrieves existing action items from the table in the document.
  function getExistingActions(table) {
    const existingActions = [];
    const numRows = table.getNumRows();
  
    for (let i = 0; i < numRows; i++) { // Start from 0 as there is no header row now
      const row = table.getRow(i);
      const actionContent = row.getCell(0).getText();
      existingActions.push({ action: actionContent });
    }
  
    return existingActions;
  }
  
  // Helper function that checks if an action item already exists in the table.
  function checkIfActionExistsInTable(existingActions, actionItem) {
    for (const existingAction of existingActions) {
      if (existingAction.action === actionItem.action) {
        return true;
      }
    }
    return false;
  }
  
  // Helper function that populates an action item into the table.
  function populateActionInTable(table, actionItem) {
    const numRows = table.getNumRows();
  
    for (let i = 0; i < numRows; i++) { // Start from 0 as there is no header row now
      const row = table.getRow(i);
      const cell = row.getCell(0);
  
      if (cell.getText() === '') {
        cell.setText(actionItem.action);
        return;
      }
    }
  
    const newRow = table.appendTableRow();
    newRow.appendTableCell().setText(actionItem.action); // Action cell
    table.setBorderColor('#FFFFFF'); // Set border color to be transparent
  }
  
  // Primary function that extracts action items from an existing table in the document and populates them into the Action Items table.
  function actionsFromTable() {
    const document = DocumentApp.getActiveDocument();
    const body = document.getBody();
    let existingTable = findExistingActionTable(body);
  
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
  
  // Helper function that searches for an existing action item table in the document.
  function findExistingActionTable(body) {
    const tables = body.getTables();
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      const numRows = table.getNumRows();
  
      if (numRows > 0) {
        const firstRowFirstCellText = table.getRow(0).getCell(0).getText();
  
        if (firstRowFirstCellText.startsWith('□')) {
          return table;
        }
      }
    }
  
    return null;
  }
  
  // Helper function: extractActionsFromTable
  function extractActionsFromTable(table) {
    const actionList = [];
    
    if (table) {
      const numRows = table.getNumRows();
  
      for (let i = 0; i < numRows; i++) { // Start from 0 as there is no header row now
        const row = table.getRow(i);
        const actionContent = row.getCell(0).getText();
        actionList.push({ action: actionContent });
      }
    }
  
    return actionList;
  }
  
  // Run the actionsFromParagraphs function
  actionsFromParagraphs();
  
  // Run the actionsFromTable function
  actionsFromTable();
  
  // function to run both actionsFromParagraphs and actionsFromTable functions for the "Populate Actions" button.
  function runBothActionItems() {
    // Run actionsFromParagraphs function
    actionsFromParagraphs();
  
    // Run actionsFromTable function
    actionsFromTable();
  }
  
  