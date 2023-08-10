// Purpose:
// This script processes a Google Document to identify and catalog action items. 
// It looks for attendees' names and action items in two primary ways: directly within the document's paragraphs and within existing tables. 
// If the script identifies action items from paragraphs, it either adds them to an existing action table or creates a new table if one doesn't exist. 
// When examining tables, the script checks for duplicates to avoid adding redundant action items. 
// The main function processDocument(documentId) initiates this procedure by examining both paragraphs and tables for action items.

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
  function actionsFromTable(documentId) {
    const document = DocumentApp.openById(documentId);
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
  function actionsFromParagraphs(documentId) {
    const document = DocumentApp.openById(documentId);
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
  function runBothActionItems(documentId) {
    // Run actionsFromParagraphs function
    actionsFromParagraphs(documentId);
  
    // Run actionsFromTable function
    actionsFromTable(documentId);
  }
  
  function processDocument(documentId){
    runBothActionItems(documentId)
  }
  