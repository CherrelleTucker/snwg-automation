// Purpose: Container script to populate action items to an end 3 column action item table from either a preceding table or a paragraphed document. 
// To use as library script, add inDocActionItems script id to container document library, then transfer function libraryCall and function customMenu to the container script. 

// function to call inDocActionItems library script 
// function libraryCall(){
  //inDocActionItems.runBothActionItems();
  //inDocActionItems.organizeAttendees();
//}

// function to create custom menu with buttons
function customMenu() {
    DocumentApp.getUi() // 
        .createMenu('Custom Items')
        .addItem('Populate Actions','runBothActionItems')      
        .addItem ('Organize Attendees','organizeAttendees')
        .addToUi();
  }
  
  // Primary function: actionsFromParagraphs
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
  
  // Helper function: extractAttendees
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
  
  // Helper function: createSecondTable
  function createSecondTable(body) {
    const table = body.appendTable();
    const header = table.appendTableRow();
    header.appendTableCell("Status");
    header.appendTableCell("Owner");
    header.appendTableCell("Action");
    return table;
  }
  
  // Helper function: getExistingActions
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
  
  // Helper function: checkIfActionExistsInTable
  function checkIfActionExistsInTable(existingActions, actionItem) {
    for (const existingAction of existingActions) {
      if (existingAction.name === actionItem.name && existingAction.action === actionItem.action) {
        return true;
      }
    }
    return false;
  }
  
  // Helper function: populateActionInTable
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
  
  // Primary function: actionsFromTable
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
  
  // Helper function: findExistingActionTable
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
  
  // Secondary function: organizeAttendees 
  function organizeAttendees() {
    var document = DocumentApp.getActiveDocument();
    var body = document.getBody();
    var attendeesPhrase = 'Attendees: ';
    var attendeesText = body.findText(attendeesPhrase);
  
    if (attendeesText) {
      var attendeesElement = attendeesText.getElement();
      var attendeesString = attendeesElement.asText().getText().substring(attendeesPhrase.length);
      var attendeesArray = attendeesString.split(',').map(function (name) {
        return name.trim();
      });
  
      Logger.log('Attendees found:');
      Logger.log(attendeesArray);
  
      var sortedAttendees = attendeesArray.sort(function (a, b) {
        return a.charAt(0).localeCompare(b.charAt(0));
      });
  
      var sortedAttendeesString = sortedAttendees.join(', ');
  
      // Check if the attendees are already in the correct order
      if (sortedAttendeesString === attendeesString) {
        Logger.log('Attendees are already sorted.');
        return; // Stop further processing
      }
  
      attendeesElement.asText().setText(attendeesPhrase + sortedAttendeesString);
  
      Logger.log('Attendees sorted:');
      Logger.log(sortedAttendees);
    }
  }
  
  // Run the actionsFromParagraphs function
  actionsFromParagraphs();
  
  // Run the actionsFromTable function
  actionsFromTable();
  
  //Run both action functions
  function runBothActionItems() {
    // Run actionsFromParagraphs function
    actionsFromParagraphs();
  
    // Run actionsFromTable function
    actionsFromTable();
  }
  
  // Run the organizeAttendees function
  organizeAttendees();
  