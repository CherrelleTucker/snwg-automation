// v1.0 Organize attendees, get action items. To recognize names, Action Owner must be listed in attendees.
// Issue: loses hyperlinked text when populating table


// Primary function: getActions
function getActions() {
  const document = DocumentApp.getActiveDocument();
  const body = document.getBody();
  const actionsPhrase = 'Action:';
  const actionList = [];

  const names = extractAttendees(body);

  // Search for actions in paragraphs
  const paragraphs = body.getParagraphs();
  paragraphs.forEach(paragraph => {
    const text = paragraph.getText();
    const startIndex = text.indexOf(actionsPhrase);

    if (startIndex !== -1) {
      const sentence = text.substring(startIndex + actionsPhrase.length);
      const words = sentence.split(" ");
      const foundName = names.find(name => words.join(' ').includes(name));

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

  const secondTable = secondTableExists(body) ? body.getTables()[0] : createSecondTable(body);
  const existingActions = getExistingActions(secondTable);

  for (const actionItem of actionList) {
    const isDuplicate = checkIfActionExistsInTable(existingActions, actionItem);

    if (!isDuplicate) {
      populateActionInTable(secondTable, actionItem);
    }
  }

  Logger.log('Actions populated.');
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

  const newRow = table.appendTableRow();
  newRow.appendTableCell().setText('Not Started'); // Status cell
  newRow.appendTableCell().setText(actionItem.name); // Owner cell
  newRow.appendTableCell().setText(actionItem.action); // Action cell
}

// Helper function: secondTableExists
function secondTableExists(body) {
  return body.getTables().length > 0;
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

// Helper function: getFirstEmptyRow
function getFirstEmptyRow(table) {
  for (let i = 0; i < table.getNumRows(); i++) {
    if (table.getCell(i, 1).getText().trim() === '') {
      return i;
    }
  }

  return table.getNumRows();
}

// Run the getActions function
getActions();

// Primary function: sortAttendees
function sortAttendees() {
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