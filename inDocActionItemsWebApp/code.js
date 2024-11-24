/*
Script Name: inDocActionItemsWebApp

  Description: 
  This script is designed to parse a Google Document to identify and extract action items, then populate them in an Action Tracking table within the document. It processes action items mentioned in both paragraph text and existing tables, sorts them by the owner's name, and ensures no duplication in the final action item table.

  Prerequisites: 
  - A Google Document containing action items in paragraphs or tables.
  - The Google Document should have a section with attendees listed, starting with 'Attendees:'.

  Setup:
  1. Open the Google Apps Script editor linked to the Google Document.
  2. Paste this entire script into the script editor.
  3. Save the script.

  Execution: 
  To run the script, either use the `testScriptWithDocumentUrl` function with a valid Google Document URL or call the `processDocument` function with a specific document ID.

  Script Functions: 
  - `extractAttendees(body)`: Extracts attendees' names from the document.
  - `createSecondTable(body)`: Creates a new action item table.
  - `getExistingActions(table)`: Retrieves existing action items from a table.
  - `checkIfActionExistsInTable(existingActions, actionItem)`: Checks for the existence of an action item in the table.
  - `populateActionInTable(table, actionItem)`: Populates an action item into the table.
  - `extractActionsFromTable(table)`: Extracts action items from an existing table.
  - `findExistingActionTable(body)`: Searches for an existing action item table.
  - `sortActionItemsByName(actionItems)`: Sorts action items by the owner's name.
  - `actionsFromTable(documentId)`: Extracts action items from a table and populates them into the Action Items table.
  - `actionsFromParagraphs(documentId)`: Extracts action items from paragraphs and populates them into the Action Items table.
  - `runBothActionItems(documentId)`: Runs both `actionsFromParagraphs` and `actionsFromTable`.
  - `processDocument(documentId)`: Main function to process the document.
  - `extractDocumentIdFromUrl(url)`: Extracts the document ID from a Google Docs URL.
  - `testScriptWithDocumentUrl(url)`: Test function for running the script with a document URL.

  Outputs: 
  - An updated Google Document with an Action Tracking table containing sorted action items from both paragraphs and tables.

  Post-Execution: 
  - Review the populated Action Tracking table in the Google Document for accuracy.

  Troubleshooting:
  - Ensure the Google Document contains the 'Attendees:' section.
  - Check for proper formatting of action items in the document.
  - Use logging statements (`Logger.log`) to debug issues with specific functions.

  Notes:
  - Future development includes preserving links in task items when copying to the action table.
  - The script currently assumes a specific format for action items and attendees.
*/

///////////////////////////////////////////////////

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
    const ownerCell = row.getCell(1);
    const actionCell = row.getCell(2);

    // Check if both Owner and Action cells are empty
    if (ownerCell.getText().trim() === '' && actionCell.getText().trim() === '') {
      // Populate the empty row
      row.getCell(0).setText('Not Started'); // Populate Status cell, assuming you want to set it to 'Not Started'
      ownerCell.setText(actionItem.name);
      actionCell.setText(actionItem.action);
      return; // Exit the function as we have populated the action item
    }
  }

  // If no empty row was found, append a new row at the end
  const newRow = table.appendTableRow();
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

// Helper function to sort action items by name.
function sortActionItemsByName(actionItems) {
  return actionItems.sort((a, b) => a.name.localeCompare(b.name));
}

// Helper function to check if any valid action items were found
function validateActionItems(actionList) {
    if (!actionList || actionList.length === 0) {
        Logger.log('No action items found');
        throw new Error('No action items found');
    }
    
    // Validate each action item has required properties
    const invalidItems = actionList.filter(item => 
        !item.name || !item.action || 
        item.name.trim() === '' || 
        item.action.trim() === ''
    );
    
    if (invalidItems.length > 0) {
        Logger.log('Invalid action items found');
        throw new Error('Invalid action items found');
    }
    
    return true;
}

// Modified actionsFromParagraphs function with validation
function actionsFromParagraphs(documentId) {
    try {
        const document = DocumentApp.openById(documentId);
        const body = document.getBody();
        const actionsPhrase = /Action:/i;
        let actionList = [];
        
        const paragraphs = body.getParagraphs();
        
        paragraphs.forEach(paragraph => {
            const text = paragraph.getText();
            const actionIndex = text.search(actionsPhrase);
            
            if (actionIndex !== -1) {
                const actionText = text.substring(actionIndex + 7).trim();
                const words = actionText.split(' ');
                
                if (words.length > 1) { // Ensure we have both owner and action
                    const potentialOwner = words[0];
                    const actionDescription = words.slice(1).join(' ').trim();
                    
                    if (actionDescription) {
                        const isDuplicate = actionList.some(item => 
                            item.name === potentialOwner && 
                            item.action === actionDescription
                        );
                        
                        if (!isDuplicate) {
                            actionList.push({ 
                                name: potentialOwner, 
                                action: actionDescription 
                            });
                        }
                    }
                }
            }
        });
        
        // Validate action items before proceeding
        validateActionItems(actionList);
        
        // Sort and populate the table if validation passes
        actionList = sortActionItemsByName(actionList);
        const existingTable = findExistingActionTable(body) || createSecondTable(body);
        
        actionList.forEach(actionItem => {
            const existingActions = getExistingActions(existingTable);
            const isDuplicate = checkIfActionExistsInTable(existingActions, actionItem);
            if (!isDuplicate) {
                populateActionInTable(existingTable, actionItem);
            }
        });
        
        Logger.log('Actions successfully populated from paragraphs.');
        return true;
        
    } catch (error) {
        Logger.log('Error in actionsFromParagraphs: ' + error.message);
        throw error;
    }
}

// Modified actionsFromTable function with validation
function actionsFromTable(documentId) {
    try {
        const document = DocumentApp.openById(documentId);
        const body = document.getBody();
        const existingTable = findExistingActionTable(body);
        
        if (existingTable) {
            let actionList = extractActionsFromTable(existingTable);
            
            // Validate action items before proceeding
            validateActionItems(actionList);
            
            // Sort and populate if validation passes
            actionList = sortActionItemsByName(actionList);
            
            const existingActions = getExistingActions(existingTable);
            for (const actionItem of actionList) {
                const isDuplicate = checkIfActionExistsInTable(existingActions, actionItem);
                if (!isDuplicate) {
                    populateActionInTable(existingTable, actionItem);
                }
            }
            
            Logger.log('Actions successfully populated from table.');
            return true;
            
        } else {
            Logger.log('No existing action item table found.');
            throw new Error('No existing action item table found');
        }
        
    } catch (error) {
        Logger.log('Error in actionsFromTable: ' + error.message);
        throw error;
    }
}

// Modified runBothActionItems function with error handling
function runBothActionItems(documentId) {
    try {
        let actionsFound = false;
        
        try {
            actionsFromParagraphs(documentId);
            actionsFound = true;
        } catch (paragraphError) {
            Logger.log('No actions found in paragraphs: ' + paragraphError.message);
        }
        
        try {
            actionsFromTable(documentId);
            actionsFound = true;
        } catch (tableError) {
            Logger.log('No actions found in tables: ' + tableError.message);
        }
        
        if (!actionsFound) {
            throw new Error('No action items found');
        }
        
        return 'Success! Action items collected.';
        
    } catch (error) {
        Logger.log('Error in runBothActionItems: ' + error.message);
        throw error;
    }
}

// Keep the original processDocument function
function processDocument(documentId) {
    return runBothActionItems(documentId);
}

////////////////Testing ////////////////

// Testing Function to extract document ID from Google Docs URL
function extractDocumentIdFromUrl(url) {
  var match = /\/d\/([a-zA-Z0-9-_]+)/.exec(url);
  return match ? match[1] : null;
}

// Testing function that takes a Google Docs URL
function testScriptWithDocumentUrl(url) {
  var documentId = extractDocumentIdFromUrl(url);
  
  if (documentId) {
    Logger.log('Document ID extracted: ' + documentId);
    processDocument(documentId); // Assuming processDocument is your main function
  } else {
    Logger.log('Invalid URL or unable to extract document ID.');
  }
}

testScriptWithDocumentUrl('https://docs.google.com/document/d/1TklfjJCp4tt8scjSOlQn7pxrByI_6KTf0JMnD8y7bl8/edit');
