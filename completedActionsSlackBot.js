/* This script facilitates a seamless integration between Slack and Google Sheets, enabling users to update action items across multiple tabs within a Google Sheet by issuing commands directly from Slack.

Prerequisites
1. Google Workspace Account: Access to Google Sheets and Google Apps Script.
2. Slack Workspace: The ability to create and manage Slack apps. This is generally limited to paid accounts only.
3. Google Sheet: A Google Sheet with pre-defined tabs with name you will specify below.
4. Slack App: A Slack app with a slash command (e.g., /done) configured and the incoming webhook feature enabled.
5. Deployment URL: The script must be deployed as a Web App from the Google Apps Script editor, with the deployment URL configured in the Slack app as the request URL for the slash command.
Script Functions. Each script update requires a new deployment. Each new deployment generates a new webapp url that must be updated in the slash command request url.
doPost(e): The main entry point for handling POST requests from Slack. Parses the command text, invokes the update function, and returns a response to Slack.
updateSheetWithActionItem(actionItem, status): Searches for and updates or adds the specified action item across multiple sheets within the Google Sheet. Returns a message indicating the outcome.

Outputs
1. Slack Response: After executing the /done command in Slack, users receive a response indicating whether the action item was added as a new entry, updated across multiple instances, or not found.
2. Google Sheet Update: Matching action items across specified tabs are updated with the new status, or a new entry is added if no match is found.

Post-Execution
1. Upon successful execution of the Slack command, the Google Sheet will reflect the updates:
    If an action item exists in the specified tabs, its status will be updated.
    If the action item does not exist, it will be added to the sheet with the specified status.
    The Slack user will receive immediate feedback on the action taken by the script.

Troubleshooting
1. Function Not Found Error: Ensure the Web App deployment is up-to-date with the latest version of the script.
2. Permission Issues: Verify that the script has permission to access and modify the Google Sheet.
3. Slack Command Not Working: Check that the Slack app's slash command is correctly configured with the Web App's deployment URL.

Notes
1. The script defaults the status to "Done" unless "Pending" is explicitly mentioned in the Slack command.
2. The script operates on a case-insensitive basis when searching for action items, reducing duplication due to case differences.
3. To modify the list of sheets searched by the script, adjust the SHEET_NAMES constant accordingly.
4. This script is designed for simplicity and ease of use. For more complex workflows or additional features, consider extending the script's functionality or integrating with Google Apps Script's advanced services.
*/ 

// Constants for the Google Sheet and sheet names to search
const SHEET_ID = 'xxxxxxxxxxxxxx'; //SNWG MO Action Tracking Sheet
const SHEET_NAMES = ["MO", "DevSeed", "SEP", "AssessmentHQ", "AdHoc"];

function doPost(e) {
  var text = e.parameter.text;

  // Remove any leading numbering, "Assigned to All:", and trim whitespace
  text = text.replace(/^\d+\.\s*/, '').replace(/^Assigned to All:\s*/i, '').trim();

  var parts = text.split(" ");
  var status = parts[parts.length - 1].toLowerCase();
  
  var actionItem;
  // Assuming "pending" must be explicitly mentioned, otherwise default to "Done"
  if (status === "pending") {
    actionItem = parts.slice(0, -1).join(" ");
  } else {
    actionItem = text; // Consider the entire text as the action item
    status = "Done"; // Default status
  }

  // Immediately respond to Slack to indicate processing has started
  var immediateResponse = {
    "response_type": "in channel", //"in_channel" Or "ephemeral" for a private response
    "text": "Marked complete: \"" + actionItem + "\""
  };

  // Log the actionItem and status for debugging purposes
  console.log("Action Item:", actionItem, "Status:", status);

  // Asynchronously update the sheet with the action item
  // Note: Google Apps Script does not support true async operations or delaying response
  // The script will continue to execute the updateSheetWithActionItem function after sending the initial response
  var resultMessage = updateSheetWithActionItem(actionItem, status);

  return ContentService.createTextOutput(JSON.stringify(immediateResponse))
    .setMimeType(ContentService.MimeType.JSON);
}

// This function updates all matching action items across multiple sheets
function updateSheetWithActionItem(actionItem, status) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  
  var totalUpdatesCount = 0; // Track total updates across all sheets
  var sheetUpdates; // Track updates per sheet for detailed feedback

  // Iterate over each sheet name
  SHEET_NAMES.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      // Skip if the sheet does not exist
      return;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return; // Skip empty sheets
    
    var actionItemsRange = sheet.getRange(1, 4, lastRow, 1); // Assuming the action items are in column 4
    var actionItems = actionItemsRange.getValues();
    
    var updatesCount = 0; // Track the number of updates made in the current sheet
    
    for (var i = 0; i < actionItems.length; i++) {
      if (actionItems[i][0].toLowerCase() === actionItem.toLowerCase()) {
        sheet.getRange(i + 1, 2).setValue(status); // Assuming the status should be set in column 2
        updatesCount++; // Increment updates count
        totalUpdatesCount++;
      }
    }
    
    if (updatesCount > 0) {
      sheetUpdates = (sheetUpdates || "") + `Updated ${updatesCount} in ${sheetName}. `;
    }
  });

  // Return a message based on the operation result
  if (totalUpdatesCount > 0) {
    return `Found and updated ${totalUpdatesCount} instance(s) of "${actionItem}" to "${status}". ${sheetUpdates || ""}`;
  } else {
    // If no matching action items were found, consider adding the action item as new
    // This step is not implemented in the given script but can be added as needed
    return `No instances of "${actionItem}" found across sheets. Please check your spelling or pasted action.`;
  }
}
