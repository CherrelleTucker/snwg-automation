/* Google Apps Script for Slack Notifications
This script fetches open action items from a specified Google Sheet and sends a formatted message to a designated Slack channel or user, ensuring stakeholders are kept informed about their current tasks.
It incorporates with a Slack WORKFLOW, not a Slack App. 

Prerequisites
1. Access to Google Sheets where the action items are tracked.
2. A Slack Workspace with an incoming webhook set up for the target channel/user. This is generally limited to paid accounts only.
3. A Slack Workflow set up with: 
  A. "Starts with a webhook"
    - Set up Variables: Key: "Action" Data type: text
        Example HTTP body in Google Apps Script:   
        {
            "Actions": "Example text"
          }
    - Web request URL (Generated in the Slack Workflow)
  B. "Send a message to @User"
    -"Add a message": "Insert a Variable" - "Actions"<- Set up in previous step. Save.
4. Google Apps Script environment linked to the Google Sheet.

Script Functions
1. prepareName(name): Normalizes a name string by removing specific characters and converting to lowercase.
2. getActionItems(sheetId, tabName, regex): Retrieves and filters unique action items based on the assignee, avoiding duplicates.
3. formatActionItems(actionItems): Formats the list of action items into a numbered list.
4. sendToSlack(actionOwnersName, slackWebhookUrl, message): Constructs and sends the payload to Slack via the specified webhook URL.
5. getMessage(actionOwnersName, actionItems): Determines the appropriate message based on the presence or absence of action items.
6. sendActionItemsToSlack(): Orchestrates the process from fetching action items to sending the notification in Slack.

Outputs
1. Slack Message: A detailed message sent to Slack, listing all open action items for the specified owner or indicating the absence of open items.

Post-Execution
After execution, the script:
1. Fetches action items from the Google Sheet.
2. Formats these items and sends a notification to Slack.
3. Updates processedActionDescriptions to avoid duplicate notifications in future executions.

Troubleshooting
1. Incorrect Slack Notifications: Ensure the regex pattern correctly matches the action owner's name or project.
2. No Notifications Sent: Verify the webhook URL and ensure the script has permission to access the Google Sheet and the internet.
3. Duplicate Notifications: Check processedActionDescriptions handling and ensure names/projects are uniquely identified.

Notes
1. The script is designed for simplicity and specific use cases. Modifications may be required for different naming conventions or additional functionality.
2. Ensure the webhook URL is kept secure and not exposed to unauthorized users.
3. Regularly review and update the action items in the Google Sheet to keep the notifications accurate and relevant.
*/

// Global variables for action owner's name components and Slack webhook URL
var sheetId = "xxxxxxxxxxxxxxxxxxxxx"; //SNWG MO Action Tracking Google Sheet
var tabName = "All Open Action Items";
var firstName = "First";
var lastName = "Last";
var projectName = "Project";
var slackWebhookUrl = "xxxxxxxxxxxxxxxxxxxxx";

// Keep track of processed action descriptions
var processedActionDescriptions = {};

// Function to prepare the name components by trimming them to lowercase
function prepareName(name) {
  return name.replace(/['’è]/g, '').toLowerCase().trim();
}

// Retrieves action items from the specified Google Sheet, avoiding duplicates
function getActionItems(sheetId, tabName, regex) {
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(tabName);
  var range = sheet.getDataRange();
  var values = range.getValues();
  var uniqueActionItems = [];

  for (var i = 1; i < values.length; i++) {
    var assignee = values[i][2];
    var description = values[i][3];

    if (regex.test(assignee)) {
      if (assignee.toLowerCase() === 'all') {
        description = "Assigned to All: " + description;
      }
      var formattedDescription = description.toLowerCase().trim();

      if (!processedActionDescriptions[formattedDescription]) {
        uniqueActionItems.push(description);
        processedActionDescriptions[formattedDescription] = true;
      }
    }
  }

  return uniqueActionItems;
}

// Formats action items text with a numbered list.
function formatActionItems(actionItems) {
  var actionItemsText = actionItems.map(function(item, index) {
    return (index + 1) + ". " + item;
  }).join("\n\n");

  return actionItemsText;
}

// Helper function responsible for constructing and sending the Slack message using the incoming webhook
function sendToSlack(actionOwnersName, slackWebhookUrl, message) {
  var payload = {
    "Actions": message
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(slackWebhookUrl, options);
}

// Determines the appropriate message to send based on the presence or absence of action items.
function getMessage(actionOwnersName, actionItems) {
  if (actionItems.length > 0) {
    // Prepare the action items text with a numbered list
    var actionItemsText = formatActionItems(actionItems);

    // Return the message with action items listed
    return "--------------------------------------------------\nGood morning, " + actionOwnersName + "! As of today, the following are your open action items from the SNWG MO:\n\n" + actionItemsText;
  } else {
    // If no unique action items found, return a message indicating so
    return "--------------------------------------------------\nGood morning, " + actionOwnersName + "! As of today, you have no recorded open action items from the SNWG MO.";
  }
}

// Main function that coordinates the entire process of interacting with the Google Sheet, formatting the data, and invoking sendToSlack
function sendActionItemsToSlack() {

  // Prepare the name components
  var preparedFirstName = prepareName(firstName);
  var preparedLastName = prepareName(lastName);
  var preparedProjectName = prepareName(projectName);

  // Regular expression to match variations of the action owner's name or "all"
  var regex = new RegExp("(?:@)?" + preparedFirstName + "(?:\\s+" + preparedLastName + ")?|(?:@)?" + preparedProjectName + "|all", "i");

  // Get unique action items for the action owner
  var actionItems = getActionItems(sheetId, tabName, regex);

  // Get the appropriate message to send
  var message = getMessage(firstName, actionItems);

  // Send the Slack message
  sendToSlack(firstName, slackWebhookUrl, message);
}
