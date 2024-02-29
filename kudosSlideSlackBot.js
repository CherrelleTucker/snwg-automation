/* Google Apps Script for Slack Notifications

Prerequisites
1. Google account with access to Google Apps Script.
2. Slack workspace with a configured Slack app.
3. Google Calendar with events following the format 'PI XX.X Sprint X'.
4. Google Drive folder containing Sprint Review files.

Script Functions
1. main(): Main function to execute the script.
2. findCurrentSprintEvent(): Finds the current sprint event for today.
3. findSprintFileInDrive(sprintTitle): Finds the corresponding Sprint Review file in Google Drive.
4. doPost(e): Handles POST requests from Slack events.
5. fetchPostsAndCreateTextBoxes(fileId): Fetches posts from Slack channel and creates text boxes in Google Slides.
6. postThankYouReply(thread_ts): Posts a thank-you reply to a specific message in the Slack channel.
7. prepareMessageBoxText(message, userInfoMap): Prepares text content for the text box based on the message format.
8. fetchUserInfo(messages): Fetches user information for all mentioned users in the messages.
9. getRandomNumber(min, max): Generates a random number within a range.
10. createColoredTextBoxInPresentation(message, index): Creates a text box with a message in Google Slides.

Outputs
1. Updated Google Slides presentation with text boxes containing Slack messages.

Post-Execution
After execution, the script:
1. Retrieves the current sprint event from Google Calendar.
2. Searches for the corresponding Sprint Review file in Google Drive.
3. Fetches posts from the Slack channel and inserts them into the Google Slides presentation.

Troubleshooting
1. If no events are found for today on the calendar, ensure events are properly scheduled.
2. If no file is found for the current sprint in the specified folder, verify the folder ID.
3. If there are errors fetching messages or posting replies, check network connectivity and API permissions.

Notes
1. Adjust the regex pattern in findCurrentSprintEvent() if the event title format changes.
2. Modify the folderId variable to match the ID of the Google Drive folder.
3. Customize the script further as needed for specific requirements.
*/

// Required global variables
var TOKEN = 'xoxb-andSomeOtherNumbers'; // Slack Bot User OAuth Access Token
var KUDOS_CHANNEL_ID = 'XXXXXXX'; // #kudos IMPACT channel id
var calendarId = 'xxxxxxx'; // IMPACT PI Google Calendar ID
var folderId = 'xxxxxxxx'; // IMPACT Sprint Review FY24 Google Drive Folder
var SLIDES_ID = ''; // This will be dynamically set based on the sprint file found
var currentSprintEvent = null;
var currentSprintFileId = null;

// Main function for executing the script. Find current sprint using the IMPACT PI Calendar. Find current sprint presentation in the Sprint Review folder. Execute 1function fetchPostsAndCreateTextBoxes
function main() {
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize today's date to start of the day for comparison

  // Log the date being searched for the sprint event
  Logger.log("Searching for current sprint event on " + today.toISOString());

  // Attempt to find the current sprint event for today
  var events = CalendarApp.getCalendarById(calendarId).getEventsForDay(today);
  if (events.length === 0) {
    Logger.log("No events found for today on the calendar.");
    return;
  }
  
  // Log all events found for today for diagnostic purposes
  Logger.log("Events found on the calendar for today:");
  events.forEach(function(event) {
    Logger.log(event.getTitle());
  });

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    // Log each event title to inspect for variations or inconsistencies
    Logger.log("Inspecting event title: " + event.getTitle());
    if (/^PI \d{2}\.\d{1,2} Sprint \d/.test(event.getTitle())) {
      currentSprintEvent = event;
      Logger.log("Current sprint event found: " + event.getTitle());
      break;
    }
  }

  if (!currentSprintEvent) {
    Logger.log("No current sprint event found for today that matches the expected format.");
    return;
  }

  // Extracting the sprint information from the event title
  var sprintInfo = currentSprintEvent.getTitle().match(/^PI (\d{2})\.(\d{1,2}) Sprint (\d)/);
  if (!sprintInfo) {
    Logger.log("Failed to extract sprint information from event title: " + currentSprintEvent.getTitle());
    return;
  }
  var sprintTitle = "IMPACT Sprint Review_PI " + sprintInfo[1] + "." + sprintInfo[2] + "." + sprintInfo[3];

  // Log the specific file name being searched for
  Logger.log("Searching for file with title containing: " + sprintTitle);

  // Attempt to find the corresponding file in the Drive folder
  var files = DriveApp.getFolderById(folderId).getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().includes(sprintTitle)) {
      currentSprintFileId = file.getId();
      Logger.log("Found file for current sprint: " + file.getName());
      Logger.log("Using file ID: " + currentSprintFileId); // Log the file ID
      break;
    }
  }

  if (!currentSprintFileId) {
    Logger.log("No file found for the current sprint in the specified folder.");
    return;
  }

  // Proceed with updating the slides using the file ID
    SLIDES_ID = currentSprintFileId; // Set SLIDES_ID here
  fetchPostsAndCreateTextBoxes(currentSprintFileId); // Call fetchPostsAndCreateTextBoxes() with the retrieved file ID
}

// Finds the current sprint event in the Google Calendar.
// @return {string} The title of the current sprint event.
function findCurrentSprintEvent() {
  var now = new Date();
  var events = CalendarApp.getCalendarById(calendarId).getEventsForDay(now);
  var sprintEventPattern = /PI\d{2}\.\d Sprint \d/; // Adjust regex pattern if needed

  for (var i = 0; i < events.length; i++) {
    if (sprintEventPattern.test(events[i].getTitle())) {
      return events[i].getTitle();
    }
  }

  throw new Error('No current sprint event found for today.');
}

/**
 * Finds the corresponding sprint file in Google Drive.
 * @param {string} sprintTitle - The title of the sprint file.
 * @return {string} The ID of the found sprint file.
 */
function findSprintFileInDrive(sprintTitle) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var fileTitlePattern = new RegExp(sprintTitle.replace(/ /g, '.*'), 'i'); // Convert sprint title to regex, making it flexible

  while (files.hasNext()) {
    var file = files.next();
    if (fileTitlePattern.test(file.getName())) {
      return file.getId(); // Return the ID of the found file
    }
  }

  throw new Error('No file found for the current sprint in the specified folder.');
}

// Define five distinct light colors using RGB values for the text boxes
var COLORS = [
  {red: 255, green: 255, blue: 224}, // Light Yellow
  {red: 144, green: 238, blue: 144}, // Light Green
  {red: 173, green: 216, blue: 230}, // Light Blue
  {red: 255, green: 182, blue: 193}, // Light Pink
  {red: 216, green: 191, blue: 216}  // Light Purple
];

// Helper function to convert RGB values to Hex
function rgbToHex(r, g, b) {
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

// Handles POST requests from Slack events
function doPost(e) {
  var json = JSON.parse(e.postData.contents);

  // Handle verification challenge for Slack Event API
  if (json.type === 'url_verification' && json.challenge) {
    return ContentService.createTextOutput(json.challenge);
  }
  
  // Security check: verify the token
  if (json.token !== TOKEN) {
    return ContentService.createTextOutput('Invalid token').setMimeType(ContentService.MimeType.TEXT);
  }

  // Process a message event in the #kudos channel
  if (json.event && json.event.type === 'message' && json.event.channel === KUDOS_CHANNEL_ID) {
    console.log('Message from #kudos channel:', json.event.text);
    // You can add additional processing here if needed

    return ContentService.createTextOutput(JSON.stringify({ "response": "Message received" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default response for non-kudos messages
  return ContentService.createTextOutput(JSON.stringify({ "response": "Not a #kudos message" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Fetches posts from the current day from the Slack channel and creates text boxes in Google Slides.
 * @param {string} fileId - The ID of the Google Slides presentation file.
 */
function fetchPostsAndCreateTextBoxes(fileId) {
  // Get the current date in the same format as Slack timestamps
  var currentDate = new Date();
  currentDate.setHours(0, 0, 0, 0); // Set to start of the current day
  var unixStartTime = Math.floor(currentDate.getTime() / 1000); // Convert to Unix timestamp
  
  // Increase the limit if necessary to make sure you capture all messages from the day
  var apiUrl = 'https://slack.com/api/conversations.history?channel=' + KUDOS_CHANNEL_ID + '&limit=20';
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + TOKEN
    },
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(apiUrl, options);
  var json = JSON.parse(response.getContentText());
  var copiedMessageCount = 0; // Initialize the count of copied messages

  if (json.ok && json.messages) {
    // Filter messages to include only those from the current day
    var todayMessages = json.messages.filter(message => {
      // Convert Slack timestamp to milliseconds and then to date for comparison
      var messageDate = new Date(parseFloat(message.ts) * 1000);
      return messageDate >= currentDate;
    });

    if (todayMessages.length > 0) {
      var userInfoMap = fetchUserInfo(todayMessages);
      
      todayMessages.forEach((message, index) => {
        var messageText = prepareMessageBoxText(message, userInfoMap);
        createColoredTextBoxInPresentation(messageText, index);
        // Increment the count of copied messages
        copiedMessageCount++;
        // Call the function to post a thank-you reply to each message
        postThankYouReply(message.ts); // Use 'ts' value as 'thread_ts' to reply in thread
      });

      // Log the count of copied messages
      Logger.log("Number of messages copied over: " + copiedMessageCount);
    } else {
      console.log('No messages found for the current day.');
    }
  } else {
    console.log('Error fetching messages:', json.error);
  }
}

/*
 * Posts "Thank you for recognizing your teammate!" reply to a specific message in the #kudos channel.
 * @param {string} thread_ts - The timestamp of the message to reply to.
 */
function postThankYouReply(thread_ts) {
  var postMessageUrl = 'https://slack.com/api/chat.postMessage';
  var payload = {
    'channel': KUDOS_CHANNEL_ID,
    'text': 'Thank you for recognizing your teammate!',
    'thread_ts': thread_ts // Ensures the message is posted as a reply
  };
  
  var options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + TOKEN,
      'Content-Type': 'application/json; charset=UTF-8'
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(postMessageUrl, options);
  var jsonResponse = JSON.parse(response.getContentText());
  if (!jsonResponse.ok) {
    console.log('Error posting thank you reply:', jsonResponse.error);
  }
}

/**
 * Prepares text content for the text box based on the message format.
 * @param {Object} message - The Slack message object.
 * @param {Object} userInfoMap - The map containing user information.
 * @return {string} The prepared text for the text box.
 */
function prepareMessageBoxText(message, userInfoMap) {
  var finalText = "";

  // Detect if the message is from KudosBot
  if (message.text.includes('received kudos from')) {
    // Extract the users and appreciation text from the KudosBot message
    var parts = message.text.split("\n"); // Split by new lines to separate main content from the note
    var appreciationText = parts[1].trim(); // The "few words of appreciation"
    var givingUserMatch = message.text.match(/from @([^ ]+)/);
    var receivingUserMatch = message.text.match(/@([^ ]+) received kudos/);

    if (givingUserMatch && receivingUserMatch) {
      var givingUser = givingUserMatch[1].trim(); // User giving kudos
      var receivingUser = receivingUserMatch[1].trim(); // User receiving kudos

      finalText = `Kudos for: @${receivingUser}\n${appreciationText}\n- @${givingUser}`;
    }
  } else {
    // Handle standard user posts
    var textWithUserNames = message.text.replace(/<@([A-Z0-9]+)>/g, function(match, userId) {
      var userName = userInfoMap[userId] ? userInfoMap[userId] : "Unknown User";
      return "@" + userName;
    });

    var posterName = userInfoMap[message.user] ? userInfoMap[message.user] : "Unknown Poster";
    finalText = `${textWithUserNames}\n- ${posterName}`;
  }

  return finalText;
}

/**
 * Fetches user information for all mentioned users in the messages.
 * @param {Object[]} messages - The array of Slack message objects.
 * @return {Object} The map containing user information.
 */
function fetchUserInfo(messages) {
  var userIds = new Set();
  messages.forEach(message => {
    userIds.add(message.user); // Add user who posted the message
    // Add mentioned users
    var matches = message.text.match(/<@([A-Z0-9]+)>/g);
    if (matches) {
      matches.forEach(match => userIds.add(match.slice(2, -1)));
    }
  });

  var userInfoMap = {};
  userIds.forEach(userId => {
    var userInfoUrl = 'https://slack.com/api/users.info?user=' + userId;
    var options = {
      'headers': {
        'Authorization': 'Bearer ' + TOKEN
      },
      'muteHttpExceptions': true
    };
    var userInfoResponse = UrlFetchApp.fetch(userInfoUrl, options);
    var userInfoJson = JSON.parse(userInfoResponse.getContentText());
    if (userInfoJson.ok) {
      userInfoMap[userId] = userInfoJson.user.real_name || userInfoJson.user.name;
    }
  });
  return userInfoMap;
}

/**
 * Generates a random number within a range.
 * @param {number} min - The minimum value of the range.
 * @param {number} max - The maximum value of the range.
 * @return {number} The random number generated.
 */
function getRandomNumber(min, max) {
  return Math.random() * (max - min) + min;
}

/**
 * Creates a text box with a message in the Google Slides presentation.
 * @param {string} message - The message to display in the text box.
 * @param {number} index - The index of the text box.
 */
function createColoredTextBoxInPresentation(message, index) {
  var presentation = SlidesApp.openById(SLIDES_ID);
  var slide = presentation.getSlides()[4]; // Assuming you are still working with the 5th slide

  // Define the bounds for the random position
  // These values should be adjusted based on the slide size and desired text box placement area
  var minX = 50; // Minimum X position
  var maxX = 650; // Maximum X position, assuming slide width is around 700
  var minY = 50; // Minimum Y position
  var maxY = 350; // Maximum Y position, assuming slide height is around 400

  // Generate random positions within the defined bounds
  var posX = getRandomNumber(minX, maxX);
  var posY = getRandomNumber(minY, maxY);

  // Define text box size
  var textBoxWidth = 100; // Adjust as needed
  var textBoxHeight = 50; // Adjust as needed

  // Choose a random color for the text box
  var colorIndex = Math.floor(Math.random() * COLORS.length);
  var color = COLORS[colorIndex];
  
  // Create the text box at a random position with the chosen color
  var textBox = slide.insertTextBox(message, posX, posY, textBoxWidth, textBoxHeight);
  textBox.getFill().setSolidFill(rgbToHex(color.red, color.green, color.blue));
  textBox.getText().getTextStyle().setFontSize(8); // Adjust font size as needed
  textBox.getText().getTextStyle().setFontFamily("Roboto"); // Set the font family to Roboto
}
