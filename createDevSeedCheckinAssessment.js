// Purpose:
// This Google Apps Script is designed to automate the creation and population of a weekly agenda document for the "SNWG/DevSeed Weekly Check-In" meeting. The script operates by duplicating the most recent agenda, calculating the meeting date (next future Thursday), renaming the new agenda file with the meeting date, and updating the meeting date within the document content. Additionally, the script clears and repopulates a specific cell within the agenda table, appending bullet points for various updates related to IMPACT development, DevSeed, DCD, and SNWG. The primary goal of the script is to streamline the process of generating consistent agendas for recurring meetings.

// To note: 
// This script is developed as a Google Apps Script standalone script. It is developed to operate independently and does not require any external application or service to function. It is a self-contained piece of code with a time-based weekly trigger.

// To use:
// 1. Setting Up Google Drive Folder:Create a Google Drive folder where you want to store the generated agenda documents.
// Note down the Folder ID of this newly created folder. You can find the Folder ID in the URL of the Google Drive folder.
// 2. Accessing Google Apps Script: Open your Google Drive. Click on "New" > "More" > "Google Apps Script" to open the Google Apps Script editor.
// 3. Copying and Pasting Script: In the script editor, delete any existing code and paste the provided script into the editor.
// 4. Configuring Folder ID: Locate the line: var AGENDA_FOLDER_ID = "1Bvj1b1u2LGwjW5fStQaLRpZX5pmtjHSj";
// Replace the existing Folder ID with the Folder ID of the Google Drive folder you created in Step 1.
// 5. Running the Script: Save your script by clicking on the floppy disk icon or pressing Ctrl + S (Cmd + S on Mac). Close the script editor.
// 6. Setting up a Time-Driven Trigger: Open the Google Apps Script editor again by navigating to your Google Drive, right-clicking the script file, and selecting "Open with" > "Google Apps Script."
    // Click on the clock icon in the toolbar to open the Triggers page.
    // Click on "+ Add Trigger" at the bottom right corner.
    // Configure the trigger settings as follows:
    // Choose which function to run: createAndPopulateNewAgenda
    // Choose which deployment should run: "Head"
    // Select event source: "Time-driven"
    // Select type of time based trigger: "Day timer"
    // Select time of day: Choose the time of day when you want the agenda to be generated.
    // Click "Save."
// 7. Authorization and Permissions: The first time the script runs, it might ask for authorization. Follow the prompts to grant the necessary permissions to the script.
// 8. Viewing Logs:You can view logs to check the progress and any potential issues. In the script editor, click on the bug icon in the toolbar to open the logs.
// 9. Review the Output: The script will create a new duplicate agenda document in the specified folder with the appropriate date and populated cell.
// The script will now run automatically at the specified time every day, creating a new agenda document for your "SNWG/DevSeed Weekly Check-In" meetings. Make sure to review the generated documents to ensure everything is correct. 

//////////////////////////////////////////////////////

// Global constant for Agenda Destination Folder ID
var AGENDA_FOLDER_ID = "1Bvj1b1u2LGwjW5fStQaLRpZX5pmtjHSj"; // Assessment SCRIPT Development>SCRIPT RGT Meeting Notes Google Drive folder

// Helper function: Find the most recent agenda file
function findMostRecentAgenda() {
  var folder = DriveApp.getFolderById(AGENDA_FOLDER_ID);
  var query = 'title contains "DevSeed Checkin"';
  var agendaFiles = folder.searchFiles(query);
  var mostRecentAgenda = null;
  var mostRecentDate = new Date(0); // Initialize with an early date

  while (agendaFiles.hasNext()) {
    var agendaFile = agendaFiles.next();
    var fileName = agendaFile.getName();
    var dateString = fileName.match(/\d{4}-\d{2}-\d{2}/);
    
    if (dateString) {
      var agendaDate = new Date(dateString[0]);
      
      if (agendaDate > mostRecentDate) {
        mostRecentDate = agendaDate;
        mostRecentAgenda = agendaFile;
      }
    }
  }

  return mostRecentAgenda;
}

// Duplicate that agenda in the same folder
function duplicateMostRecentAgenda() {
  var mostRecentAgenda = findMostRecentAgenda();
  
  if (!mostRecentAgenda) {
    // No agenda found
    Logger.log("No agenda file found in the folder.");
    return;
  }

  // Duplicate the most recent agenda
  var folder = DriveApp.getFolderById(AGENDA_FOLDER_ID);
  var newAgenda = mostRecentAgenda.makeCopy("Agenda - Duplicate of " + mostRecentAgenda.getName(), folder);
  Logger.log("New agenda created: " + newAgenda.getName());
}

// Helper function: Get the meeting date - Next future Thursday
function getMeetingDate() {
  var today = new Date();
  var thursdayOffset = 4; // Thursday is the 5th day (0-indexed) of the week
  var daysToNextThursday = (thursdayOffset - today.getDay() + 7) % 7 || 7;
  var meetingDate = new Date(today.getTime());
  meetingDate.setDate(today.getDate() + daysToNextThursday);
  return meetingDate;
}

// Helper function: Rename to YYYY-MM-DD "SNWG/DevSeed Checkin" where YYYY-MM-DD is the meeting date
function renameAgendaWithMeetingDate(agenda, meetingDate) {
  var formattedDate = Utilities.formatDate(meetingDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var newName = formattedDate + " SNWG/DevSeed Checkin TEST";
  agenda.setName(newName);
}

// helper function to update meeting date within the newly created document
function updateAgendaDate(agendaFile,newDate) {
  var document = DocumentApp.openById(agendaFile.getId()); // Open the file as a Google Docs document
  var body = document.getBody();
  
  // Define a regular expression pattern to match dates in Month day, year format
  var datePattern = /\b(January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2}, \d{4}\b/;
    
  // Iterate through the paragraphs in the body
  var paragraphs = body.getParagraphs();
  for (var i = 0; i < paragraphs.length; i++) {
    var paragraph = paragraphs[i];
    var text = paragraph.getText();
    var match = text.match(datePattern);
    
    if (match) {
      // Replace the matched text with the new formatted date
      var formattedNewDate = Utilities.formatDate(newDate, Session.getScriptTimeZone(), "MMMM dd, yyyy");
      var updatedText = text.replace(datePattern, formattedNewDate);
      paragraph.setText(updatedText);
    }
  }
}

// Helper function to update specific placeholders
function updatePlaceholders(agendaFile) {
  var document = DocumentApp.openById(agendaFile.getId());
  var body = document.getBody();

  var placeholders = [
    { search: "IMPACT dev updates", replacement: "IMPACT dev updates (Iksha)" },
    { search: "DevSeed updates", replacement: "DevSeed updates (Will)" },
    { search: "DCD updates", replacement: "DCD updates (Essence)" }
  ];

  for (var i = 0; i < placeholders.length; i++) {
    var placeholder = placeholders[i];
    var textToSearch = placeholder.search;

    var foundElement = null;
    var searchResult = body.findText(textToSearch);

    while (searchResult) {
      if (searchResult.getElement().getParent().getType() === DocumentApp.ElementType.TABLE_CELL) {
        foundElement = searchResult.getElement();
        break;
      }
      searchResult = body.findText(textToSearch, searchResult);
    }

    if (foundElement) {
      var textElement = foundElement.asText();
      var startIndex = searchResult.getStartOffset();
      var endIndex = startIndex + textToSearch.length;

      textElement.deleteText(startIndex, endIndex);
      textElement.insertText(startIndex, placeholder.replacement);
    }
  }

  // No need for document.saveAndClose() here
}

// Helper function to get events for Katrina from the designated calendar
function getKatrinasEvents() {
  var now = new Date();
  var sixWeeksLater = new Date(now.getTime() + 42 * 24 * 60 * 60 * 1000); // Add 42 days to the current date
  sixWeeksLater.setMonth(now.getMonth() + 2);
  
  var calendarId = 'c_365230bc41700e58e23f74b286db1773d395e4bc6807c81a4c78658df5db423e@group.calendar.google.com';
  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(now, sixWeeksLater);
  
  var eventDetails = [];
  for (var i = 0; i < events.length; i++) {
    var event = events[i];

    // Check if the event title includes "Katrina"
    if (event.getTitle().includes("Katrina")) {
      var startDate = event.getStartTime();
      var endDate = event.getEndTime();

      if (event.isAllDayEvent()) {
        endDate = new Date(endDate.getTime() - 24*60*60*1000); // Subtract one day from the end date
      }

      var formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMMM d");
      var formattedEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MMMM d");

      var dateRange = (formattedStartDate === formattedEndDate) ? formattedStartDate : (formattedStartDate + " - " + formattedEndDate);
      
      eventDetails.push({
        title: event.getTitle(),
        date: dateRange
      });
    }
  }
  return eventDetails;
}

// Helper Function to populate Katrina's schedule in the agenda document
function populateKatrinasScheduleInAgenda(agendaFile) {
  var document = DocumentApp.openById(agendaFile.getId());
  var body = document.getBody();

  // This pattern matches "Katrina:" followed by any characters up to the end of the line
  var pattern = "Katrina:.*";

  var sectionStart = body.findText(pattern);
  if (sectionStart) {
    Logger.log("Found placeholder 'Katrina:' in the document.");

    var events = getKatrinasEvents();
    var eventsText = ""; // Start without "Katrina:"
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      eventsText += event.title + " - " + event.date + "\n";
    }

    // Replace the pattern with the events text
    body.replaceText(pattern, eventsText.trim());

  } else {
    Logger.log("Did not find placeholder 'Katrina:' in the document.");
  }
}

// Primary function: create and populate new agenda
function createAndPopulateNewAgenda() {
  // Find the most recent agenda
  var mostRecentAgenda = findMostRecentAgenda();

  if (!mostRecentAgenda) {
    // No agenda found
    Logger.log("No agenda file found in the folder.");
    return;
  }

  Logger.log("Most recent agenda: " + mostRecentAgenda.getName());

  // Duplicate the most recent agenda
  var newAgenda = mostRecentAgenda.makeCopy("Agenda - Duplicate of " + mostRecentAgenda.getName());
  Logger.log("New agenda created: " + newAgenda.getName());

  // Get the meeting date
  var meetingDate = getMeetingDate();
  Logger.log("Meeting date: " + meetingDate.toDateString());

  // Rename the new agenda with the meeting date
  renameAgendaWithMeetingDate(newAgenda, meetingDate);
  Logger.log("New agenda renamed to: " + newAgenda.getName());

  // Update the agenda date in the content
  updateAgendaDate(newAgenda, mostRecentAgenda.getDateCreated(), meetingDate);
  Logger.log("Agenda date updated.");

  // Populate Katrina's schedule in the agenda
  populateKatrinasScheduleInAgenda(newAgenda);
  Logger.log("Katrina's schedule updated.");

  // Update specific placeholders
  updatePlaceholders(newAgenda);
  Logger.log("Specific placeholders updated.");

}
