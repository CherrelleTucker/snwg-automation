// Purpose: 
// create and populate a new Internal Planning meeting agenda based on a template document. It accomplishes this by creating a copy of the template, fetching links to relevant files (past and future) from specific folders, calculating the current Program Increment (PI), and replacing placeholders in the document with the obtained data, including Team Schedules. It has an external trigger to execute each Sunday

// To Note: 
// This script is developed as a Google Apps Script standalone script. It is designed to operate independently and does not require any external application or service to function. It is a self-contained piece of code with a time-based weekly trigger.

// To Use: 
// Copy the Script: Open the Google Apps Script editor by clicking on "Extensions" in the Google Docs/Sheets/Slides menu, then selecting "Apps Script." Delete any existing code in the editor and paste the entire script provided into the editor.
// Set Up Template: Replace the value of the `templateId` variable with the ID of your own Google Docs template for the Internal Planning Meeting agenda. To get the template ID, open your template, copy the document's URL, and extract the unique document ID from the URL.
// Configure Folder IDs: Replace the folder IDs (`previousAgendaFolderId`, `operaTagUpFolderId`, `snwgMonthlyFolderId`, and `dmprFolderId`) with the IDs of the respective folders where your past agendas, OPERA tag-up files, SNWG/NASA monthly reports, and DMPR files are stored. You can find the folder ID in the URL when you open the folder in Google Drive.
// Save and Deploy: Save the script and click on the "Deploy" button. Choose "New deployment" and configure the settings as per your requirement. For simplicity, choose "Web app" and set access to "Anyone, even anonymous." Click "Deploy" and grant the necessary permissions.
// Run the Script: After deploying, a URL will be generated for the script's web app. Open that URL in a browser, and it will run the script to create a new Internal Planning Meeting agenda based on the template. The agenda will include links to the most recent past and closest future files from the specified folders, along with the current Program Increment (PI) information.
// Automate Regularly: For convenience, you can set up a trigger to run the script automatically at a specific time or interval. In the script editor, click on the "Triggers" icon (the clock), and set up a new trigger to run the `createNewInternalAgenda` function at your desired frequency (e.g., weekly on Mondays).
// Customize Further (Optional): If you want to modify the appearance or content of the generated agenda, you can customize the template document to suit your needs. Just ensure that the placeholders (`{{...}}`) in the template match the names used in the script.
// Remember to keep your template document up to date and ensure that the relevant files are present in the specified folders to get accurate information in the generated agenda. Happy planning!

////////////////////////////////////////////////

// Helper Function: Create a copy of a template document
function createCopyOfTemplate(templateId) {
  return DriveApp.getFileById(templateId).makeCopy();
}

// Helper Function: Get the Monday following a given date
function getMondayFollowingDate(date) {
  var day = date.getDay();
  var diff = (day === 0 ? 1 : 8) - day;
  var nextMonday = new Date(date.getTime() + diff * 24 * 60 * 60 * 1000);
  var formattedDate = Utilities.formatDate(nextMonday, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return formattedDate;
}

// Helper Function: Replace a placeholder with a formatted date
function replaceWithFormattedDate(documentBody, placeholderText, currentDate) {
  var dateForInternal = new Date(currentDate);
  dateForInternal.setDate(dateForInternal.getDate() + 1);
  var formattedDate = Utilities.formatDate(dateForInternal, Session.getScriptTimeZone(), "EEEE, MMMM dd, yyyy");
  documentBody.replaceText(placeholderText, formattedDate);
}

// Helper Function: Get the link of the most recent or next file from a folder
function getMostRecentFileLink(folderId, excludeId, isFuture) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var relevantPastFile = null;
  var relevantFutureFile = null;
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileId = file.getId();
    
    if (!fileName.includes("Template")) { // Exclude files with "Template" in their name
      
      var fileDate = parseDateFromFileName(fileName); // Implement this function to parse the date from the filename
      
      if (fileDate instanceof Date) {
        var fileStr = Utilities.formatDate(fileDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        
        if (fileStr < todayStr) {
          // Check if the file is in the past and more recent than the current most recent past file
          if (!relevantPastFile || fileDate > parseDateFromFileName(relevantPastFile.getName())) {
            relevantPastFile = file;
          }
        } else if (fileStr > todayStr) {
          // Check if the file is in the future and closer to today than the current closest future file
          if (!relevantFutureFile || fileDate < parseDateFromFileName(relevantFutureFile.getName())) {
            relevantFutureFile = file;
          }
        }
      }
      
    }
  }
  
  return isFuture ? (relevantFutureFile ? relevantFutureFile.getUrl() : '') : (relevantPastFile ? relevantPastFile.getUrl() : '');
}

// Helper function: parse the date from the filename
function parseDateFromFileName(fileName) {
  // Regular expression to match the date format "YYYY-MM-DD" at the beginning of the filename
  var dateRegex = /^(\d{4}-\d{2}-\d{2})/;
  
  var match = fileName.match(dateRegex);
  if (match && match.length > 1) {
    var dateString = match[1];
    var dateComponents = dateString.split("-");
    var year = parseInt(dateComponents[0]);
    var month = parseInt(dateComponents[1]) - 1; // Months in JavaScript are 0-indexed (0 = January, 11 = December)
    var day = parseInt(dateComponents[2]);

    // Check if the parsed date components are valid
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      return new Date(year, month, day);
    }
  }

  // Return null if the date cannot be parsed from the filename
  return null;
}

// Helper Function: Get the relevant file based on the specified conditions
function getRelevantFile(isFuture, todayStr, fileName, fileId, excludeId, currentRelevantFile, currentRelevantFileName) {
  var fileIsRelevant = (isFuture && fileName.localeCompare(todayStr) > 0) || (!isFuture && fileName.localeCompare(todayStr) <= 0);
  var fileIsMoreRecent = currentRelevantFileName === '' || (isFuture ? fileName.localeCompare(currentRelevantFileName) < 0 : fileName.localeCompare(currentRelevantFileName) > 0);
  if (fileIsRelevant && fileIsMoreRecent && fileId !== excludeId && !fileName.includes("Template")) {
    return DriveApp.getFileById(fileId);
  }
  return currentRelevantFile;
}

// Helper function: Get Teeam Schedule events from the designated calendar
function getCalendarEvents() {
  var now = new Date();
  var sixWeeksLater = new Date(now.getTime() + 42 * 24 * 60 * 60 * 1000); // Add 42 days to the current date
  
  var calendarId = 'c_365230bc41700e58e23f74b286db1773d395e4bc6807c81a4c78658df5db423e@group.calendar.google.com'; // SNWG Team Schedules Google Calendar
  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(now, sixWeeksLater);
  
  var eventDetails = [];
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    
    var startDate = event.getStartTime();
    var endDate = event.getEndTime();

    // Adjust for all-day events
    if (event.isAllDayEvent()) {
      endDate = new Date(endDate.getTime() - 24*60*60*1000); // Subtract one day from the end date
    }

    var formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMMM d");
    var formattedEndDate = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MMMM d");
    
    Logger.log("Event: " + event.getTitle());
    Logger.log("Start Date: " + startDate + " Formatted: " + formattedStartDate);
    Logger.log("End Date: " + endDate + " Formatted: " + formattedEndDate);
    
    // Check if start date and end date are the same
    var dateRange = (formattedStartDate === formattedEndDate) ? formattedStartDate : (formattedStartDate + " - " + formattedEndDate);
    
    eventDetails.push({
      title: event.getTitle(),
      date: dateRange
    });
  }
  return eventDetails;
}

// Helper Function: Replace placeholder text in a document with a hyperlink
function replaceWithHyperlink(documentBody, placeholderText, url) {
  var foundElement = documentBody.findText(placeholderText);
  if (foundElement) {
    var startOffset = foundElement.getStartOffset();
    var endOffset = foundElement.getEndOffsetInclusive();
    var textElement = foundElement.getElement().asText();
    if (url !== '') {
      var fileId = url.split('/')[5];
      var file = DriveApp.getFileById(fileId);
      var fileName = file.getName();
    } else {
      var fileName = 'Not found';
    }
    textElement.deleteText(startOffset, endOffset);
    textElement.insertText(startOffset, fileName).setLinkUrl(startOffset, startOffset + fileName.length - 1, url);
  }
}

// Helper Function: Get the link of the DMPR file corresponding to a given month
function getDMPRLink(folderId, currentDate) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var dmprFile = null;
  var dmprFileMonth = currentDate.substring(0, 7);

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    if (fileName.startsWith(dmprFileMonth) && !fileName.includes("Template")) {
      dmprFile = file;
      break;
    }
  }
  return dmprFile ? dmprFile.getUrl() : '';
}

// Primary function to create and populate a new Internal Planning meeting agenda
function createNewInternalAgenda() {
  var templateId = "1tE6xNFeMLVpcGwWMB9GuYpGom5F4Bi_81dUsp_W3jDQ"; // Template: YYYY-MM-DD Internal SNWG Meeting Agenda
  var newDocument = createCopyOfTemplate(templateId);
  var newDocumentId = newDocument.getId();
  var document = DocumentApp.openById(newDocumentId);
  var documentBody = document.getBody();

  // Get the date for the next Monday following the current date
  var currentDate = getMondayFollowingDate(new Date());
  var newDocumentName = currentDate + " SNWG MO Internal Planning Meeting"; 
  document.setName(newDocumentName);

  // IDs of folders where specific files are stored
  var previousAgendaFolderId = "1WKYw4jnP6ejRkOLAIPoPvbEYClaLE4eR"; // Weekly Internal Planning>FY23 SNWG MO Google Drive Folder
  var operaTagUpFolderId = "1M3EMWLCxhkqcPKLmH7grDu2zhSEuOvmc"; // OPERA> FY23 SNWG MO Google Drive Folder 
  var snwgMonthlyFolderId = "1HPjhc2LADvS9j3W_K3riq4RQPBngfqGY"; // Monthly Project Status Update> FY23 SNWG MO Google Drive Folder
  var dmprFolderId = "1y2vjwf52HBJpTeSIg7sYPaLSSzGAnWZU"; // ST10 DMPR> FY23 IMPACT Google Drive Folder

  // Get links to specific files from their respective folders
  var nextOperaTagUpLink = getMostRecentFileLink(operaTagUpFolderId, newDocumentId, true);
  var previousAgendaLink = getMostRecentFileLink(previousAgendaFolderId, newDocumentId, false);
  var previousOperaTagUpLink = getMostRecentFileLink(operaTagUpFolderId, newDocumentId, false);
  var snwgMonthlyLink = getMostRecentFileLink(snwgMonthlyFolderId, newDocumentId, false);
  var dmprLink = getDMPRLink(dmprFolderId, currentDate);

  // Get Team Schedules and replace {{Team Schedules}} placeholder
  var sectionStart = documentBody.findText("{{Team Schedules}}");
    if (sectionStart) {
      Logger.log("Found placeholder {{Team Schedules}} in the document.");

      var events = getCalendarEvents();
      var eventsText = ""; // Text to replace the placeholder
      for (var i = 0; i < events.length; i++) {
        var event = events[i];
        eventsText += event.title + " - " + event.date + "\n";
      }

      documentBody.replaceText("{{Team Schedules}}", eventsText.trim());
    } else {
      Logger.log("Did not find placeholder {{Team Schedules}} in the document.");
    }
  // Get the current PI (Program Increment) for the document using the PI calculator library script
  var adjustedDate = new Date(); // Use a valid date object here or pass the required date
  var adjustedPI = adjustedPIcalculator.getPI(adjustedDate);

  // Call the function to replace the placeholder texts in the document
  adjustedPIcalculator.replacePlaceholderWithPI(document, adjustedPI); // <-- Use the 'document' object here, not 'targetDocument'

  // Replace placeholders in the document body with the obtained links and PI information
  replaceWithHyperlink(documentBody, "{{Link to Previous Agenda}}", previousAgendaLink);
  replaceWithHyperlink(documentBody, "{{link to last OPERA tag up}}", previousOperaTagUpLink);
  replaceWithHyperlink(documentBody, "{{link to next OPERA tag up}}", nextOperaTagUpLink);
  replaceWithHyperlink(documentBody, "{{link to last SNWG/NASA monthly}}", snwgMonthlyLink);
  replaceWithHyperlink(documentBody, "{{link to current DMPR}}", dmprLink);

  //documentBody.replaceText("{{Adjusted PI}}", adjustedPI); 

  replaceWithFormattedDate(documentBody, "{{Internal Date}}", currentDate);
}
