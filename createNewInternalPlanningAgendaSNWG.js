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

// Dev Seed folder:https://drive.google.com/drive/folders/1Bvj1b1u2LGwjW5fStQaLRpZX5pmtjHSj; {{link to Assessment DevSeed Agenda}}
// Assessment folder: https://drive.google.com/drive/folders/1dmN0oYQZwGFu83BwOGT90I_GFtGH1aup; {{link to Assessment HQ Agenda}}
// SEP folder: https://drive.google.com/drive/folders/1Cw_sdH_IleGbtW1mVoWnJ0yqoyzr4Oe0; {{link to SEP Agenda}}
// OPERA folder: https://drive.google.com/drive/folders/1AX95NPrIYiLvn_1l8a6G4JwI6wW0viD8; {{link to last OPERA tag up}} {{link to next OPERA tag up}} 
// DMPR folder: https://drive.google.com/drive/folders/1NCH6-V9pMA8pOivX0XD5ZtGLT-8OQZ6A; {{link to current DMPR}}
// NASA SNWG folder: https://drive.google.com/drive/folders/1r52FELtJWytcp5Iw7F01wSxuXeRYcm78; {{link to last SNWG/NASA monthly}} 
// SNWG Internal folder: https://drive.google.com/drive/folders/1SRIUs7CUEdGUw0r1PI52e0OJpfXYN0z8; {{Link to Previous Agenda}} 
// Team Schedules calendar: c_365230bc41700e58e23f74b286db1773d395e4bc6807c81a4c78658df5db423e@group.calendar.google.com; {{Team Schedules}}
// IMPACT PI calendar = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com'; // IMPACT Project Increment Google Calendar

var impactPIcalendar = 'xxxxxxxxxxxx@group.calendar.google.com'; // IMPACT Project Increment Google Calendar

// helper function to get current week Mon-Fri dates
function getCurrentWeekDates() {
  try {
    var today = new Date();
    var startOfWeek = new Date(today.setDate(today.getDate() - today.getDay() + 1));
    var endOfWeek = new Date(today.setDate(today.getDate() - today.getDay() + 5));
    return {
      start: startOfWeek,
      end: endOfWeek,
      formatted: Utilities.formatDate(startOfWeek, 'GMT', 'MM/dd/yy') + " - " + Utilities.formatDate(endOfWeek, 'GMT', 'MM/dd/yy')
    };
  } catch (error) {
    Logger.log("Error getting current week dates: " + error);
    throw error;
  }
}

// Helper function to find and return the Current Sprint event.
function getCurrentSprintEvent() {
  var now = new Date();
  var calendar = CalendarApp.getCalendarById(impactPIcalendar);
  var events = calendar.getEventsForDay(now, { search: 'Sprint' });
  Logger.log('Number of events found: ' + events.length); // Log the number of events found

  if (events.length > 0) {
    var event = events[0];
    Logger.log('Selected event title: ' + event.getTitle()); // Log the title of the selected event
    Logger.log('Event start time: ' + event.getStartTime()); // Log start time
    Logger.log('Event end time: ' + event.getEndTime()); // Log end time
    return event;
  } else {
    Logger.log('No events found for the current day.');
    return null;
  }
}

// Function to get FY.PI.Sprint from IMPACT PI Calendar with Week Number
function getPiFromImpactPiCalendar(internalDate) {
  var calendar = CalendarApp.getCalendarById(impactPIcalendar);
  var weekStart = new Date(internalDate.getTime());
  weekStart.setDate(internalDate.getDate() - internalDate.getDay() + 1); // Set to Monday of the week
  var weekEnd = new Date(weekStart.getTime());
  weekEnd.setDate(weekStart.getDate() + 4); // Set to Friday of the week

  var events = calendar.getEvents(weekStart, weekEnd);
  var piRegex = /PI \d{2}\.\d Sprint \d/; // Regex to match "PI YY.Q Sprint S" format

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    if (piRegex.test(event.getTitle())) {
      var eventTitle = event.getTitle();
      var eventStartDate = event.getStartTime();
      var weekNumber = determineWeekNumber(eventStartDate, internalDate);
      return eventTitle + " Week " + weekNumber;
    }
  }
  return "No PI Event Found";
}

// Adjusted function to determine the week number based on the internal date
function determineWeekNumber(eventStartDate, internalDate) {
  var oneWeekDuration = 7 * 24 * 60 * 60 * 1000; // One week in milliseconds
  var timeDifference = internalDate.getTime() - eventStartDate.getTime();

  if (timeDifference < oneWeekDuration) {
    return 1; // First week of the sprint
  } else {
    return 2; // Second week of the sprint
  }
}


// Modified helper function to populate {{Adjusted PI}} with the current PI information.
function populateAdjustedPI(document, internalDate) {
  var piInfo = getPiFromImpactPiCalendar(internalDate);

  // Replace the placeholder text in the document with the obtained PI information
  var documentBody = document.getBody();
  documentBody.replaceText('{{Adjusted PI}}', piInfo);
}

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

// Helper Function: Get the link of the Assessment HQ file within the same week
function getAssessmentHQLink(folderId, currentDate) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var assessmentHQFile = null;

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    // Extract the date part from the filename (assuming it's in YYYY-MM-DD format)
    var fileDateStr = fileName.match(/\d{4}-\d{2}-\d{2}/);

    if (fileDateStr && !fileName.includes("Template")) {
      // Check if the file date falls within the same week as the current date
      var fileDate = new Date(fileDateStr[0]);
      var weekStartDate = new Date(currentDate);
      var weekEndDate = new Date(currentDate);
      weekEndDate.setDate(weekEndDate.getDate() + 6); // Add 6 days to get the end of the week

      if (fileDate >= weekStartDate && fileDate <= weekEndDate) {
        Logger.log("Found Assessment HQ file for Date: " + currentDate);
        assessmentHQFile = file;
        break;
      }
    }
  }

  if (assessmentHQFile) {
    Logger.log("Assessment HQ File Found for Date: " + currentDate);
    return assessmentHQFile.getUrl();
  } else {
    Logger.log("Assessment HQ File Not Found for Date: " + currentDate);
    return '';
  }
}

// Helper Function: Get the link of the Assessment DevSeed file within the same week
function getAssessmentDevSeedLink(folderId, currentDate) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var assessmentDevSeedFile = null;

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    // Extract the date part from the filename (assuming it's in YYYY-MM-DD format)
    var fileDateStr = fileName.match(/\d{4}-\d{2}-\d{2}/);

    if (fileDateStr && !fileName.includes("Template")) {
      // Check if the file date falls within the same week as the current date
      var fileDate = new Date(fileDateStr[0]);
      var weekStartDate = new Date(currentDate);
      var weekEndDate = new Date(currentDate);
      weekEndDate.setDate(weekEndDate.getDate() + 6); // Add 6 days to get the end of the week

      if (fileDate >= weekStartDate && fileDate <= weekEndDate) {
        Logger.log("Found Assessment DevSeed file for Date: " + currentDate);
        assessmentDevSeedFile = file;
        break;
      }
    }
  }

  if (assessmentDevSeedFile) {
    Logger.log("Assessment DevSeed File Found for Date: " + currentDate);
    return assessmentDevSeedFile.getUrl();
  } else {
    Logger.log("Assessment DevSeed File Not Found for Date: " + currentDate);
    return '';
  }
}

// Helper Function: Get the link of the SEP file within the same week
function getSEPLink(folderId, currentDate) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var sepFile = null;

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();

    // Extract the date part from the filename (assuming it's in YYYY-MM-DD format)
    var fileDateStr = fileName.match(/\d{4}-\d{2}-\d{2}/);

    if (fileDateStr && !fileName.includes("Template")) {
      // Check if the file date falls within the same week as the current date
      var fileDate = new Date(fileDateStr[0]);
      var weekStartDate = new Date(currentDate);
      var weekEndDate = new Date(currentDate);
      weekEndDate.setDate(weekEndDate.getDate() + 6); // Add 6 days to get the end of the week

      if (fileDate >= weekStartDate && fileDate <= weekEndDate) {
        Logger.log("Found SEP file for Date: " + currentDate);
        sepFile = file;
        break;
      }
    }
  }

  if (sepFile) {
    Logger.log("SEP File Found for Date: " + currentDate);
    return sepFile.getUrl();
  } else {
    Logger.log("SEP File Not Found for Date: " + currentDate);
    return '';
  }
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

// Helper function: Get Team Schedule events from the designated calendar
function getCalendarEvents() {
  var now = new Date();
  var fourWeeksLater = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000); // Add 30 days to the current date
  
  var calendarId = 'xxxxxxxxxxxxxxx@group.calendar.google.com'; // SNWG Team Schedules Google Calendar
  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(now, fourWeeksLater);
  
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

// Helper function to get today's date
function getToday() {
  return new Date();
}

// Helper function to calculate the next most recent Integrated Master Schedule(IMS) meeting date
function getNextIMSDate() {
  const startDate = new Date("January 16, 2024 11:30:00");
  const today = getToday();

  // Calculate the difference in days from the start date
  const daysDiff = Math.floor((today - startDate) / (24 * 60 * 60 * 1000));

  // Calculate the number of bi-weeks since the start date
  const biWeeksSinceStart = Math.floor(daysDiff / 14);

  // Calculate the next due date, ensuring it's on Tuesday
  const nextIMSDate = new Date(startDate);
  nextIMSDate.setDate(startDate.getDate() + (biWeeksSinceStart * 14)); 
 // nextIMSDate.setHours(11, 30, 0, 0); // Set the time to 11:30 AM

 // Check if the next IMSDate is before today, if so, add 14 days to get the next bi-weekly date
  if (nextIMSDate < today) {
    nextIMSDate.setDate(nextIMSDate.getDate() + 13);
  }

  // If the calculated date is not a Tuesday (2 corresponds to Tuesday), adjust it
  /*while (nextIMSDate.getDay() !== 2) { // 2 corresponds to Tuesday
    nextIMSDate.setDate(nextIMSDate.getDate()); // Subtract one day until it's Tuesday
  }*/

  return nextIMSDate;
}

// Helper function to format a date in the standard format
function formatDate(date) {
  const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  return date.toLocaleDateString('en-US', options);
}

// Primary function to create and populate a new Internal Planning meeting agenda
function createNewInternalAgenda() {
  var templateId = "xxxxxxxxxxxxxxxx"; // Template: YYYY-MM-DD Internal SNWG Meeting Agenda
  var newDocument = createCopyOfTemplate(templateId);
  var newDocumentId = newDocument.getId();
  var document = DocumentApp.openById(newDocumentId);
  var documentBody = document.getBody();

  // Get the date for the next Monday following the current date
  var currentDate = getMondayFollowingDate(new Date());
  var newDocumentName = currentDate + " SNWG MO Internal Planning Meeting"; 
  document.setName(newDocumentName);

  // IDs of folders where specific files are stored
  var previousAgendaFolderId = "xxxxxxxxxxxx"; // Weekly Internal Planning>FY24 SNWG MO Google Drive Folder
  var operaTagUpFolderId = "xxxxxxxxxxxxxxx"; // OPERA> FY24 SNWG MO Google Drive Folder 
  var snwgMonthlyFolderId = "xxxxxxxxxxxx"; // Monthly Project Status Update> FY24 SNWG MO Google Drive Folder
  var dmprFolderId = "xxxxxxxxxxxxxxxxx"; // ST10 DMPR> FY24 IMPACT Google Drive Folder
  var assessmentDevSeedFolderID = "xxxxxxxxxxxxxxxx" // DevSeed FY24 Google Drive Folder
  var assessmentHQFolderID = "xxxxxxxxxxxxxxxx"; // AssessmentHQ Weekly CY24 Google Drive Folder
  var sepFolderID = "xxxxxxxxxxxxxxxxxxx"; // Stakeholder Engagement Program FY24 Google Drive folder

  // Get links to specific files from their respective folders
  var previousAgendaLink = getMostRecentFileLink(previousAgendaFolderId, newDocumentId, false);
  var snwgMonthlyLink = getMostRecentFileLink(snwgMonthlyFolderId, newDocumentId, false);
  var sepLink = getSEPLink(sepFolderID, currentDate);
  var assessmentHQLink = getAssessmentHQLink(assessmentHQFolderID, currentDate);
  var assessmentDevSeedLink = getAssessmentDevSeedLink(assessmentDevSeedFolderID, currentDate);
  var nextOperaTagUpLink = getMostRecentFileLink(operaTagUpFolderId, newDocumentId, true);
  var previousOperaTagUpLink = getMostRecentFileLink(operaTagUpFolderId, newDocumentId, false);
  var dmprLink = getDMPRLink(dmprFolderId, currentDate);

    // Get the next IMS meeting date using getNextIMSDate() function
  var nextIMSDate = getNextIMSDate();

  // Format the next IMS meeting date using the formatDate() function
  var formattedNextIMSDate = formatDate(nextIMSDate);

  // Replace the placeholder text {{Next IMS Meeting Date}} with the formatted date
  replaceWithFormattedDate(documentBody, "{{Next IMS Meeting Date}}", formattedNextIMSDate);

  // Replace the placeholder text {{Adjusted PI}}
  // Convert string to Date object
  var internalDate = new Date(currentDate);
  populateAdjustedPI(document, internalDate);

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
    

  // Replace placeholders in the document body with the obtained links and PI information
  replaceWithHyperlink(documentBody, "{{Link to Previous Agenda}}", previousAgendaLink);
  replaceWithHyperlink(documentBody, "{{link to last OPERA tag up}}", previousOperaTagUpLink);
  replaceWithHyperlink(documentBody, "{{link to next OPERA tag up}}", nextOperaTagUpLink);
  replaceWithHyperlink(documentBody, "{{link to last SNWG/NASA monthly}}", snwgMonthlyLink);
  replaceWithHyperlink(documentBody, "{{link to current DMPR}}", dmprLink);
  replaceWithHyperlink(documentBody, "{{link to Assessment DevSeed Agenda}}", assessmentDevSeedLink);
  replaceWithHyperlink(documentBody, "{{link to SEP Agenda}}", sepLink);
  replaceWithHyperlink(documentBody, "{{link to Assessment HQ Agenda}}", assessmentHQLink);

  replaceWithFormattedDate(documentBody, "{{Internal Date}}", currentDate);
}
