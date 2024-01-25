// Purpose: 
// automatically create new agenda documents for specific events in a Google Calendar. It scans the calendar for future events within the next 30 days and checks if the event's creator matches a specific owner's email address. For eligible events, it generates a new agenda document by making a copy of a provided template, populating it with event details like meeting name, date, time, and location. The script also excludes events with a "custom agenda" description and tracks processed events to avoid duplicating agendas for the same event.

// Future Development: Automatically place agenda in the appropriate project folder. 

// To Note: 
// This script is developed as a Google Apps Script standalone script. It is designed to operate independently and does not require any external application or service to function. It is a self-contained piece of code with a time-based daily trigger.

// To Use: 
// 1. Prepare Google Calendar and Template: Ensure you have the required access and permissions to create events in Google Calendar. Also, prepare a template for the agenda in Google Docs and note down its Template ID.
// 2. Open the Script Editor: Access the Google Apps Script editor by clicking on "Extensions" in the top menu and choosing "Apps Script.
// 3. Copy and Paste the Script: Copy the entire script provided above and paste it into the Google Apps Script editor
// 4. Configure Global Variables: Adjust the global variables at the beginning of the script to match your Template ID, the calendar ID for the "IMPACT conference room calendar," and the folder ID where you want to store the new agenda documents
// 5. Save the Script: Save the script by clicking on the floppy disk icon or pressing "Ctrl + S" (Windows) or "Cmd + S" (Mac)
// 6. Run the Script: To generate new agenda documents, click on the play button ▶️ in the Google Apps Script editor or manually run the `createDocFromFutureEvents()` function.
// 7. Agenda Generation: The script will find future events within the next 30 days in the specified calendar and create new agenda documents based on the provided template. Each document will be named with the event date and title.
// 8. Customize Agenda Exclusion (Optional): If you want to exclude certain events from having agendas generated, modify the `isEventRecurring()` function as indicated in the comments.
// 9. Hyperlink to Previous Agenda: The script will automatically search for the most recent previous month's agenda in the folder (excluding the newly created one) and replace the placeholder "{{link to last SNWG/NASA monthly}}" with a hyperlink to that agenda.
// 10. Schedule Automation (Optional): If you want to generate agendas automatically, set up a trigger to run the script at a specific interval (e.g., daily). To do this, click on the clock icon in the Google Apps Script editor, add a new trigger, and set it to run the `createDocFromFutureEvents()` function on the desired schedule.
// 11. Ensure Permission: When running the script for the first time, it may ask for permission to access your Google Calendar and Drive. Click "Continue" and grant the necessary permissions.
// Before executing the script, carefully review its functionality, and customize it according to your requirements. Also, ensure that the script has access to the appropriate folders and files in your Google Drive and Calendar. 

///////////////////////////////////////////////

// Global variables
var calendarId = 'mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com'; // IMPACT conference room calendar
var templateId = '1nZN2y09ad2V7CqK0QIqrFXB_lZvv4XyXtefbOzdpdpE'; // SNWG MO Template: YYYY-MM-DD Meeting Agenda
var targetFolderId = '16oXc4f1Rlvwk0vsXa27Dvy-niH1E1IOr'; // Meeting Notes> Auto generated New Agendas SNWG MO Google drive folder
var specificOwnerEmail = 'cherrelle.j.tucker@nasa.gov'; // specific owner's email address

// Helper function: Check if the specific owner is an attendee of the event
function isOwnerAnAttendee(event, ownerEmail) {
  var attendees = event.getGuestList();
  return attendees.some(function(attendee) {
    return attendee.getEmail() === ownerEmail;
  });
}

// Helper function: Get future events from a Google Calendar
function getFutureEvents(calendarId) {
  var today = new Date();
  var aMonthFromNow = new Date();
  aMonthFromNow.setDate(today.getDate() + 30);
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(today, aMonthFromNow);

  return events;
}

// Helper function: Create a new document by making a copy of a template document
function createDocumentFromTemplate(templateId, folder, documentName) {
  var template = DriveApp.getFileById(templateId);
  var newFile = template.makeCopy(documentName, folder);

  return DocumentApp.openById(newFile.getId());
}

// Helper function: Replace placeholders in the document with event details
function replacePlaceholdersInDoc(doc, event) {
  var eventTitle = event.getTitle();
  var eventDate = event.getStartTime();
  var eventEndDate = event.getEndTime();
  
  var body = doc.getBody();
  body.replaceText('{{Meeting Name}}', eventTitle);
  
  var longFormattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'EEEE, MMMM dd, yyyy');
  body.replaceText('{{Date of Calendar Event}}', longFormattedDate);
  
  var eventTime = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'HH:mm') + 
                    ' - ' + Utilities.formatDate(eventEndDate, Session.getScriptTimeZone(), 'HH:mm z');
  body.replaceText('{{Time of Calendar Event}}', eventTime);
  
  var eventLocation = event.getLocation() || 'No location specified'; //Get the location or set a default message
  body.replaceText('{{Conference Room}}', eventLocation);
  
  body.replaceText('{{Meeting Owner Name}}', event.getCreators()[0]); // {{Meeting Owner Name}} is replaced with the email address of the event creator and not the owner name, as Google Apps Script does not have a direct method to get the creator's name.

  doc.saveAndClose();
}

// Helper function: Check if an event has a custom agenda template or otherwise should not be made; remove this function if not if not exculding routine events with custom agendas
function isEventRecurring(event) {
  return event.getDescription().toLowerCase().includes('custom agenda');
}

// Primary function: Find future events and create a general agenda document from them
function createDocFromFutureEvents() {
  var events = getFutureEvents(calendarId);
  var scriptProperties = PropertiesService.getScriptProperties();
  var processedEvents = scriptProperties.getProperty('processedEvents');

  processedEvents = processedEvents ? JSON.parse(processedEvents) : {};
  var targetFolder = DriveApp.getFolderById(targetFolderId);

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    var eventId = event.getId();

    if (processedEvents[eventId] || isEventRecurring(event)) {
      continue;
    }

    if (!isOwnerAnAttendee(event, specificOwnerEmail)) {
      continue;
    }

    var eventCreatorEmail = event.getCreators()[0]; //Get the creator's email address
    
    // Check if the event's creator is the specific owner
    if (eventCreatorEmail !== specificOwnerEmail) {
      continue;
    }

    var eventTitle = event.getTitle();
    var eventDate = event.getStartTime();
    var formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var docTitle = formattedDate + ' ' + eventTitle;

    var doc = createDocumentFromTemplate(templateId, targetFolder, docTitle);
    replacePlaceholdersInDoc(doc, event);
    
    processedEvents[eventId] = true;
  }

  scriptProperties.setProperty('processedEvents', JSON.stringify(processedEvents));
}
