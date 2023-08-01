// Purpose: general agenda is created and stored in a top-level folder for the day's Google Calendar (Conference Room Reservation) events
// How to use: 
// In your Google Workspace: create a Google Docs template document.
// In the template document: create fields {{Meeting Name}} {{Date of Calendar Event}}	{{Time of Calendar Event}} {{Conference Room}} {{Conference Room}} placeholders
// In Google Apps Script (https://script.google.com/home): create a new script with the following information.
// in the script: replace General meeting agenda Template ID with chosen template ID
// in the script: replace Folder id for newly created document to be placed in with the chosen folder ID
// in the script: replace 'custom agenda' with your chosen exclusion phrase to keep new agenda from being made 
// in the script: select "Triggers" tab (alarm clock) and add a trigger
// in the IMPACT conference room calendar: create a future event with the title of the meeting and location. 
// in the script: either run function createDocFromFutureEvents or wait for your trigger to run. 

// IMPACT Conference Room Calendar ID: mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com
// General meeting agenda Template ID: 1nZN2y09ad2V7CqK0QIqrFXB_lZvv4XyXtefbOzdpdpE <--replace with your chosen template
// Folder id for newly created document to be placed in: 16oXc4f1Rlvwk0vsXa27Dvy-niH1E1IOr <--replace with your chosen folder

// event is created in IMPACT conference room calendar - calendar ID: mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com 
// Search calendar for events for the day after current day and later.
// Check "Created by" id (owner email address) against specified email address.
  // If no match, abort action
  // If match, create document from specified template and place in specified folder.
// Find date of calendar event. Format in YYYY-MM-DD; name document "<date of the calendar event> + title of the event"
// Find name of calendar event. Replace {{Meeting Name}} placeholder text in new document.
// Find date of calendar event. Replace {{Date of Calendar Event}} placeholder text in new document. Format the date as "Day, Month Date, Year"
// Find time of calendar event. Replace {{Time of Calendar Event}} placeholder text in new document. Format the time as "beginning HH:MM - ending HH:MM Timezone". Timezone = CT
// Find location of calendar event. Replace {{Conference Room}} placeholder text in new document.
// Find "created by" email address. Replace {{Meeting Owner Name}} placeholder text in new document.
  // Note that {{Meeting Owner Name}} is replaced with the email address of the event creator and not the owner name, as Google Apps Script does not have a direct method to get the creator's name.
// trigger: set in script settings

/////////////////////////////////////////////

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
  
  body.replaceText('{{Meeting Owner Name}}', event.getCreators()[0]); //Replace with the creator's email address

  doc.saveAndClose();
}

// Helper function: Check if an event has a custom agenda template or otherwise should not be made; remove this function if not if not exculding routine events with custom agendas
function isEventRecurring(event) {
  return event.getDescription().toLowerCase().includes('custom agenda'); // <--replace with your chosen exclusion phrase to keep new agenda from being made 
}

// Primary function: Find future events and create a general agenda document from them
function createDocFromFutureEvents() {
  var calendarId = 'mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com'; 
  var templateId = '1nZN2y09ad2V7CqK0QIqrFXB_lZvv4XyXtefbOzdpdpE'; // <-- replace with desired template
  var targetFolderId = '16oXc4f1Rlvwk0vsXa27Dvy-niH1E1IOr'; // <-- replace with desired target folder

  // Add the specific owner's email address here
  var specificOwnerEmail = 'cherrelle.j.tucker@nasa.gov'; // <-- replace with desired meeting creator's email

  var events = getFutureEvents(calendarId);
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var processedEvents = scriptProperties.getProperty('processedEvents');
  processedEvents = processedEvents ? JSON.parse(processedEvents) : {};

  var targetFolder = DriveApp.getFolderById(targetFolderId);

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    var eventId = event.getId();

    if (processedEvents[eventId] || isEventRecurring(event)) { // <<-- Remove "|| isEventRecurring(event)" if not exculding routine events with custom agendas
      continue;
    }

    var eventCreatorEmail = event.getCreators()[0]; // Get the creator's email address
    
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
