// Purpose: general agenda is created and stored in a top-level folder for the day's Google Calendar (Conference Room Reservation) events
// Calendar ID: mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com
// Template ID: 1nZN2y09ad2V7CqK0QIqrFXB_lZvv4XyXtefbOzdpdpE
// Folder id for newly created document to be placed in: 16oXc4f1Rlvwk0vsXa27Dvy-niH1E1IOr. 

// Done: Make a copy of the template document
// Done: Set the new document's name with the date of the calendar event + title of the event; format YYYY-MM-DD
// Done: Replace Placeholder text: {{Meeting Name}} with the same name as the meeting on the calendar; 
// Done: Replace placeholder text: {{Date of Calendar Event}} with the same date as the calendar event; Format the date as "Day, Month Date, Year"  
// Done: Replace placeholder text: {{Time of Calendar Event}} with the same time range as the calendar event; Format the time as "beginning HH:MM - ending HH:MM Timezone"
// Done: Replace placeholder text {{Conference Room}} with the Location of the created meeting.
// Done: Replace the place holder text {{Meeting Owner Name}} with the email address of the person who created the meeting. 
//      Note that {{Meeting Owner Name}} is replaced with the email address of the event creator, as Google Apps Script does not have a direct method to get the creator's name.
// Done: triggers creation of documents for events that have a specific owner email address. The first time you run createTimeDrivenTrigger function, it will create a time-driven trigger to execute the createDocFromFutureEvents function every 5 minutes. 
// Done: Added custom menu: 
// Future Development: Exlude routine meetings? ie: Internal and Monthly have custom agendas and scripts. 
// Future Development: Eventually would like thtis to automatically sort to the appropriate project folder. 

function onOpen(){
  DocumentApp.getUi()
  .createMenu('Action Items')
  .addItem('Get Actions','createDocFromFutureEvents')
  .addToUi();
}

function createDocFromFutureEvents() {
  var calendarId = 'mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com'; //Test calendar; replace with mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com when ready to deploy
  var templateId = '1nZN2y09ad2V7CqK0QIqrFXB_lZvv4XyXtefbOzdpdpE';
  var targetFolderId = '16oXc4f1Rlvwk0vsXa27Dvy-niH1E1IOr';
  
  // Add the specific owner's email address here
  var specificOwnerEmail = 'cherrelle.j.tucker@nasa.gov';

  var calendar = CalendarApp.getCalendarById(calendarId);
  var today = new Date();
  var aMonthFromNow = new Date();
  aMonthFromNow.setDate(today.getDate() + 30);

  var events = calendar.getEvents(today, aMonthFromNow);
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var processedEvents = scriptProperties.getProperty('processedEvents');
  processedEvents = processedEvents ? JSON.parse(processedEvents) : {};

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    var eventId = event.getId();

    if (processedEvents[eventId]) {
      continue;
    }

    var eventCreatorEmail = event.getCreators()[0]; //Get the creator's email address
    
    // Check if the event's creator is the specific owner
    if (eventCreatorEmail !== specificOwnerEmail) {
      continue;
    }

    var eventTitle = event.getTitle();
    var eventDate = event.getStartTime();
    var eventEndDate = event.getEndTime();

    var formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var docTitle = formattedDate + ' ' + eventTitle;

    var targetFolder = DriveApp.getFolderById(targetFolderId);
    var file = DriveApp.getFileById(templateId).makeCopy(docTitle, targetFolder);

    var doc = DocumentApp.openById(file.getId());
    var body = doc.getBody();

    body.replaceText('{{Meeting Name}}', eventTitle);

    var longFormattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'EEEE, MMMM dd, yyyy');
    body.replaceText('{{Date of Calendar Event}}', longFormattedDate);

    var eventTime = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'HH:mm') + 
                    ' - ' + Utilities.formatDate(eventEndDate, Session.getScriptTimeZone(), 'HH:mm z');
    body.replaceText('{{Time of Calendar Event}}', eventTime);

    var eventLocation = event.getLocation() || 'No location specified'; //Get the location or set a default message
    body.replaceText('{{Conference Room}}', eventLocation);

    body.replaceText('{{Meeting Owner Name}}', eventCreatorEmail); //Replace with the creator's email address

    doc.saveAndClose();
    
    processedEvents[eventId] = true;
  }

  scriptProperties.setProperty('processedEvents', JSON.stringify(processedEvents));
}

  

function createDocFromFutureEvents() {
  var calendarId = 'mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com'; //Test calendar; replace with mn9msmqj2nqobs0md4gmgfnabk@group.calendar.google.com when ready to deploy
  var templateId = '1nZN2y09ad2V7CqK0QIqrFXB_lZvv4XyXtefbOzdpdpE';
  var targetFolderId = '16oXc4f1Rlvwk0vsXa27Dvy-niH1E1IOr';
  
  // Add the specific owner's email address here
  var specificOwnerEmail = 'cherrelle.j.tucker@nasa.gov';

  var calendar = CalendarApp.getCalendarById(calendarId);
  var today = new Date();
  var aMonthFromNow = new Date();
  aMonthFromNow.setDate(today.getDate() + 30);

  var events = calendar.getEvents(today, aMonthFromNow);
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var processedEvents = scriptProperties.getProperty('processedEvents');
  processedEvents = processedEvents ? JSON.parse(processedEvents) : {};

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    var eventId = event.getId();

    if (processedEvents[eventId]) {
      continue;
    }

    var eventCreatorEmail = event.getCreators()[0]; //Get the creator's email address
    
    // Check if the event's creator is the specific owner
    if (eventCreatorEmail !== specificOwnerEmail) {
      continue;
    }

    var eventTitle = event.getTitle();
    var eventDate = event.getStartTime();
    var eventEndDate = event.getEndTime();

    var formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var docTitle = formattedDate + ' ' + eventTitle;

    var targetFolder = DriveApp.getFolderById(targetFolderId);
    var file = DriveApp.getFileById(templateId).makeCopy(docTitle, targetFolder);

    var doc = DocumentApp.openById(file.getId());
    var body = doc.getBody();

    body.replaceText('{{Meeting Name}}', eventTitle);

    var longFormattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'EEEE, MMMM dd, yyyy');
    body.replaceText('{{Date of Calendar Event}}', longFormattedDate);

    var eventTime = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), 'HH:mm') + 
                    ' - ' + Utilities.formatDate(eventEndDate, Session.getScriptTimeZone(), 'HH:mm z');
    body.replaceText('{{Time of Calendar Event}}', eventTime);

    var eventLocation = event.getLocation() || 'No location specified'; //Get the location or set a default message
    body.replaceText('{{Conference Room}}', eventLocation);

    body.replaceText('{{Meeting Owner Name}}', eventCreatorEmail); //Replace with the creator's email address

    doc.saveAndClose();
    
    processedEvents[eventId] = true;
  }

  scriptProperties.setProperty('processedEvents', JSON.stringify(processedEvents));
}

  