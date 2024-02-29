/*
Script Name: colorEventsAndSendNotificationEmailsIMPACT

Description: 
This script is designed to work with the IMPACT conference room Google Calendar. It recolors events on the calendar based on their locations, using predefined color codes associated with specific room identifiers. Additionally, the script identifies overlapping events (double bookings) at the same location and sends email notifications to the event organizers.

Prerequisites: 
1. Access to Google Apps Script.
2. A Google Calendar with events to manage.
3. Permissions to modify calendar events and send emails through Gmail.
 - Permission Notes:
 - See only free/busy (hide details) - inopperable for the purposes of this script
 - See all event details - can see event details, but cannot see event colors. Cannot change event colors. Minimum required for email functionality.
 - Make changes to events - can see colors and change event colors. Minimum persmission required for full functionality.  

Setup:
1. Open the Google Apps Script editor via "Tools" > "Script editor" in Google Sheets.
2. Replace 'your_calendar_id_here' in the script with the ID of the target Google Calendar.
3. Adjust the 'colorMap' object in the script to match your room identifiers and color preferences.
4. Save the script.
5. Set up a time-driven trigger for the 'colorEvents' function to run automatically (e.g., daily at a specific time).

Execution: 
The script can be executed manually or automatically. 
- For manual execution, select the 'colorEvents' function and click the play button in the Apps Script editor.
- For automatic execution, the script will run at the specified intervals set by the time-driven trigger.

Script Functions: 
1. colorEvents: Main function to recolor events and check for overlaps.
2. canEditEvents: Checks if the script has edit permissions for the calendar.
3. assignColorToLocation: Assigns colors to events based on location.
4. isVirtualEvent: Determines if an event is virtual (without a room identifier).
5. checkOverlapsForSameColorEvents: Identifies overlapping events and prepares email info.
6. sendOverlapEmails: Sends notification emails to organizers of overlapping events.
7. onCalendarUpdate (Optional): Trigger function to run 'colorEvents' when the calendar is updated.

Outputs: 
- Events in the calendar are recolored based on their locations.
- Notification emails are sent to organizers of overlapping events.

Post-Execution: 
- Check the Google Calendar to verify that events have been correctly recolored.
- Confirm that email notifications have been sent for any overlapping events.

Troubleshooting:
- Ensure that the correct Calendar ID is used.
- Verify that the script has the necessary permissions to modify calendar events and send emails.
- Check execution logs in the Apps Script editor for errors or issues.

Notes:
- This script requires that the user has sufficient permissions on the Google Calendar and Gmail.
- Be aware of Google Apps Script's quotas and limitations, especially for calendar operations and email sending.
- The effectiveness of the script depends on the accuracy of the 'colorMap' configuration.
*/

/////////////////////////////////////////////////////////////////////

// Global constant for Calendar ID
var CALENDAR_ID = 'xxxxxxxx@group.calendar.google.com'; // IMPACT Conference Room Schedule shared Google Calendar 

// Main function to recolor events, check for overlaps, check for missing locations, and send conflict found emails for the next week.
function colorEvents() {
  var today = new Date();
  var nextweek = new Date();
  nextweek.setDate(today.getDate() + 7); 

  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    Logger.log("No calendar found with the ID: " + CALENDAR_ID);
    return;
  }
  var events = calendar.getEvents(today, nextweek);

  // Google Color assignments
  var colorMap = {
    // 1: Lavendar; #7986cb
    // 2: Sage; #33b679
    '3': ["mccarthy", "1063"], // Grape; #8e24aa
    '4': ["teams"], // Flamingo; #e67c73
    '5': ["1062", "lovelace"], // Banana; #f6bf26
    '6': ["hamilton", "1063a"], // Tangerine; #f4511e
    '7': ["3098", "3rd", "CR3098", "CR3098:"], // Peacock; #039be5
    // 8: Graphite; #616161
    '9': ["3084", "CR3084"], // Blueberry; #3f51b5
    '10': ["1030", "turing"], // Basil; #0b8043
    '11': ["1034"], // Tomato; #d500000
    // 12: Default Calendar Color
  };

  var locationEventsMap = {};
    var eventsWithoutLocation = [];
    var canEditCalendar = canEditEvents(calendar);

    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      var location = event.getLocation().toLowerCase().replace('#', '');
      var locationWords = location.trim().split(/\s+/);
      var color = assignColorToLocation(locationWords, colorMap);

      // Try to set color if the calendar can be edited
      if (canEditCalendar && color !== '') {
        try {
          event.setColor(color);
          Logger.log("Event: '" + event.getTitle() + "' set to color '" + color + "'");
        } catch (e) {
          Logger.log("Error setting color for event: " + e.message);
        }
      }

      if (color === '') {
        eventsWithoutLocation.push(event);
      }

      if (!locationEventsMap[color]) {
        locationEventsMap[color] = [];
      }
      locationEventsMap[color].push(event);
    }

    // Process overlaps for same color events
    for (var color in locationEventsMap) {
      var sameColorEvents = locationEventsMap[color];
      checkOverlapsForSameColorEvents(sameColorEvents, colorMap);
    }

    // Notify for events without location
    if (eventsWithoutLocation.length > 0) {
      notifyLocationlessEvents(eventsWithoutLocation);
  }
}

//function to check if the account running the script has calendar edit permissions. 
function canEditEvents(calendar) {
  try {
    var testEvent = calendar.createEvent("Test Event", new Date(), new Date());
    testEvent.deleteEvent();
    return true;
  } catch(e) {
    Logger.log("No edit permissions for the calendar: " + e.message);
    return false;
  }
}

// funciton to assign Google You color code to an event based on its location
function assignColorToLocation(locationWords, colorMap) {
  for (var color in colorMap) {
    var phrases = colorMap[color];
    for (var k = 0; k < phrases.length; k++) {
      var phrase = phrases[k];
      if (locationWords.includes(phrase)) {
        return color;
      }
    }
  }
  return ''; // Return empty string if no color matches
}

// function to determine if an event is virtual based on the absence of room identification; 
function isVirtualEvent(event, colorMap) {
  var locationWords = event.getLocation().toLowerCase().trim().split(/\s+/);
  for (var color in colorMap) {
    var phrases = colorMap[color];
    for (var k = 0; k < phrases.length; k++) {
      var phrase = phrases[k];
      if (locationWords.includes(phrase)) {
        return false; // Event has a room identifier, not a virtual event
      }
    }
  }
  return true; // Event does not have any room identifiers, it is a virtual event
}

// function to notify organizers of events that do not have a location specified to add location or remove from calendar.
function notifyLocationlessEvents(events) {
  events.forEach(function(event) {
    var subject = "Missing Location Notification for Event";
    var message = "Dear Organizer,\n\n" +
                  "Your event titled '" + event.getTitle() + "' scheduled for " + event.getStartTime() + 
                  " does not have a location specified.\n\n" +
                  "Please update the event with the appropriate location information. If your event is virtual, consider placing it on your individual work calendar. \n\n" +
                  "Best regards,\n" +
                  "IMPACT Conference Room Schedule";

    var organizerEmail = event.getCreators()[0];
    Logger.log("Sending locationless event notification to: " + organizerEmail);
    MailApp.sendEmail(organizerEmail, subject, message);
  });
}

// function to check for overlaps within the same colored events and prepares information for notifications.
function checkOverlapsForSameColorEvents(sameColorEvents, colorMap) {
  var overlappingEventsInfo = [];

  sameColorEvents.sort(function(a, b) {
    return a.getStartTime() - b.getStartTime();
  });

  for (var j = 0; j < sameColorEvents.length - 1; j++) {
    var currentEvent = sameColorEvents[j];
    var nextEvent = sameColorEvents[j + 1];

    if (!isVirtualEvent(currentEvent, colorMap) && !isVirtualEvent(nextEvent, colorMap) && currentEvent.getEndTime() > nextEvent.getStartTime()) {
      try {
        currentEvent.setColor('8'); // Graphite color for overlap
        nextEvent.setColor('8');    // Graphite color for overlap
      } catch(e) {
        Logger.log("Error setting color for overlapping events: " + e.message);
      }

      overlappingEventsInfo.push({
        organizerEmail1: currentEvent.getCreators()[0],
        organizerEmail2: nextEvent.getCreators()[0],
        title1: currentEvent.getTitle(),
        title2: nextEvent.getTitle()
      });

      Logger.log("Overlap detected between the events: '" + currentEvent.getTitle() + "' and '" + nextEvent.getTitle() + "'");
    }
  }

  if (overlappingEventsInfo.length > 0) {
    sendOverlapEmails(overlappingEventsInfo);
  }
}

// function to send email notifications to organizers of overlapping events. 
function sendOverlapEmails(overlappingEventsInfo) {
  overlappingEventsInfo.forEach(function(info) {
    var subject = "Room Booking Overlap Notification";
    var message = "Dear Organizers,\n\n" +
                  "An overlap has been detected for the following events:\n" +
                  "Event 1: " + info.title1 + " (Organizer: " + info.organizerEmail1 + ")\n" +
                  "Event 2: " + info.title2 + " (Organizer: " + info.organizerEmail2 + ")\n\n" +
                  "Please coordinate to resolve this scheduling conflict. Conflict organizers will be reminded each day until the conflict is resolved.\n\n" +
                  "Best regards,\n" +
                  "IMPACT Conference Room Schedule \n\n\n" +
                  "This inbox is not monitored.";

    // Send email to the first organizer
    if (info.organizerEmail1) {
      Logger.log("Sending email to: " + info.organizerEmail1);
      MailApp.sendEmail(info.organizerEmail1, subject, message);
    }

    // Send email to the second organizer if different
    if (info.organizerEmail1 !== info.organizerEmail2 && info.organizerEmail2) {
      Logger.log("Sending email to: " + info.organizerEmail2);
      MailApp.sendEmail(info.organizerEmail2, subject, message);
    }

    // Log email sending information
    Logger.log("Sending email to: " + info.organizerEmail1);
    MailApp.sendEmail(info.organizerEmail1, subject, message);
    Logger.log("Email sent from " + Session.getActiveUser().getEmail() + " to " + info.organizerEmail1);

    if (info.organizerEmail1 !== info.organizerEmail2) {
      Logger.log("Sending email to: " + info.organizerEmail2);
      MailApp.sendEmail(info.organizerEmail2, subject, message);
      Logger.log("Email sent from " + Session.getActiveUser().getEmail() + " to " + info.organizerEmail2);
    }
  });
}
