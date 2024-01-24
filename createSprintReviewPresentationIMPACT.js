// Purpose: 
// To automate the creation of a presentation and Jamboard (a collaborative whiteboard tool) for Sprint Review events in the "IMPACT Project Increment" Google Calendar. The script is triggered one week before each Sprint Review, saving time and ensuring consistency in the document creation process for the IMPACT Project Increment team.

// To note: 
// This script is developed as a Google Apps Script standalone script. It is designed to operate independently and does not require any external application or service to function. It is a self-contained piece of code with a time-based daily trigger.

// To use: 
// Instructions for Using the "IMPACT Sprint Review Automation" Script
// Make a Copy of the Script:
// Open the IMPACT Sprint Review Automation Script. Click "File" > "Make a copy..." and rename it as desired.
// Configure Global Variables: Replace the source_calendar_id, template_id, jamboard_template_id, and placement_folder_id with your own Google Calendar and Drive IDs.
// Update Slide numbering to match your template numbering in the primary function createNewPresentation if needed.
// Save and Run the Script: Click the floppy disk icon (or press Ctrl + S) to save the script. Click the play button (▶) in the toolbar to run the script manually for the first time.
// Grant Necessary Permissions: If prompted, grant the script the necessary permissions to access your Google Calendar and Drive.
// Test the Script (Optional): To test the script functionality, manually call the createNewPresentation function by clicking the play button (▶) again.
// Set Up the Trigger (Optional): The script will automatically create and manage a time-driven trigger to run one week before each upcoming Sprint Review event in your Google Calendar.

///////////////////////////////////////////////////

// Global variables
var source_calendar_id = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com'; // IMPACT Project Increment Google Calendar
var template_id = '1UxcyJtzCgDWnJc0Nr3UxtMKESU1cMowblvgq6I_jf0U'; // Sprint Review Template ID
var jamboard_template_id = '1fxtfrJKvVMwOHhQkZNXkSc5KMB-gafinTDtefz-htBs'; // Jamboard Template ID
var placement_folder_id = '1UmjkjY5RTRYFOQEt10mwU8trJQ389Jum'; // IMPACT Team Meetings Google Drive folder

// Helper function to: Find and return the Current Sprint event.
function getCurrentSprintEvent() {
  var now = new Date();
  var calendar = CalendarApp.getCalendarById(source_calendar_id);
  var events = calendar.getEventsForDay(now, { search: 'Sprint' });
  Logger.log('Number of events found: ' + events.length); // Log the number of events found

  if (events.length > 0) {
    var event = events[0];
    Logger.log('Selected event title: ' + event.getTitle()); // Log the title of the selected event
    Logger.log('Event start time: ' + event.getStartTime()); // Log start time
    Logger.log('Event end time: ' + event.getEndTime()); // Log end time
    return event;
  } else {
    Logger.log('No events found for current day.');
    return null;
  }
}

// Helper function to: Duplicate the IMPACT Sprint Review Template in the Testing folder and return the newly created file.
function duplicateTemplate(currentSprintNumber) {
  var templateFile = DriveApp.getFileById(template_id);
  var newFile = templateFile.makeCopy('IMPACT Sprint Review_' + currentSprintNumber, DriveApp.getFolderById(placement_folder_id));
  return newFile;
}

// Helper function to: Duplicate the Jamboard Template in the Testing folder and return the newly created file.
function duplicateJamboardTemplate(currentSprintNumber) {
  var jamboardTemplateFile = DriveApp.getFileById(jamboard_template_id);
  var newJamboardTitle = 'IMPACT-Kudos-Board _' + currentSprintNumber;
  var newJamboard = jamboardTemplateFile.makeCopy(newJamboardTitle, DriveApp.getFolderById(placement_folder_id));
  return newJamboard;
}

// Helper function to: Update the title of the newly created Jamboard.
function updateJamboardTitle(jamboardId, currentSprintNumber) {
  var newTitle = 'IMPACT-Kudos-Board _' + currentSprintNumber;
  var jamboard = SlidesApp.openById(jamboardId);
  jamboard.setName(newTitle);
}

// Helper function to: Update the hyperlink for the shape in the Sprint Review Presentation with the newly created Jamboard hyperlink.
function updateHyperlinkInPresentation(presentationId, slideIndex, shapeText, jamboardUrl) {
  var presentation = SlidesApp.openById(presentationId);
  var slides = presentation.getSlides();

  if (slideIndex >= 0 && slideIndex < slides.length) {
    var slide = slides[slideIndex];
    var shapes = slide.getShapes();
    for (var i = 0; i < shapes.length; i++) {
      var shape = shapes[i];
      var textRange = shape.getText();
      var text = textRange.asString();
      if (text.includes(shapeText)) {
        shape.setLinkUrl(jamboardUrl);
        break;
      }
    }
  }
}

// Helper function to get today's date
function getToday() {
  return new Date();
}

// Helper function to calculate the next most recent blurb due date
function getNextDueDate() {
  const startDate = new Date("December 4, 2023 12:00:00");
  const today = getToday();

  // Calculate the difference in days from the start date
  const daysDiff = Math.floor((today - startDate) / (24 * 60 * 60 * 1000));

  // Calculate the number of bi-weeks since the start date
  const biWeeksSinceStart = Math.floor(daysDiff / 14);

  // Calculate the next due date
  const nextDueDate = new Date(startDate);
  nextDueDate.setDate(startDate.getDate() + (biWeeksSinceStart + 2) * 14);
  nextDueDate.setHours(12, 0, 0, 0);

  // Adjust the date to the previous Friday
  nextDueDate.setDate(nextDueDate.getDate() - 3);

  return nextDueDate;
}

// Helper function to format the blurb due date in the standard format
function formatDate(date) {
  const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  return date.toLocaleDateString('en-US', options);
}

// Primary function to: Create a new presentation.
function createNewPresentation() {
  var sprintEvent = getCurrentSprintEvent();
  if (sprintEvent) {
    var title = sprintEvent.getTitle(); // e.g., "23.3 Sprint 4"
    var titleComponents = title.split('Sprint'); // e.g., ["23.3 ", " 4"]
    var currentSprintNumber = titleComponents[0].trim() + '.' + titleComponents[1].trim(); // e.g., "23.3.4"

    var fiscalYearQuarter = currentSprintNumber.split('.')[0] + '.' + currentSprintNumber.split('.')[1].charAt(0); // e.g., "23.3"
    var sprintNumber = currentSprintNumber.split('.')[2]; // e.g., "4"

    var newFile = duplicateTemplate(currentSprintNumber);
    var newJamboard = duplicateJamboardTemplate(currentSprintNumber);
    var jamboardUrl = 'https://jamboard.google.com/d/' + newJamboard.getId();

    var presentation = SlidesApp.openById(newFile.getId());

    var sprintEndDate = new Date(sprintEvent.getEndTime().getTime()); // calculate Sprint end date minus one day
    sprintEndDate.setDate(sprintEndDate.getDate() - 1);

    // Slide 1
    var slide1 = presentation.getSlides()[0];
    slide1.replaceAllText('{{currentSprintNumber}}', currentSprintNumber);
    slide1.replaceAllText('{{Date of Sprint Review}}', sprintEndDate.toDateString());
    slide1.replaceAllText('{{FY}}', new Date().getFullYear().toString().slice(-2));

    // Slide 3
    var slide3 = presentation.getSlides()[2];
    var nextMonth = new Date();
    nextMonth.setMonth(nextMonth.getMonth() + 1);
    var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    slide3.replaceAllText('{{1 Month}}', monthNames[nextMonth.getMonth()]);

    var twoMonthsLater = new Date();
    twoMonthsLater.setMonth(twoMonthsLater.getMonth() + 2);
    slide3.replaceAllText('{{2 Month}}', monthNames[twoMonthsLater.getMonth()]);

    // Slide 4
    updateHyperlinkInPresentation(newFile.getId(), 3, "IMPACT Jamboard for Kudos", jamboardUrl); // Example usage to set the hyperlink for the shape on Slide 4

    // Slide 9
    var slide9 = presentation.getSlides()[8]; // because indexing starts at 0
    slide9.replaceAllText('{{FY.Q}}', fiscalYearQuarter);
    slide9.replaceAllText('{{S}}', sprintNumber);

    // Slide 25
    var nextDueDate = formatDate(getNextDueDate()); // Get the formatted next due date
    var slideForDueDate = presentation.getSlides()[24]; // 
    slideForDueDate.replaceAllText('{{Blurb due date}}', nextDueDate);

  }
}

// Trigger function to execute one week before the next Sprint Review event on the IMPACT PI Calendar
/*function executeOneWeekBeforeSprintReview() {
  var sourceCalendar = CalendarApp.getCalendarById(source_calendar_id);
  var today = new Date();
  var oneWeekFromToday = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
  var events = sourceCalendar.getEvents(today, oneWeekFromToday);

  // Find the next Sprint Review event
  var nextSprintReviewEvent = null;
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    if (event.getTitle().toLowerCase().indexOf('sprint review') !== -1) {
      nextSprintReviewEvent = event;
      break;
    }
  }

  if (nextSprintReviewEvent) {
    // Calculate the date one week before the Sprint Review event
    var oneWeekBeforeSprintReview = new Date(nextSprintReviewEvent.getStartTime().getTime() - 7 * 24 * 60 * 60 * 1000);

    // Set up the time-driven trigger to call the libraries one week before the PI Welcome event
    ScriptApp.newTrigger(createNewPresentation)
      .timeBased()
      .at(oneWeekBeforeSprintReview)
      .create();
  } else {
    Logger.log("No upcoming Sprint Review event found within the next week.");
  }
}*/