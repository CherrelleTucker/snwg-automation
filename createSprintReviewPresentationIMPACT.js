// Purpose: create and populate bi-weekly Sprint review slides and Jamboard for IMPACT Sprint reviews. Newly created files are placed in the IMPACT Team Meetings folder

// To use: run function "createNewPresentation". Built-in trigger executes "createNewPresentation" one week before the next Sprint Review 


///////////////////////////////////////////////////

// Source Calendar ID
var source_calendar_id = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com';
// Sprint Review Template ID
var template_id = '1UxcyJtzCgDWnJc0Nr3UxtMKESU1cMowblvgq6I_jf0U';
// Jamboard Template ID
var jamboard_template_id = '1fxtfrJKvVMwOHhQkZNXkSc5KMB-gafinTDtefz-htBs';
// Testing Folder ID
var placement_folder_id = '1UmjkjY5RTRYFOQEt10mwU8trJQ389Jum';

// Helper function to: Find and return the Current Sprint event.
function getCurrentSprintEvent() {
  var now = new Date();
  var calendar = CalendarApp.getCalendarById(source_calendar_id);
  var events = calendar.getEventsForDay(now, { search: 'Sprint' });
  return events.length > 0 ? events[0] : null;
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

// Main function to: Create a new presentation.
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

    // Slide 1
    var slide1 = presentation.getSlides()[0];
    slide1.replaceAllText('{{currentSprintNumber}}', currentSprintNumber);
    slide1.replaceAllText('{{Date of Sprint Review}}', sprintEvent.getEndTime().toDateString());
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
  }
}

// Trigger function to execute one week before the next Sprint Review event on the IMPACT PI Calendar
function executeOneWeekBeforeSprintReview() {
  var sourceCalendar = CalendarApp.getCalendarById(sourceCalendarId);
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
}
