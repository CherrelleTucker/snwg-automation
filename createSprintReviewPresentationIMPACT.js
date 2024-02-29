/*
Google Apps Script for Sprint Review Notifications

This script automates the creation of Sprint Review presentations from a template, tailored for the IMPACT project. It is designed to replace specific placeholders within the presentation with dynamic content related to the Sprint Review meeting.

Prerequisites:
1. Google Calendar with scheduled Sprint Review events.
2. Google Drive with a folder containing the presentation template.
3. Proper permissions set for the script to access Google Calendar and Google Drive.

Script Functions:
1. getCurrentSprintEvent - Fetches the current or next Sprint event from Google Calendar.
2. duplicateTemplate - Creates a copy of the Sprint Review presentation template.
3. populatePresentation - Fills in dynamic content in the copied template.
4. formatSprintReviewTitle - Formats the title of the Sprint Review presentation.
5. main - The main function to run the script.

Outputs:
1. A new Sprint Review presentation file in Google Drive.
2. The presentation is populated with details specific to the upcoming Sprint Review.

Post-Execution:
After execution, the script:
1. Logs the creation of the new presentation.
2. Provides a link to the newly created presentation in Google Drive.

Troubleshooting:
1. Ensure the script has the necessary permissions to access Google Calendar and Google Drive.
2. Check that the Sprint Review events are correctly named and scheduled in Google Calendar.
3. Verify that the presentation template contains the correct placeholders.

Notes:
1. This script should be run before each Sprint Review meeting.
2. Adjust the placeholders and date formats as necessary to match your template and locale.
3. Consider setting up a time-driven trigger to automate the execution of this script.
*/

///////////////////////////////////////////////////

// Global variables
var source_calendar_id = 'xxxxxxxxxxxx@group.calendar.google.com'; // IMPACT Project Increment Google Calendar
var template_id = 'xxxxxxxxxxxxxxxxx'; // Sprint Review Template ID
var placement_folder_id = 'xxxxxxxxxxxxxxxxxxxxxxxx'; // IMPACT Team Meetings Google Drive folder

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

// Jamboard functions removed 02/2024. Google is discontinuing the product later this year. 
// Helper function to: Duplicate the Jamboard Template in the Testing folder and return the newly created file.
/*function duplicateJamboardTemplate(currentSprintNumber) {
  var jamboard_template_id = '1fxtfrJKvVMwOHhQkZNXkSc5KMB-gafinTDtefz-htBs'; // Jamboard Template ID
  var jamboardTemplateFile = DriveApp.getFileById(jamboard_template_id);
  var newJamboardTitle = 'IMPACT-Kudos-Board _' + currentSprintNumber;
  var newJamboard = jamboardTemplateFile.makeCopy(newJamboardTitle, DriveApp.getFolderById(placement_folder_id));
  return newJamboard;

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
}*/

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
//    var newJamboard = duplicateJamboardTemplate(currentSprintNumber);
//    var jamboardUrl = 'https://jamboard.google.com/d/' + newJamboard.getId();

    var presentation = SlidesApp.openById(newFile.getId());

    var sprintEndDate = new Date(sprintEvent.getEndTime().getTime()); // calculate Sprint end date minus one day
    sprintEndDate.setDate(sprintEndDate.getDate());

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

    /* Slide 4
    updateHyperlinkInPresentation(newFile.getId(), 3, "IMPACT Jamboard for Kudos", jamboardUrl); // Example usage to set the hyperlink for the shape on Slide 4*/

    // Slide 10
    var slide10 = presentation.getSlides()[9]; // because indexing starts at 0
    slide10.replaceAllText('{{FY.Q}}', fiscalYearQuarter);
    slide10.replaceAllText('{{S}}', sprintNumber);

    // Slide 27
    var nextDueDate = formatDate(getNextDueDate()); // Get the formatted next due date
    var slideForDueDate = presentation.getSlides()[26]; // 
    slideForDueDate.replaceAllText('{{Blurb due date}}', nextDueDate);

  }
}

// Trigger function to check if Sprint Review event exists in the calendar and run createNewPresentation if it does. 
function checkForSprintReviewAndCreateFiles() {
  var today = new Date();
  // Check if today is Friday; 5 represents Friday in getDay() where Sunday is 0, Monday is 1, and so on.
  if (today.getDay() === 5) { 
    var calendar = CalendarApp.getCalendarById(source_calendar_id);
    var events = calendar.getEventsForDay(today, {
      search: 'Sprint Review'
    });

    if (events.length > 0) {
      Logger.log('Sprint Review found for today. Creating new presentation.');
      createNewPresentation();
    } else {
      Logger.log('No Sprint Review found for today. No presentation created.');
    }
  } else {
    Logger.log('Today is not Friday. No need to check for Sprint Review.');
  }
}

//////////////////////Testing///////////////////////////////////

function testCheckForSprintReviewAndCreateFilesOnSpecificDate(testDate) {
  var testDateObj = new Date(testDate);
  Logger.log('Testing for date: ' + testDateObj.toDateString());
  
  // Simulate as if the test date is the current date
  if (testDateObj.getDay() === 5) { // Check if the test date is a Friday
    var calendar = CalendarApp.getCalendarById(source_calendar_id);
    var events = calendar.getEventsForDay(testDateObj, {
      search: 'Sprint Review'
    });

    if (events.length > 0) {
      Logger.log('Sprint Review found for the test date. Would create new presentation.');
      // Uncomment the next line to actually create the presentation during testing
      // createNewPresentation();
    } else {
      Logger.log('No Sprint Review found for the test date. Would not create presentation.');
    }
  } else {
    Logger.log('Test date is not Friday. Script checks for Sprint Reviews only on Fridays.');
  }
}
testCheckForSprintReviewAndCreateFilesOnSpecificDate('2024-02-17'); // Replace with the date you want to test, format YYY-MM-DD
