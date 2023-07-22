// Purpose: This script automates the creation of a new PI (Program Increment) Planning Final Presentation along with related Jamboards for a specific event on a Google Calendar. The generated presentation and Jamboards serve as templates for event attendees to populate with relevant information.

// NOTE: This script is currently utilized as a library in impactPiWeekPackage

// Issue: Slide 5 dates: same calculation as Welcome Script. D:

////////////////////////////////////////////

// Source Calendar ID
var source_calendar_id = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com';
// Final Presentation Template ID
var template_id = '1j5KLp01q8pI89E3PN2QIGlvmA0-rD2O8bONWwG21IMY';
// Kudos Jamboard template ID
var kudos_jamboard_template_id = '1fxtfrJKvVMwOHhQkZNXkSc5KMB-gafinTDtefz-htBs';
// SSC Jamboard template ID
var ssc_jamboard_template_id = '1rQBgG43-PKWj2OOIl3G_Pocb7hHQRWOH46RVGK-XD-c';
// Placement Folder ID
var placement_folder_id = '1V40h1Df4TMuuGzTMiLHxyBRPC-XJhQ10'; //<<-- update to permanent folder prior to deployment


// Helper function to: Find and return the next "Final Presentation" event.
function getNextFinalPresentationEvent() {
  var now = new Date();
  var calendar = CalendarApp.getCalendarById(source_calendar_id);
  // Search for events with the title "Final Presentation" starting from the current date.
  var events = calendar.getEvents(now, new Date(now.getFullYear()+1, now.getMonth(), now.getDate()), {search: 'Final Presentation'});
  return events.length > 0 ? events[0] : null;
}

// Helper function to extract current PI number from the event title.
function extractCurrentPiNumber(eventTitle) {
  return eventTitle.split(' ')[1]; // "PI 23.4 IMPACT PI Planning Final Presentation" -> "23.4"
}

// Helper function to duplicate the IMPACT Final Presentation Template in the Placement folder and return the newly created file.
function duplicateTemplate(currentPiNumber) {
  var templateFile = DriveApp.getFileById(template_id);
  var newFile = templateFile.makeCopy('IMPACT PI Planning ' + currentPiNumber + ' Final Presentation', DriveApp.getFolderById(placement_folder_id));
  return newFile;
}

// Helper function to format the date as "MM/DD/YY"
function formatDate(date) {
  var dd = String(date.getDate()).padStart(2, '0');
  var mm = String(date.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yy = String(date.getFullYear()).substr(-2);

  return mm + '/' + dd + '/' + yy;
}

// Slide 5 crap

// Helper function to calculate the dates based on a PI structure of 5 two-week sprints followed by 2 one-week periods for Slide 11
function getNextMonday(date) {
  var day = date.getDay();
  var diff = 8 - day; // Calculate days until the next Monday (Monday = 1)
  date.setDate(date.getDate() + diff);
  return date;
}

// Helper function to get the date of the second Friday after a given date for Slide 11
function getSecondFriday(date) {
  var day = date.getDay();
  var diff = 12 - day + 7; // Calculate days until the second Friday (Friday = 5) and add 7 days for the next week
  date.setDate(date.getDate() + diff);
  return date;
}

// Helper function to: Calculate sprint dates for Slide 11
function calculateSprintDates(welcomeEventStart, welcomeEventEnd) {
  var sprintDates = [];
  var nextDate;

  var firstMonday = getNextMonday(new Date(welcomeEventEnd));
  var secondFriday = getSecondFriday(new Date(welcomeEventEnd));

  for (var i = 0; i < 5; i++) { // 5 2-week sprints
    nextDate = new Date(firstMonday);
    nextDate.setDate(nextDate.getDate() + 11);  // Increase by 11 days for a 2-week sprint (Monday to second Friday = 11 days)
    sprintDates.push(formatDate(firstMonday) + ' - ' + formatDate(nextDate));
    firstMonday = new Date(nextDate);
    firstMonday.setDate(firstMonday.getDate() + 3); // Start the next sprint on the following Monday
  }

  for (var i = 0; i < 2; i++) { // 2 1-week periods
    nextDate = new Date(secondFriday);
    nextDate.setDate(nextDate.getDate() + 10);  // Increase by 10 days for a 1-week period (Friday to second Thursday = 10 days)
    sprintDates.push(formatDate(secondFriday) + ' - ' + formatDate(nextDate));
    secondFriday = new Date(nextDate);
    secondFriday.setDate(secondFriday.getDate() + 3); // Start the next period on the following Monday
  }

  // Check if the PI Planning Week is in December and adjust for transitioning into January for Slide 11
  // ISSUE: currently not returning proper dates after the new year, returning to earlier in the PI
  if (secondFriday.getMonth() === 11 && secondFriday.getFullYear() !== welcomeEventEnd.getFullYear()) {
    for (var i = 0; i < sprintDates.length; i++) {
      var startDate = new Date(sprintDates[i].split(' - ')[0]);
      var endDate = new Date(sprintDates[i].split(' - ')[1]);
      startDate.setFullYear(startDate.getFullYear() + 1);
      endDate.setFullYear(endDate.getFullYear() + 1);
      sprintDates[i] = formatDate(startDate) + ' - ' + formatDate(endDate);
    }
  }

  return sprintDates;
}

////// Kudos Jamboard

// Helper function to duplicate the Kudos Jamboard Template in the Placement folder and return the newly created file.
function duplicateKudosJamboardTemplate(currentPiNumber) {
  var kudosJamboardTemplateFile = DriveApp.getFileById(kudos_jamboard_template_id);
  var newKudosJamboardTitle = 'IMPACT-Final-Presentation-Kudos-Board _' + currentPiNumber;
  var newKudosJamboard = kudosJamboardTemplateFile.makeCopy(newKudosJamboardTitle, DriveApp.getFolderById(placement_folder_id));
  return newKudosJamboard;
}

// Helper function to update the title of the newly created Kudos Jamboard.
function updateKudosJamboardTitle(kudosJamboardId, currentPiNumber) {
  var newTitle = 'Final-Presentation-' + currentPiNumber;
  var kudosJamboard = SlidesApp.openById(kudosJamboardId);
  kudosJamboard.setName(newTitle);
}

// Helper function to update the hyperlink for the shape in the Sprint Review Presentation with the newly created Kudos Jamboard hyperlink.
function updateKudosJamboardHyperlinkInPresentation(presentationId, slideIndex, shapeText, kudosJamboardUrl) {
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
        shape.setLinkUrl(kudosJamboardUrl);
        break;
      }
    }
  }
}

////// Stop/Start/Continue (SSC) Jamboard

// Helper function to duplicate the SSC Jamboard Template in the Testing folder and return the newly created files.
function duplicateSccJamboardTemplate(currentPiNumber) {
  var sscJamboardTemplateFile = DriveApp.getFileById(ssc_jamboard_template_id);
  var newSscJamboardTitle = 'Start/Stop/Continue for PI Planning ' + currentPiNumber + ' Retro';
  var newSscJamboard = sscJamboardTemplateFile.makeCopy(newSscJamboardTitle, DriveApp.getFolderById(placement_folder_id));
  return newSscJamboard;
}

// Helper function to update the title of the newly created SSC Jamboard.
function updateSccJamboardTitle(sccJamboardId, currentPiNumber) {
  var newTitle = 'Start/Stop/Continue for PI Planning' + currentPiNumber + 'Retro';
  var sccJamboard = SlidesApp.openById(sccJamboardId);
  sccJamboard.setName(newTitle);
}

// Helper function to update the hyperlink for the shape in the Sprint Review Presentation with the newly created SCC Jamboard hyperlink.
function updateSccJamboardHyperlinkInPresentation(presentationId, slideIndex, shapeText, sccJamboardUrl) {
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
        shape.setLinkUrl(sccJamboardUrl);
        break;
      }
    }
  }
}


// Main function to: Create a new presentation and Jamboards.
function createNewPresentation() {
  // Get the next "Final Presentation" event from the source calendar.
  var managementReviewEvent = getNextFinalPresentationEvent();
  if (managementReviewEvent) {
    // Extract the current PI number from the event title.
    var title = managementReviewEvent.getTitle(); // e.g., "PI 23.4 IMPACT PI Planning Final Presentation"
    var currentPiNumber = title.split(' ')[1]; // e.g., "23.4"
    var fiscalYear = currentPiNumber.split('.')[0]; // e.g., "23"
    var quarter = currentPiNumber.split('.')[1]; // e.g., "4"

    // Duplicate and update presentation
    var newFile = duplicateTemplate(currentPiNumber);
    var presentation = SlidesApp.openById(newFile.getId());
    var slides = presentation.getSlides();

    // Global Placeholder
    for (var i = 0; i < slides.length; i++) {
      slides[i].replaceAllText('{{FY.Q}}', fiscalYear + '.' + quarter);
    }

    // Slide 1
    var slide1 = slides[0];
    slide1.replaceAllText('{{Month Day, Year of Review event}}', managementReviewEvent.getStartTime().toLocaleDateString());

    // Slide 5: Populate Key Dates
    var slide5 = slides[4]; // Slide 5 is at index 4 (0-based index)
    var sprintDates = calculateSprintDates(managementReviewEvent.getStartTime(), managementReviewEvent.getEndTime());

    // Assigning each two week period to the respective placeholders
    for (var i = 0; i < sprintDates.length; i++) {
      if (i === sprintDates.length - 2) { // Special case for last 2 weeks
        slide5.replaceAllText('{{Week 11}}', sprintDates[i]);
      } else if (i === sprintDates.length - 1) {
        slide5.replaceAllText('{{Week 12}}', sprintDates[i]);
      } else {
        slide5.replaceAllText('{{Week ' + ((i * 2) + 1) + '-' + ((i * 2) + 2) + '}}', sprintDates[i]);
      }
    }

    // Slide 7: Duplicate and update Kudos Jamboard
    var kudosJamboard = duplicateKudosJamboardTemplate(currentPiNumber);
    updateKudosJamboardHyperlinkInPresentation(newFile.getId(), 6, 'IMPACT Jamboard for Kudos', kudosJamboard.getUrl());

    // Slide 25: Duplicate and update SSC Jamboard
    var sccJamboard = duplicateSccJamboardTemplate(currentPiNumber);
    updateSccJamboardHyperlinkInPresentation(newFile.getId(), 24, 'Start/Stop/Continue Jamboard', sccJamboard.getUrl());
  }
}