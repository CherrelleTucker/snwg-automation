// Purpose: create and populate new PI Planning Welcome presentation from template. 
// Current issue: dates on slide 12 will not progress to 2024

// NOTE: This script is currently utilized as a library in impactPiWeekPackage

////////////////////////////////////////

// Source Calendar ID
var source_calendar_id = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com';
// Welcome Template ID
var template_id = '1JtnXgRM85G7fBJ0nbM4VlNQjCxvPmh8-6jH_e7JXPiM';
// Placement Folder ID
var placement_folder_id = '169W64yI042Q24q4socXa4GhiQ7iY4a1f'; //Placed in IMPACT Presentations>PI Planning for easy identification and error checking prior to being placed in appropriate FY folder.

// Helper function to find and return the next "Welcome" event.
function getNextWelcomeEvent() {
  var now = new Date();
  var calendar = CalendarApp.getCalendarById(source_calendar_id);
  var events = calendar.getEvents(now, new Date(now.getFullYear()+1, now.getMonth(), now.getDate()), {search: 'Welcome'});
  return events.length > 0 ? events[0] : null;
}

// Helper function to: Extract current PI number.
function extractCurrentPINumber(eventTitle) {
  return eventTitle.split(' ')[1]; // "PI 23.4 IMPACT PI Planning Welcome" -> "23.4"
}

// Helper function to: Duplicate the IMPACT Welcome Template in the Placement folder and return the newly created file.
function duplicateTemplate(currentPINumber) {
  var templateFile = DriveApp.getFileById(template_id);
  var newFile = templateFile.makeCopy('IMPACT PI Planning ' + currentPINumber + ' Welcome', DriveApp.getFolderById(placement_folder_id));
  return newFile;
}

// Helper function to: Find an agenda file in the PI Planning folder.
function findAgendaFile(currentPINumber) {
  var folder = DriveApp.getFolderById(placement_folder_id);
  var files = folder.getFiles();
  var regex = new RegExp(currentPINumber + ".*Agenda", "i");
  while (files.hasNext()) {
    var file = files.next();
    if (regex.test(file.getName())) {
      return file;
    }
  }
  return null;
}

// helper function to format date in mm/dd/yy 
function formatDate(date) {
  var dd = String(date.getDate()).padStart(2, '0');
  var mm = String(date.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yy = String(date.getFullYear()).substr(-2);

  return mm + '/' + dd + '/' + yy;
}


// Helper function to get the date of the next Monday after a given date for Slide 11
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
  // Assuming a PI structure of 5 2-week sprints followed by 2 1-week periods.
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

  // Check if the PI Planning Week is in December and adjust for transitioning into January for Slide 11. ISSUE: currently not returning proper 2024 dates
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

// Main function to: Create a new presentation.
function createNewPresentation() {
  var welcomeEvent = getNextWelcomeEvent();
  if (welcomeEvent) {
    var title = welcomeEvent.getTitle(); // e.g., "PI 23.4 IMPACT PI Planning Welcome"
    var currentPINumber = title.split(' ')[1]; // e.g., "23.4"
    var fiscalYear = currentPINumber.split('.')[0]; // e.g., "23"
    var quarter = currentPINumber.split('.')[1]; // e.g., "4"
    
    var newFile = duplicateTemplate(currentPINumber);
    
    var presentation = SlidesApp.openById(newFile.getId());
    var slides = presentation.getSlides();
    
    // Global Placeholder
    for (var i = 0; i < slides.length; i++) {
      slides[i].replaceAllText('{{FY.Q}}',fiscalYear + '.' + quarter);
    }
    
    // Slide 1
    var slide1 = slides[0];
    slide1.replaceAllText('{{Month Day, Year of welcome event}}', welcomeEvent.getStartTime().toLocaleDateString());
    
    // Slide 12
    var slide12 = slides[11];
    var sprintDates = calculateSprintDates(welcomeEvent.getStartTime(), welcomeEvent.getEndTime());

    // Assigning each two week period to the respective placeholders
    for (var i = 0; i < sprintDates.length; i++) {
      if (i === sprintDates.length - 2) { // Special case for last 2 weeks
        slide12.replaceAllText('{{Week 11}}', sprintDates[i]);
      } else if (i === sprintDates.length - 1) {
        slide12.replaceAllText('{{Week 12}}', sprintDates[i]);
      } else {
        slide12.replaceAllText('{{Week ' + ((i * 2) + 1) + '-' + ((i * 2) + 2) + '}}', sprintDates[i]);
      }
    }
  }
}