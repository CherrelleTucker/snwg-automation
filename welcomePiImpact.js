// Purpose: create new PI Planning Welcome welcome presentation from template. 
// Current issue: dates on slide 12 will not progress in to 2024

// source calendar ID: c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com
// Welcome template id: 1JtnXgRM85G7fBJ0nbM4VlNQjCxvPmh8-6jH_e7JXPiM
// placement folder: 1V40h1Df4TMuuGzTMiLHxyBRPC-XJhQ10

// Access IMPACT PI calendar. Find today. Identify the next "Welcome" event. This event title - "PI FY.Q IMPACT PI Planning Welcome" contains the currentPINumber;  example: "PI 23.4 IMPACT PI Planning Welcome" = currentPINumber "23.4" Format: FY.PI This information will create the currentPINumber that informs the rest of the script. 
// Global placeholder text for all slides: FY = Fiscal year, Q = Quarter
// // Access IMPACT PI Welcome Template. Duplicate IMPACT Welcome Template in the Testing folder

// Name newly created file "IMPACT PI Planning {{FY.Q}} Welcome"

// Slide 1 {{FY.Q}}, {{Month Day, Year of welcome event}} Calendar accessed, welcome event found, date returned in day month, year format
// Slide 2 {{FY.Q}}
// Slide 3 {{FY.Q}}
// Slide 4 {{FY.Q}}
// Slide 5 {{FY.Q}}
// Slide 7 {{FY.Q}}
// Slide 11 hyperlink shape containing "Full agenda HERE" - search in PI Planning folder for document that matches the current FY.Q and contains "Agenda" in the title
// Slide 12 {{Q}} access calendar, identify sprint events and duration. Format MM/DD/YY - MM/DD/YY
  // {{Week 1-2}}
  // {{Week 3-4}}
  // {{Week 5-6}}
  // {{Week 7-8}}
  // {{Week 9-10}}
  // {{Week 11}}
  // {{Week 12}}
// Slide 12 {{Q}}
// Slide 14 {{FY.Q}}

// Source Calendar ID
var source_calendar_id = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com';
// Welcome Template ID
var template_id = '1JtnXgRM85G7fBJ0nbM4VlNQjCxvPmh8-6jH_e7JXPiM';
// Placement Folder ID
var placement_folder_id = '1V40h1Df4TMuuGzTMiLHxyBRPC-XJhQ10';

// Helper function to: Find and return the next "Welcome" event.
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

function formatDate(date) {
  var dd = String(date.getDate()).padStart(2, '0');
  var mm = String(date.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yy = String(date.getFullYear()).substr(-2);

  return mm + '/' + dd + '/' + yy;
}


// Helper function to get the date of the next Monday after a given date
function getNextMonday(date) {
  var day = date.getDay();
  var diff = 8 - day; // Calculate days until the next Monday (Monday = 1)
  date.setDate(date.getDate() + diff);
  return date;
}

// Helper function to get the date of the second Friday after a given date
function getSecondFriday(date) {
  var day = date.getDay();
  var diff = 12 - day + 7; // Calculate days until the second Friday (Friday = 5) and add 7 days for the next week
  date.setDate(date.getDate() + diff);
  return date;
}

// Helper function to: Calculate sprint dates.
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

  // Check if the PI Planning Week is in December and adjust for transitioning into January
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