// Purpose: create new PI Planning Management Review Template for event attendees to populate
// source calendar ID: c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com
// Management Review Template ID: 1M9R_Yds6OO6TAmEtmsXrtOSmeaCyDh7NBIYswRmIhwI
// Placement folder ID: 1V40h1Df4TMuuGzTMiLHxyBRPC-XJhQ10 <<-- update to permenant folder prior to deployment

// Access IMPACT PI calendar. Find today. Identify the next "Management Review" event. This event title - ""PI FY.Q IMPACT PI Planning Management Review" contains the currentPINumber;  example: "PI 24.1 IMPACT PI Planning Management Review" = currentPINumber "24.1" Format: FY.Q This information will create the currentPINumber that informs the rest of the script. 

// Global placeholder text for all slides: FY = Fiscal year, Q = Quarter
// // Access IMPACT PI Management Review Template. Duplicate IMPACT Management Review Template in the Placement folder

// Name newly created file "IMPACT PI Planning {{FY.Q}} Management Review"
 
// replace {{FY.Q}} in title and all slides
// replace {{Month Day, Year of Review Event}} on slide 1: Calendar accessed, Management Review event found, date returned in day month, year format

// Source Calendar ID
var source_calendar_id = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com';
// Management Review Template ID
var template_id = '1M9R_Yds6OO6TAmEtmsXrtOSmeaCyDh7NBIYswRmIhwI';
// Placement Folder ID
var placement_folder_id = '1V40h1Df4TMuuGzTMiLHxyBRPC-XJhQ10'; //<<-- update to permenant folder prior to deployment

// Helper function to: Find and return the next "Management Review" event.
function getNextManagementReviewEvent() {
  var now = new Date();
  var calendar = CalendarApp.getCalendarById(source_calendar_id);
  var events = calendar.getEvents(now, new Date(now.getFullYear()+1, now.getMonth(), now.getDate()), {search: 'Management Review'});
  return events.length > 0 ? events[0] : null;
}

// Helper function to: Extract current PI number.
function extractCurrentPINumber(eventTitle) {
  return eventTitle.split(' ')[1]; // "PI 23.4 IMPACT PI Planning Management Review" -> "23.4"
}

// Helper function to: Duplicate the IMPACT Management Review Template in the Placement folder and return the newly created file.
function duplicateTemplate(currentPINumber) {
  var templateFile = DriveApp.getFileById(template_id);
  var newFile = templateFile.makeCopy('IMPACT PI Planning ' + currentPINumber + ' Management Review', DriveApp.getFolderById(placement_folder_id));
  return newFile;
}

function formatDate(date) {
  var dd = String(date.getDate()).padStart(2, '0');
  var mm = String(date.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yy = String(date.getFullYear()).substr(-2);

  return mm + '/' + dd + '/' + yy;
}

// Main function to: Create a new presentation.
function createNewPresentation() {
  var managementReviewEvent = getNextManagementReviewEvent();
  if (managementReviewEvent) {
    var title = managementReviewEvent.getTitle(); // e.g., "PI 23.4 IMPACT PI Planning Management Review"
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
    slide1.replaceAllText('{{Month Day, Year of Review event}}', managementReviewEvent.getStartTime().toLocaleDateString());
    
  }
}