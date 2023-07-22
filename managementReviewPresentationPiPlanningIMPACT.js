// Purpose: create and populate new PI Planning Management Review presentation for event attendees to populate

// NOTE: This script is currently utilized as a library in impactPiWeekPackage

//////////////////////////////////////////

// Source Calendar ID
var source_calendar_id = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com';
// Management Review Template ID
var template_id = '1M9R_Yds6OO6TAmEtmsXrtOSmeaCyDh7NBIYswRmIhwI';
// Placement Folder ID
var placement_folder_id = '169W64yI042Q24q4socXa4GhiQ7iY4a1f';

// Helper function to: Find and return the next "Management Review" event.
function getNextManagementReviewEvent() {
  var now = new Date();
  var calendar = CalendarApp.getCalendarById(source_calendar_id);
  var events = calendar.getEvents(now, new Date(now.getFullYear()+1, now.getMonth(), now.getDate()), {search: 'Management Review'});
  return events.length > 0 ? events[0] : null;
}

// Helper function to: Extract current PI number from file title.
function extractCurrentPINumber(eventTitle) {
  return eventTitle.split(' ')[1]; // "PI 23.4 IMPACT PI Planning Management Review" -> "23.4"
}

// Helper function to: Duplicate the IMPACT Management Review Template in the Placement folder and return the newly created file.
function duplicateTemplate(currentPINumber) {
  var templateFile = DriveApp.getFileById(template_id);
  var newFile = templateFile.makeCopy('IMPACT PI Planning ' + currentPINumber + ' Management Review', DriveApp.getFolderById(placement_folder_id));
  return newFile;
}

// helper function to format date in mm/dd/yy 
function formatDate(date) {
  var dd = String(date.getDate()).padStart(2, '0');
  var mm = String(date.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yy = String(date.getFullYear()).substr(-2);

  return mm + '/' + dd + '/' + yy;
}

// Primary function to: Create a new presentation.
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