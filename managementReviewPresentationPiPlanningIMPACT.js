// Purpose:
// Automate the creation of a new presentation for an "IMPACT PI Planning Management Review" based on the next scheduled event in the "IMPACT PI Google calendar." It retrieves the relevant data, such as the Program Increment (PI) number and the date of the review event, from the calendar event's title. The script then duplicates a specified template and populates it with the extracted data, including the fiscal year, quarter, and date of the review event. The end result is a new presentation ready for the management review, streamlining the process and ensuring consistency in the generated slides.

// TO NOTE:  
// Although this script is developed as a Google Apps Script standalone script. it is currently utilized as a library in impactPiWeekPackage. It is designed to operate independently and does not require any external application or service to function. 
 
// To Use:
// 1. Make a copy of the Google Apps Script: Open the script editor in your Google Workspace (formerly G Suite) account. Create a new script file and copy-paste the entire script into it.

// 2. Set up calendar and folder IDs: Replace the placeholder values in the global variables section with your specific calendar and folder IDs.Update the 'source_calendar_id' with the ID of your IMPACT PI Calendar. Update the 'placement_folder_id' with the ID of the folder where you want to store the generated presentations and Jamboards.

// 3. Template IDs: If you have your own presentation templates, replace the 'template_id with the IDs of your templates.

// 4. Save the script: Save the script and give it a descriptive name.

// 5. Run the 'createNewPresentation()' function: Click the "Run" button or use the keyboard shortcut "Ctrl + Enter" (Windows) or "Cmd + Enter" (Mac) to execute the script.

// 6. Grant permissions: The script will request permission to access your Google Calendar, Google Drive, and Google Slides. Click "Continue" and grant the necessary permissions.

// 7. Enjoy the automation:The script will automatically create a new "Management Review" presentation for the given PI. The presentation will be populated with relevant data.

// 8. Schedule the script (optional): If you want this process to run automatically, you can set up a time-based trigger to run the 'createNewPresentation()' function at specific intervals (e.g., weekly) to generate presentations.

// Please note that you need to be familiar with Google Apps Script and have the necessary permissions to access and modify Google Calendar, Google Drive, and Google Slides to use this script effectively.
// Make sure to review and customize the script to fit your specific use case before running it.


//////////////////////////////////////////

// Global variables
var source_calendar_id = 'xxxxxxxxxxxxxxxxxxxxxxxxxxx@group.calendar.google.com'; // IMPACT PI Google calendar
var template_id = 'xxxxxxxxxxxxxxxx'; // Management Review Template ID
var placement_folder_id = 'xxxxxxxxxxxxxxxxxxxxx'; // IMPACT PI Planning Google Drive folder

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
