// Purpose:
// This script automates the creation of an "Internal Planning Meeting Agenda" presentation for a specific Program Increment (PI) within the IMPACT project. It utilizes the IMPACT PI Calendar to extract the current PI number from the next "Final Presentation" event and then duplicates a presentation template. The script populates the new presentation with relevant data, including the PI number, key dates, Kudos Jamboard hyperlink, and Start/Stop/Continue (SSC) Jamboard hyperlink for retrospective purposes. Additionally, the script calculates and displays the dates for five two-week sprints followed by two one-week periods in Slide 5 of the presentation. Moreover, it creates two Jamboards for the PI Planning Retro, one for Kudos and one for SSC. The script duplicates the respective Jamboard templates, updates their titles with the current PI number, and provides hyperlinks to these Jamboards in the presentation.

// Future development:
// The script will need enhancement to ensure the dates on Slide 5 continue progressing beyond the end of 2023, without looping back to the beginning of the PI.

// To Note: While this script is currently utilized as a library in "impactPiWeekPackage," it is developed as a Google Apps Script standalone script. It is designed to operate independently and does not require any external application or service to function.

// To Use:
// 1. Make a copy of the Google Apps Script: Open the script editor in your Google Workspace (formerly G Suite) account. Create a new script file and copy-paste the entire script into it.
// 2. Set up calendar and folder IDs: Replace the placeholder values in the global variables section with your specific calendar and folder IDs. Update the 'source_calendar_id' with the ID of your IMPACT PI Calendar. Update the 'placement_folder_id' with the ID of the folder where you want to store the generated presentations and Jamboards.
// 3. Template IDs: If you have your own presentation and Jamboard templates, replace the 'template_id' and 'kudos_jamboard_template_id', with the IDs of your templates.
// 4. Save the script: Save the script and give it a descriptive name.
// 5. Run the 'createNewPresentation()' function: Click the "Run" button or use the keyboard shortcut "Ctrl + Enter" (Windows) or "Cmd + Enter" (Mac) to execute the script.
// 6. Grant permissions: The script will request permission to access your Google Calendar, Google Drive, and Google Slides. Click "Continue" and grant the necessary permissions.
// 7. Enjoy the automation: The script will automatically create a new "Internal Planning Meeting Agenda" presentation and two Jamboards for the given PI. The presentation will be populated with relevant data, and the Jamboards will be hyperlinked in the slides as needed.
// 8. Schedule the script (optional): If you want this process to run automatically, you can set up a time-based trigger to run the 'createNewPresentation()' function at specific intervals (e.g., weekly) to generate new agendas and Jamboards.
// Please note that you need to be familiar with Google Apps Script and have the necessary permissions to access and modify Google Calendar, Google Drive, and Google Slides to use this script effectively. Make sure to review and customize the script to fit your specific use case before running it.

////////////////////////////////////////

// Global variables
var source_calendar_id = 'c_e6e532cefc5ddfdd7f3c715e7a07326607cd240d951991f6a4e3b87653e67ef3@group.calendar.google.com'; // IMPACT PI calendar
var template_id = '1JtnXgRM85G7fBJ0nbM4VlNQjCxvPmh8-6jH_e7JXPiM'; // Welcome Template ID
var placement_folder_id = '169W64yI042Q24q4socXa4GhiQ7iY4a1f'; // IMPACT Presentations>PI Planning for easy identification and error checking prior to being placed in appropriate FY folder.

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

// helper function to calculate the previous fiscal year and quarter
function calculatePreviousFYQ(fiscalYear, quarter) {
  var previousQuarter, previousFiscalYear;

  if (quarter === '1') {
    previousQuarter = '4';
    previousFiscalYear = (parseInt(fiscalYear) - 1).toString(); // Decrement the fiscal year
  } else {
    previousQuarter = (parseInt(quarter) - 1).toString(); // Decrement the quarter
    previousFiscalYear = fiscalYear;
  }

  return previousFiscalYear + '.' + previousQuarter;
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

    // Calculate the previous fiscal year and quarter
    var previousFYQ = calculatePreviousFYQ(fiscalYear, quarter);
    
    var newFile = duplicateTemplate(currentPINumber);
    
    var presentation = SlidesApp.openById(newFile.getId());
    var slides = presentation.getSlides();
    
    // Update placeholders in the presentation
    for (var i = 0; i < slides.length; i++) {
      slides[i].replaceAllText('{{FY.Q}}', fiscalYear + '.' + quarter);
      slides[i].replaceAllText('{{FY.Q-1}}', previousFYQ); // Replace the new placeholder
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

/////////////Testing//////////////////

// Testing function to console log output for specific placeholders
function testSprintDateCalculations(testDate) {
  var simulatedWelcomeEventStart = new Date(testDate);
  var simulatedWelcomeEventEnd = new Date(testDate); // You can adjust this if the end date is different

  var sprintDates = calculateSprintDates(simulatedWelcomeEventStart, simulatedWelcomeEventEnd);

  // Log the outputs for the placeholders
  console.log('Test Date: ' + testDate);
  console.log('{{Week 1-2}}: ' + sprintDates[0]);
  console.log('{{Week 3-4}}: ' + sprintDates[1]);
  console.log('{{Week 5-6}}: ' + sprintDates[2]);
  console.log('{{Week 7-8}}: ' + sprintDates[3]);
  console.log('{{Week 9-10}}: ' + sprintDates[4]);
  console.log('{{Week 11}}: ' + sprintDates[5]);
  console.log('{{Week 12}}: ' + sprintDates[6]);
}

// Example usage of the testing function
testSprintDateCalculations('2024-012-15'); // Replace with any date you want to test
