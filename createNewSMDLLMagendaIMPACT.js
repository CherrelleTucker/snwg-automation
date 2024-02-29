// Purpose:  
// Create a bi-weekly agenda for the "SMD Large Language Model WG" (Working Group) by utilizing a template agenda document and populating it with the appropriate meeting date. The script will only create the agenda if the current week number is even (bi-weekly) to align with the group's bi-weekly meeting schedule.

// Future development: 
// none currently identified

// To note: 
// This script is developed as a Google Apps Script standalone script. It is developed to operate independently and does not require any external application or service to function. It is a self-contained piece of code with a time-based daily trigger.

// To use: 
// 1. Modify the global variables AGENDA_TEMPLATE_ID and AGENDA_FOLDER_ID to the respective IDs of your agenda template document and destination folder where the agendas should be stored. These variables should be set with the appropriate values specific to your Google Drive.
// 2. The primary function createBiweeklyAgenda() will be triggered automatically if set up with a time-based daily trigger. It will check if the current week number is even, and if so, it will create a new agenda document for the "SMD Large Language Model WG."
// 3. To test the agenda creation manually, use the test function testEvenWeekAgendaCreation(). This function simulates an even week by setting the weekNumber variable to an even value (e.g., 32). When calling this test function, it will create an agenda document for the simulated even week.
// 4. When deploying the script as a standalone app, make sure to set up a time-based daily trigger for the primary function createBiweeklyAgenda(). This trigger will automatically create the agenda at the appropriate time, provided that it's an even week.
// 5. The script will create a new agenda document, customize its title with the appropriate meeting date (the following Wednesday), and replace the placeholder text "{{Day, Month Date, YYYY}}" within the document's body with the calculated meeting date in "Day, Month Date, YYYY" format.
// 6. The agenda creation process is automated and efficient, saving time and effort for the "SMD Large Language Model WG" by eliminating manual document generation and date calculation.

//////////////////////////////////////////////////////

// Global constant for Agenda Template ID and Agenda Destination Folder ID
var AGENDA_TEMPLATE_ID = "xxxxxxxxxxxxxxxx"; // Template Agenda for SMD {{short date}}
var AGENDA_FOLDER_ID = "xxxxxxxxxxxxxxxxxxxxxx"; // SMD Large Language Model WG Google Drive folder

// Helper Function: Get the next meeting date (Wednesday)
function getNextWednesday() {
  var currentDate = new Date();
  var daysUntilNextWednesday = (3 - currentDate.getDay() + 7) % 7;
  return new Date(currentDate.getTime() + (daysUntilNextWednesday * 24 * 60 * 60 * 1000));
}

// Helper Function: Format the meeting date as "M.d.yy"
function formatMeetingDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "M.d.yy");
}

// Helper Function: Replace a placeholder with a formatted date
function replaceWithFormattedDate(documentBody, placeholderText, currentDate) {
  var dateForInternal = new Date(currentDate);
  dateForInternal.setDate(dateForInternal.getDate());
  var formattedDate = Utilities.formatDate(dateForInternal, Session.getScriptTimeZone(), "EEEE, MMMM dd, yyyy");
  documentBody.replaceText(placeholderText, formattedDate);
}

// Helper Function: Create a copy of a template document with the desired title and destination folder
function createDocumentCopy(templateId, title, destinationFolderId) {
  var newFile = DriveApp.getFileById(templateId).makeCopy(title, DriveApp.getFolderById(destinationFolderId));
  return DocumentApp.openById(newFile.getId());
}

// Helper Function: Get the week number of a date
function getWeekNumber(date) {
  var onejan = new Date(date.getFullYear(), 0, 1);
  return Math.ceil(((date - onejan) / 86400000 + onejan.getDay() + 1) / 7);
}

// Primary function to create the agenda
function createBiweeklyAgenda() {
  // Check if the current week number is even (bi-weekly)
  var today = new Date();
  var weekNumber = getWeekNumber(today);
  if (weekNumber % 2 === 0) {
    
    // Create a new agenda document
    var newDocument = createDocumentCopy(AGENDA_TEMPLATE_ID, "Temporary Title", AGENDA_FOLDER_ID);
    var newDocumentId = newDocument.getId();

    // Get the meeting date for the following Wednesday
    var nextWednesday = getNextWednesday();

    // Customize the title of the agenda
    var agendaTitle = "Agenda for SMD LLM " + formatMeetingDate(nextWednesday);
    DriveApp.getFileById(newDocumentId).setName(agendaTitle);

    // Replace placeholder with the calculated meeting date in "Day, Month Date, YYYY" format
    var document = DocumentApp.openById(newDocumentId);
    var body = document.getBody();
    replaceWithFormattedDate(body, "{{Day, Month Date, YYYY}}", nextWednesday);
  }
}

// Test function for creating an agenda during an even week
function testEvenWeekAgendaCreation() {
  // Set weekNumber to an even value (e.g., 32) for testing
  var weekNumber = 32; // Set this to an even value to simulate an even week

  // Override the getWeekNumber function to return the test week number
  var originalGetWeekNumber = getWeekNumber;
  getWeekNumber = function() {
    return weekNumber;
  };

  // Call the createBiweeklyAgenda function to test agenda creation
  createBiweeklyAgenda();

  // Reset the getWeekNumber function to its original implementation
  getWeekNumber = originalGetWeekNumber;
}
