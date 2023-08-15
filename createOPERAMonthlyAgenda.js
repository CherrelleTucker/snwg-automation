// Purpose: 
// Automate the creation of a new monthly agenda for the "SNWG/OPERA Monthly Tag-up" in the designated "SNWG-OPERA tag-up" folder on Google Drive. It uses a specified template with placeholders to generate the agenda document, populates it with the relevant information, such as the date of the meeting and links to the previous month's agenda, and then places the new agenda in the designated folder. The script ensures that the agenda is consistently formatted and up-to-date, simplifying the process for users and maintaining organization within the monthly meeting records.

// To Note:
// This script is developed as a Google Apps Script standalone script. It is designed to operate independently and does not require any external application or service to function. It is a self-contained piece of code with a time-based monthly trigger.

// To Use: 
// 1. Prepare Google Drive and Template: Ensure you have the necessary access and permissions to create files in Google Drive. Prepare a template for the monthly agenda in Google Docs and make a note of its Template ID.

// 2. Open the Script Editor: Open the Google Apps Script editor by clicking on "Extensions" in the top menu and selecting "Apps Script."

// 3. Copy and Paste the Script: Copy the entire script provided above and paste it into the Google Apps Script editor.

// 4. Configure Global Variables: Modify the global variables at the beginning of the script to match your Template ID and the folder ID of the "Monthly Project Status Update" folder where you want to store the agendas.

// 5. Save the Script: Save the script by clicking on the floppy disk icon or pressing "Ctrl + S" (Windows) or "Cmd + S" (Mac).

// 6. Run the Script: To generate a new agenda, click on the play button ▶️ in the Google Apps Script editor or manually run the `createNewMonthlyAgenda()` function.

// 7. Agenda Generation: The script will create a new agenda by making a copy of the provided template, populating it with the relevant details, and placing it in the specified folder. The agenda's name will be in the format "YYYY-MM-DD SNWG MO Monthly Project Update Meeting" based on the fourth Monday of the current month.

// 8. Hyperlink to Previous Agenda: The script will automatically search for the most recent previous month's agenda in the folder (excluding the newly created one) and replace the placeholder "{{link to last SNWG/NASA monthly}}" with a hyperlink to that agenda.

// 9. Schedule Automation (Optional): If you want to generate agendas automatically, set up a trigger to run the script at a specific interval (e.g., monthly). To do this, click on the clock icon in the Google Apps Script editor, add a new trigger, and set it to run the `createNewMonthlyAgenda()` function on the desired schedule.

// 10. Ensure Permission: When running the script for the first time, it may ask for permission to access your Google Drive. Click "Continue" and grant the necessary permissions.

// Please review the script, understand its functionality, and make any required adjustments before running it. Additionally, ensure that the script has access to the necessary folders and files in your Google Drive.

////////////////////////////////////////////

// Global variables
var templateId = "1vr50ZU0skOM7_UOYaFKQtfDkFlkE5O7KzlItsdP4BOM"; // Template: YYYY-MM-DD SNWG/OPERA Tag-Up
var operaMonthlyFolderId = "1M3EMWLCxhkqcPKLmH7grDu2zhSEuOvmc"; // SNWG-OPERA Tag-up Folder>Meeting-Notes>2023
var snwgMonthlyFolder = '1HPjhc2LADvS9j3W_K3riq4RQPBngfqGY'; // SNWG Monthly Project Updates

// helper function to return the date of the second Thursday of a given month, or the current month if not specified.
function getsecondThursday(date) {
     date = date || new Date();
    var timezoneOffset = date.getTimezoneOffset() * 60000;
    date = new Date(date.getTime() - timezoneOffset);
  if (isNaN(date.getTime())) { // check if date is valid
      throw new Error("Invalid date object");
  }
  // Find the second Thursday of the current month
  date.setDate(1);
  var day = date.getDay();
  var diff = (day <= 4 ? 4 : 11) - day;  // Adjusted this line to get the first Thursday
  var firstThursday = new Date(date.getTime() + diff * 24 * 60 * 60 * 1000);

  // Add one week ( 7 days to get the second Thursday)
  var secondThursday = new Date(firstThursday.getTime() + 7 * 24 * 60 * 60 * 1000);

  return secondThursday;
}

// helper function to Return the date of the second Thursday of the next month, or the next month after the specified date.
function getNextsecondThursday(date) {
     date = date || new Date();
    var timezoneOffset = date.getTimezoneOffset() * 60000;
    date = new Date(date.getTime() - timezoneOffset);
  if (isNaN(date.getTime())) { // check if date is valid
      throw new Error("Invalid date object");
  }
  
  // Set the date to the first day of the next month
  date.setMonth(date.getMonth() + 1);
  date.setDate(1);
  
  var day = date.getDay();
  var diff = (day <= 4 ? 4 : 11) - day;  // Adjusted this line to get the first Thursday
  var firstThursday = new Date(date.getTime() + diff * 24 * 60 * 60 * 1000);

  // Add one week (7 days to get the second Thursday)
  var secondThursday = new Date(firstThursday.getTime() + 7 * 24 * 60 * 60 * 1000);

  return secondThursday;
}

// helper function to Replace occurrences of searchText in the bodyElement with a hyperlink having the URL linkUrl and display text linkText.
function replaceWithHyperlink(bodyElement, searchText, linkUrl, linkText) {
  var paragraphs = bodyElement.getParagraphs();
  for (var i in paragraphs) {
      var text = paragraphs[i].editAsText();
      var foundOffset = text.findText(searchText);
      if (foundOffset !== null) {
          var start = foundOffset.getStartOffset();
          var end = foundOffset.getEndOffsetInclusive();
          text.insertText(start, linkText).setLinkUrl(start, start + linkText.length - 1, linkUrl);
          text.deleteText(start + linkText.length, end + linkText.length);
      }
  }
}

//helper function to Find the most recent file in the specified folder, excluding the one with excludeId, and returns its URL and file name, replacing placeholderText in the process.
function getMostRecentFileLink(folderId, excludeId, placeholderText) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var newestFile = null;
  var newestFileName = '';
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      var fileId = file.getId();

      // Skip files whose title date is in the future
      if (fileName.localeCompare(todayStr) > 0) {
          continue;
      }

      if (fileName.localeCompare(newestFileName) > 0 && fileId !== excludeId && !fileName.includes("Template")) {
          newestFile = file;
          newestFileName = fileName;
      }
  }

  if (newestFile !== null) {
      var placeholderLink = "{{Link to last SNWG/NASA monthly}}";
      if (placeholderText === placeholderLink && newestFile.getId() === excludeId) {
          return ["", ""];
      } else {
          return [newestFile.getUrl(), newestFileName];
      }
  } else {
      return ["", ""];
  }
}

// primary function to generate, populate, and place new agenda
function createNewMonthlyAgenda() {
  var operaMonthlyFolder = DriveApp.getFolderById(operaMonthlyFolderId);
  var snwgFolder = DriveApp.getFolderById(snwgMonthlyFolder); // Added this line
  var templateFile = DriveApp.getFileById(templateId);
  var newDocument = templateFile.makeCopy(templateFile.getName(), operaMonthlyFolder);
  var newDocumentId = newDocument.getId();
  var document = DocumentApp.openById(newDocumentId);
  var documentBody = document.getBody();
  
  var today = new Date();
  var secondThursdayThisMonth = getsecondThursday(today);
  
  // Check if today's date has surpassed the second Thursday of the current month
  var monthlyDate;
  if (today > secondThursdayThisMonth) {
    monthlyDate = getNextsecondThursday(today);
  } else {
    monthlyDate = secondThursdayThisMonth;
  }
  
  var newDocumentName = Utilities.formatDate(monthlyDate, Session.getScriptTimeZone(), "yyyy-MM-dd") + " SNWG/OPERA Tag-Up";
  document.setName(newDocumentName);
  
  var formattedDate = Utilities.formatDate(monthlyDate, Session.getScriptTimeZone(), "EEEE, MMMM dd, yyyy");
  documentBody.replaceText("{{Monthly Day}}", formattedDate);
  
  var [operaMonthlyLink, operaNewestFileName] = getMostRecentFileLink(operaMonthlyFolderId, newDocumentId, "{{Link to Previous Agenda}}");
  replaceWithHyperlink(documentBody, "{{Link to Previous Agenda}}", operaMonthlyLink, operaNewestFileName);
  
  // New section for the SNWG Monthly link replacement
  var [snwgMonthlyLink, snwgNewestFileName] = getMostRecentFileLink(snwgMonthlyFolder, newDocumentId, "{{Link to last SNWG/NASA monthly}}");
  replaceWithHyperlink(documentBody, "{{Link to last SNWG/NASA monthly}}", snwgMonthlyLink, snwgNewestFileName);
  
  documentBody.replaceText("{{Monthly Date}}", formattedDate);  
  
  var nextMonthlyDate = getNextsecondThursday(monthlyDate);
  var formattedNextDate = Utilities.formatDate(nextMonthlyDate, Session.getScriptTimeZone(), "EEEE, MMMM dd, yyyy");
  documentBody.replaceText("{{next monthly meeting}}", formattedNextDate);
}
