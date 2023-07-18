// Purpose: Create new monthly agenda for the SNWG Monthly Update Meeting in the Monthly meeting folder
// Done: Duplicate template agenda in 2023 folder
// Done: Document name populates with fourth Monday of the current month
// Done: Populate last Monthly status update agenda 
// Done: Monthly Date formats in Day Month, Date, Year format. 
// Done: Populates date of next meeting.
 

function getFourthMonday(date) {
  date = date || new Date();
  if (isNaN(date.getTime())) { // check if date is valid
      throw new Error("Invalid date object");
  }
  // Find the fourth Monday of the current month
  date.setDate(1);
  var day = date.getDay();
  var diff = (day === 0 ? 1 : 8) - day;
  var firstMonday = new Date(date.getTime() + diff * 24 * 60 * 60 * 1000);

  // Add three weeks (21 days to get the fourth Monday)
  var fourthMonday = new Date(firstMonday.getTime() + 21 * 24 * 60 * 60 * 1000);

  // Format the date as "YYYY-MM-DD"
  var formattedDate = Utilities.formatDate(fourthMonday, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return formattedDate;
}

function getNextFourthMonday(date) {
  date = date || new Date();
  if (isNaN(date.getTime())) { // check if date is valid
      throw new Error("Invalid date object");
  }
  
  // Set the date to the first day of the next month
  date.setMonth(date.getMonth() + 1);
  date.setDate(1);
  
  var day = date.getDay();
  var diff = 1+(day === 0 ? 1 : 8) - day;
  var firstMonday = new Date(date.getTime() + diff * 24 * 60 * 60 * 1000);

  // Add three weeks (21 days to get the fourth Monday)
  var fourthMonday = new Date(firstMonday.getTime() + 21 * 24 * 60 * 60 * 1000);

  // Format the date as "YYYY-MM-DD"
  var formattedDate = Utilities.formatDate(fourthMonday, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return formattedDate;
}

function replaceWithHyperlink(bodyElement, searchText, linkUrl, linkText) {
  // Get all the elements in the document body
  var paragraphs = bodyElement.getParagraphs();
  for (var i in paragraphs) {
      var text = paragraphs[i].editAsText();
      // Find the position of the search text
      var foundOffset = text.findText(searchText);
      if (foundOffset !== null) {
          var start = foundOffset.getStartOffset();
          var end = foundOffset.getEndOffsetInclusive();
          // Insert the hyperlink text
          text.insertText(start, linkText).setLinkUrl(start, start + linkText.length - 1, linkUrl);
          // Delete the placeholder text
          text.deleteText(start + linkText.length, end + linkText.length);
      }
  }
}

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

      // If this file is more recent than our current newest file, and is not the template or excluded file, update newestFile and newestFileName
      if (fileName.localeCompare(newestFileName) > 0 && fileId !== excludeId && !fileName.includes("Template")) {
          newestFile = file;
          newestFileName = fileName;
      }
  }

  if (newestFile !== null) {
      var placeholderLink = "{{Link to Previous Agenda}}";
      if (placeholderText === placeholderLink && newestFile.getId() === excludeId) {
          return ["", ""];
      } else {
          return [newestFile.getUrl(), newestFileName];
      }
  } else {
      return ["", ""];
  }
}

function createNewMonthlyAgenda() {
  // Set the ID of the template document
  var templateId = "1J1hij-_8Bh9ygQdt7mBeDDk-p6DsZm2njKoPrtnqBvc";

  // Make a copy of the template document
  var newDocument = DriveApp.getFileById(templateId).makeCopy();

  // Get the new document's ID
  var newDocumentId = newDocument.getId();

  // Open the new document
  var document = DocumentApp.openById(newDocumentId);

  // Set the new document's name with the meeting date (fourth Monday of the month)
  var monthlyDate = getFourthMonday(new Date());
  var newDocumentName = monthlyDate + " SNWG MO Monthly Project Update Meeting";
  document.setName(newDocumentName);

  // Replace placeholders with links and date
  var snwgMonthlyFolderId = "1HPjhc2LADvS9j3W_K3riq4RQPBngfqGY";
  var [snwgMonthlyLink, newestFileName] = getMostRecentFileLink(snwgMonthlyFolderId, newDocumentId, "{{link to last SNWG/NASA monthly}}");
  var documentBody = document.getBody();

  // Replace the placeholders with hyperlinks
  replaceWithHyperlink(documentBody, "{{link to last SNWG/NASA monthly}}", snwgMonthlyLink, newestFileName);

  // Format the date as "Day, Month Date, Year" and replace the placeholder
  var formattedDate = Utilities.formatDate(new Date(monthlyDate), Session.getScriptTimeZone(), "EEEE, MMMM dd, yyyy");
  documentBody.replaceText("{{Monthly Date}}", formattedDate);  

  // Format the next month's fourth Monday date as "Day, Month Date, Year" and replace the placeholder
  var nextMonthlyDate = getNextFourthMonday(new Date());
  var formattedNextDate = Utilities.formatDate(new Date(nextMonthlyDate), Session.getScriptTimeZone(), "EEEE, MMMM dd, yyyy");
  documentBody.replaceText("{{next monthly meeting}}", formattedNextDate);  

}
