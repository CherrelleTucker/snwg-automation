// Purpose:
// to be part of a web app designed to let users input a Google Document URL or ID. 
// Once provided, the app processes the Google Document to identify and catalog action items.

// Google Apps Script function to serve as an html page. 
function doGet() {
    return HtmlService.createHtmlOutputFromFile('Page');
  }
  
// accepts an input which can either be a direct Google Document ID or a full Google Document URL and extracts the document ID, calls the processDocument() function (from code.js) to process the Google Document by searching for and cataloging action items.
  function processDocumentId(input) {
    // If the input looks like a URL, extract the ID
    var documentId = input;
    if (input.startsWith('https://')) {
      var match = input.match(/\/d\/([\w-]+)/); // Regular expression to match /d/ followed by the ID
      if (match) {
        documentId = match[1];
      } else {
        // Handle error if URL does not match expected pattern
        return 'Error: Invalid URL format'; 
      }
    }
    
    // Then, call your existing function
    processDocument(documentId); // Updated this line
    return 'Success'; // Or an appropriate message
  }
  
  