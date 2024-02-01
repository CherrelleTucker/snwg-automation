function doGet() {
  return HtmlService.createHtmlOutputFromFile('Page');
}

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
  
  // Call existing function
  processDocument(documentId); 
  return 'Success!'; // Or an appropriate message
}