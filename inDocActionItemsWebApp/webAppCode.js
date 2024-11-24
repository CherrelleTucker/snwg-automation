function doGet() {
    // Create HTML output from the HTML file named 'Page'
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Page')
        .setTitle('Action Item Collector')
        .setFaviconUrl('https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/08d5035760893ed829b6e3ac0ed80404260743b6/action_favicon.png');
    return htmlOutput;
}

function processDocumentId(input) {
    try {
        // Extract document ID from URL if necessary
        let documentId = input;
        if (input.startsWith('https://')) {
            const match = input.match(/\/d\/([\w-]+)/);
            if (!match) {
                return 'Error: Invalid URL format';
            }
            documentId = match[1];
        }
        
        try {
            // Call the main processing function
            const result = processDocument(documentId);
            return result || 'Success! Action items collected.';
        } catch (error) {
            if (error.message.includes('No action items found')) {
                return 'No action items found. Please review the document for proper formatting of the keyword ("action: ") or action occurrences.';
            }
            return 'Error processing document: ' + error.message;
        }
        
    } catch (error) {
        Logger.log('Error in processDocumentId: ' + error.message);
        return 'Error: ' + error.message;
    }
}
