/**
 * Entry point for the web app to serve the HTML page.
 * @return {HtmlOutput} - The HTML page for the web app.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Action Item Collector')
      .setFaviconUrl('https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/08d5035760893ed829b6e3ac0ed80404260743b6/action_favicon.png');
}

/**
 * Main function to process the document from the provided URL.
 * @param {string} docUrl - The URL of the Google Doc to process.
 * @returns {Object} - Result object containing success status and message.
 */
function processDocument(docUrl) {
  try {
    if (!docUrl) {
      throw new Error('No document URL provided.');
    }

    Logger.log('Received docUrl: ' + docUrl);  // Log received URL

    const docId = extractDocId(docUrl);
    Logger.log('Extracted docId: ' + docId);  // Log extracted document ID

    const docContent = getDocumentContent(docId);
    Logger.log('Document content retrieved. Length: ' + docContent.length);  // Log the length of document content

    const docTitle = DocumentApp.openById(docId).getName();  // Get the document title
    Logger.log('Document title: ' + docTitle);  // Log the document title

    const results = parseDocument(docContent);  // Parse the document for keyphrases
    Logger.log('Parsed document results: ' + JSON.stringify(results));  // Log parsed results

    const githubResults = updateGitHubIssues(results, docUrl, docTitle);  // Update GitHub issues based on parsed results
    Logger.log('GitHub update results: ' + JSON.stringify(githubResults));  // Log GitHub results

    return { success: true, message: githubResults };
  } catch (error) {
    Logger.log('Error in processDocument: ' + error.message);  // Log error message
    return { success: false, message: 'Error: ' + error.message };
  }
}

/**
 * Helper function to extract the Google Doc ID from the URL.
 * @param {string} docUrl - The URL of the Google Doc.
 * @returns {string} - The extracted document ID.
 */
function extractDocId(docUrl) {
  const regex = /[-\w]{25,}/;
  const matches = docUrl.match(regex);
  if (!matches) {
    throw new Error('Invalid Google Doc URL');
  }
  return matches[0];
}

/**
 * Function to retrieve the content of a Google Doc.
 * @param {string} docId - The ID of the Google Doc.
 * @returns {string} - The text content of the document.
 */
function getDocumentContent(docId) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  return body.getText();
}

/**
 * Function to parse the document for specific keywords and relevant info.
 * Iterates through the document content looking for "Update issue" and "New issue".
 * @param {string} docContent - The text content of the document.
 * @returns {Object} - Parsed results with update and new issues.
 */
function parseDocument(docContent) {
  const lines = docContent.split('\n');
  const results = { updateIssues: [], newIssues: [] };
  const updateIssueRegex = /update\s*issue[:\-]*\s*/i;  // Regex to match "Update issue" in various formats
  const newIssueRegex = /new\s*issue[:\-]*\s*/i;  // Regex to match "New issue" in various formats

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    
    // Handle "Update issue"
    if (updateIssueRegex.test(line)) {
      Logger.log('Found "Update issue" keyphrase at line ' + i);
      let issueLinkLine = line.replace(updateIssueRegex, '');  // Remove the keyphrase to isolate potential link
      if (!issueLinkLine.startsWith('https://github.com')) {
        i++;
        if (i < lines.length) {
          issueLinkLine = lines[i].trim();
        }
      }

      Logger.log('Checking line for GitHub link: ' + issueLinkLine);
      const issueNumber = extractIssueNumberFromLine(issueLinkLine);
      if (issueNumber) {
        const comment = line;
        results.updateIssues.push({ issueNumber, comment });
        continue;
      }
    }

    // Handle "New issue"
    if (newIssueRegex.test(line)) {
      Logger.log('Found "New issue" keyphrase at line ' + i);
      let issueDetails = line.replace(newIssueRegex, '').trim();  // Remove the keyphrase to isolate potential title and description

      if (issueDetails === "" && i + 1 < lines.length) {
        issueDetails = lines[++i].trim();  // Take the next line if the current line is empty after keyphrase
      }

      if (issueDetails !== "") {
        const issueParts = issueDetails.split(':');  // Split title and description by ":"
        const title = issueParts[0].trim();
        const description = issueParts.length > 1 ? issueParts[1].trim() : '';
        
        if (title) {
          results.newIssues.push({ title, description });
        }
      }
    }
  }

  return results;
}

/**
 * Helper function to extract issue number from a line containing a GitHub hyperlink.
 * @param {string} line - The line containing the GitHub hyperlink.
 * @returns {string} - The extracted issue number.
 */
function extractIssueNumberFromLine(line) {
  const regex = /https:\/\/github\.com\/\w+\/\w+\/issues\/(\d+)/;  // Regex to match GitHub issue URL and extract number
  const matches = line.match(regex);
  if (matches && matches.length > 1) {
    Logger.log('Extracted issue number: ' + matches[1]);
    return matches[1];
  } else {
    Logger.log('No valid GitHub issue link found in line: ' + line);
    return null;
  }
}

/**
 * Function to update GitHub issues based on parsed document results.
 * @param {Object} results - The parsed results from the document.
 * @param {string} docUrl - The URL of the Google Doc.
 * @param {string} docTitle - The title of the Google Doc.
 * @returns {Array} - Messages indicating the actions performed on GitHub.
 */
function updateGitHubIssues(results, docUrl, docTitle) {
  const githubToken = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');  // Securely get the token
  const repo = 'CherrelleTucker/TestAssessment';
  const apiUrl = `https://api.github.com/repos/${repo}`;

  const headers = {
    'Authorization': `token ${githubToken}`,
    'Accept': 'application/vnd.github.v3+json'
  };

  let messages = [];

  // Process update issues
  results.updateIssues.forEach(issue => {
    const issueUrl = `${apiUrl}/issues/${issue.issueNumber}`;
    const commentsUrl = `${issueUrl}/comments`;
    const issueLink = `https://github.com/${repo}/issues/${issue.issueNumber}`;

    try {
      // Check for existing comments
      const existingComments = UrlFetchApp.fetch(commentsUrl, { headers }).getContentText();
      if (existingComments.includes(issue.comment)) {
        messages.push(`Content already exists in issue <a href="${issueLink}" target="_blank">#${issue.issueNumber}</a>`);
      } else {
        // Add comment to the issue with link to the source document
        const payload = JSON.stringify({ body: `${issue.comment}\n[${docTitle}](${docUrl})` });
        UrlFetchApp.fetch(commentsUrl, { method: 'POST', headers, payload });
        messages.push(`Updated issue <a href="${issueLink}" target="_blank">#${issue.issueNumber}</a> with a new comment.`);
      }
    } catch (e) {
      Logger.log('Error updating GitHub issue #' + issue.issueNumber + ': ' + e.message);
    }
  });

  // Process new issues
  results.newIssues.forEach(issue => {
    const issuesUrl = `${apiUrl}/issues`;
    try {
      const payload = JSON.stringify({ title: issue.title, body: `${issue.description}\n[${docTitle}](${docUrl})` });
      const response = UrlFetchApp.fetch(issuesUrl, { method: 'POST', headers, payload });
      const newIssue = JSON.parse(response.getContentText());
      const newIssueLink = newIssue.html_url;
      messages.push(`Created new issue: <a href="${newIssueLink}" target="_blank">${issue.title}</a>`);
    } catch (e) {
      Logger.log('Error creating new GitHub issue titled "' + issue.title + '": ' + e.message);
    }
  });

  return messages;
}

/**
 * Test function to manually run and check the script with logging.
 */
function testProcessDocument() {
  const testUrl = 'https://docs.google.com/document/d/1FE1d6oPtxj0Leo4Hi-Cz5IC9CIkua8m4hqDX0nTsNDg/edit';
  const result = processDocument(testUrl);
  Logger.log('Test result: ' + JSON.stringify(result));
}
