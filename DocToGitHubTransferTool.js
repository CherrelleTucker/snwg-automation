/**
 * Entry point for the web app to serve the HTML page.
 * @return {HtmlOutput} - The HTML page for the web app.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Doc to GH Transfer Tool')
      .setFaviconUrl('https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/38b13230c6f746ef09a0a583f6746d706cad6233/DocToGitHubTransferToolIcon.png');
}

function processDocument(docUrl) {
  try {
    if (!docUrl) {
      throw new Error('No document URL provided.');
    }

    Logger.log('Received docUrl: ' + docUrl);

    const docId = extractDocId(docUrl);
    Logger.log('Extracted docId: ' + docId);

    const docContent = getDocumentContent(docId);
    Logger.log('Document content retrieved. Length: ' + docContent.length);

    const docTitle = DocumentApp.openById(docId).getName();
    Logger.log('Document title: ' + docTitle);

    const results = parseDocument(docContent);
    Logger.log('Parsed document results: ' + JSON.stringify(results));

    const githubResults = updateGitHubIssues(results, docUrl, docTitle);
    Logger.log('GitHub update results: ' + JSON.stringify(githubResults));

    return { success: true, message: githubResults };
  } catch (error) {
    Logger.log('Error in processDocument: ' + error.message);
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
 * Function to retrieve the content of a Google Doc and convert it into Markdown format.
 * @param {string} docId - The ID of the Google Doc.
 * @returns {string} - The text content of the document in Markdown format.
 */
function getDocumentContent(docId) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();
  return convertToMarkdown(body);
}

/**
 * Function to convert Google Doc elements to Markdown format.
 * @param {Body} body - The body of the Google Doc.
 * @returns {string} - The content of the body in Markdown format.
 */
function convertToMarkdown(body) {
  let markdownContent = '';
  const numChildren = body.getNumChildren();
  
  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = element.asParagraph();
      
      if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
        // Handle headings
        const heading = getHeadingMarkdown(paragraph.getHeading());
        markdownContent += heading + paragraph.getText() + '\n\n';
      } else {
        // Handle normal text paragraphs
        markdownContent += paragraph.getText() + '\n\n';
      }
    } else if (element.getType() === DocumentApp.ElementType.LIST_ITEM) {
      // Handle list items (bulleted or numbered)
      const listItem = element.asListItem();
      markdownContent += '- ' + listItem.getText() + '\n';
    }
  }
  
  return markdownContent.trim();  // Remove trailing spaces or newlines
}

/**
 * Helper function to get the Markdown representation of a heading.
 * @param {DocumentApp.ParagraphHeading} heading - The Google Docs heading type.
 * @returns {string} - The Markdown heading prefix.
 */
function getHeadingMarkdown(heading) {
  switch (heading) {
    case DocumentApp.ParagraphHeading.HEADING1:
      return '# ';
    case DocumentApp.ParagraphHeading.HEADING2:
      return '## ';
    case DocumentApp.ParagraphHeading.HEADING3:
      return '### ';
    default:
      return '';
  }
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
  const updateIssueRegex = /update\s*issue[:\-]*\s*/i;
  const newIssueRegex = /new\s*issue[:\-]*\s*/i;

  let currentComment = '';
  let currentTitle = '';
  let isUpdateIssue = false;
  let isNewIssue = false;
  let issueNumber = '';

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    Logger.log(`Processing line ${i}: ${line}`);

    // Detect "Update issue"
    if (updateIssueRegex.test(line)) {
      finalizeCurrentComment(results, currentTitle, currentComment, issueNumber, isUpdateIssue, isNewIssue);

      Logger.log('Found "Update issue" keyphrase at line ' + i);
      issueNumber = extractIssueNumberFromLine(line.replace(updateIssueRegex, '').trim());
      isUpdateIssue = true;
      isNewIssue = false;
      currentComment = '';  // Reset comment collection for updates
      continue;
    }

    // Detect "New issue"
    if (newIssueRegex.test(line)) {
      finalizeCurrentComment(results, currentTitle, currentComment, issueNumber, isUpdateIssue, isNewIssue);

      Logger.log('Found "New issue" keyphrase at line ' + i);
      currentTitle = line.replace(newIssueRegex, '').trim().split(':')[0];
      currentComment = '';  // Start collecting content for the new issue body
      isUpdateIssue = false;
      isNewIssue = true;
      continue;
    }

    // Convert bullets to Markdown bullets if detected and append to comment
    if (line.startsWith('•') || line.startsWith('-')) {
      currentComment += `- ${line.replace(/^[•-]\s*/, '')}\n`;
    } else if (line !== '') {
      currentComment += `${line}\n`;
    }
  }

  finalizeCurrentComment(results, currentTitle, currentComment, issueNumber, isUpdateIssue, isNewIssue);

  Logger.log('Final parsed document results: ' + JSON.stringify(results));
  return results;
}

/**
 * Helper function to finalize the current comment or issue and add it to the results.
 * @param {Object} results - The results object where issues or comments will be stored.
 * @param {string} currentTitle - The current issue title (for new issues).
 * @param {string} currentComment - The accumulated comment or description.
 * @param {string} issueNumber - The issue number (for update issues).
 * @param {boolean} isUpdateIssue - Whether it's an update issue.
 * @param {boolean} isNewIssue - Whether it's a new issue.
 */
function finalizeCurrentComment(results, currentTitle, currentComment, issueNumber, isUpdateIssue, isNewIssue) {
  if (isUpdateIssue && issueNumber && currentComment.trim()) {
    Logger.log(`Finalizing update issue #${issueNumber} with comment: ${currentComment.trim()}`);
    results.updateIssues.push({ issueNumber, comment: currentComment.trim() });
  } else if (isNewIssue && currentTitle && currentComment.trim()) {
    Logger.log(`Finalizing new issue with title "${currentTitle}" and description: ${currentComment.trim()}`);
    results.newIssues.push({ title: currentTitle, description: currentComment.trim() });
  }
}

function updateGitHubIssues(results, docUrl, docTitle) {
  const githubToken = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const repo = 'CherrelleTucker/TestAssessment';
  const apiUrl = `https://api.github.com/repos/${repo}`;
  
  const headers = {
    'Authorization': `token ${githubToken}`,
    'Accept': 'application/vnd.github.v3+json'
  };
  
  let messages = [];
  
  const commentMarker = `[Generated by Doc to GH Transfer Tool](https://script.google.com/a/macros/nasa.gov/s/AKfycbwqpBs1GyGnBHd34MOlRaizKdisqmqbgNMJC669zgGwSybZ8zkn-nmtah7F1WN7r0at/exec)`; // Marker to help detect duplicates
  
  const existingIssues = getExistingIssues(repo, headers);
  
  results.updateIssues.forEach(issue => {
    const issueUrl = `${apiUrl}/issues/${issue.issueNumber}`;
    const commentsUrl = `${issueUrl}/comments`;
    const issueLink = `https://github.com/${repo}/issues/${issue.issueNumber}`;
    
    try {
      const existingCommentsResponse = UrlFetchApp.fetch(commentsUrl, { headers });
      const existingComments = JSON.parse(existingCommentsResponse.getContentText());

      Logger.log(`Checking for existing comments on issue #${issue.issueNumber}`);
      
      let commentExists = false;
      
      // Updated formatting logic for comments
      const newCommentBody = `${issue.comment.trim()}\n\nSource file: [${docTitle}](${docUrl})\n\n${commentMarker}`;
      
      existingComments.forEach(existingComment => {
        const normalizedExistingComment = normalizeString(existingComment.body);
        
        if (normalizedExistingComment.includes(commentMarker) && normalizedExistingComment.includes(issue.comment.trim())) {
          commentExists = true;
          Logger.log(`Found matching comment in issue #${issue.issueNumber}.`);
        }
      });
      
      if (commentExists) {
        messages.push(`Content already exists in issue <a href="${issueLink}" target="_blank">#${issue.issueNumber}</a>`);
      } else {
        // Add new comment with proper formatting
        const payload = JSON.stringify({
          body: newCommentBody
        });
        UrlFetchApp.fetch(commentsUrl, { method: 'POST', headers, payload });
        messages.push(`Updated issue <a href="${issueLink}" target="_blank">#${issue.issueNumber}</a> with a new comment.`);
      }
    } catch (e) {
      Logger.log('Error updating GitHub issue #' + issue.issueNumber + ': ' + e.message);
    }
  });

  // Similar formatting logic for new issues
  results.newIssues.forEach(issue => {
    const issuesUrl = `${apiUrl}/issues`;
    
    const existingIssue = existingIssues.find(existingIssue => {
      return existingIssue.title === issue.title && existingIssue.body.includes(issue.description.trim());
    });
    
    if (existingIssue) {
      const issueLink = `https://github.com/${repo}/issues/${existingIssue.number}`;
      messages.push(`Issue with title "<a href="${issueLink}" target="_blank">${issue.title}</a>" already exists.`);
    } else {
      try {
        const payload = JSON.stringify({
          title: issue.title,
          body: `${issue.description.trim()}\n\nSource file: [${docTitle}](${docUrl})\n\n${commentMarker}`
        });
        const response = UrlFetchApp.fetch(issuesUrl, { method: 'POST', headers, payload });
        const newIssue = JSON.parse(response.getContentText());
        const newIssueLink = newIssue.html_url;
        messages.push(`Created new issue: <a href="${newIssueLink}" target="_blank">${issue.title}</a>`);
      } catch (e) {
        Logger.log('Error creating new GitHub issue titled "' + issue.title + '": ' + e.message);
      }
    }
  });
  
  return messages;
}

/**
 * Helper function to normalize comment content for comparison.
 * Removes extra spaces, newlines, and ensures consistent formatting.
 * @param {string} str - The string to normalize.
 * @returns {string} - Normalized string.
 */
function normalizeString(str) {
  return str.replace(/\s+/g, ' ').trim(); // Replace multiple spaces/newlines with single space and trim
}

/**
 * Helper function to extract the first sentence of a string.
 * @param {string} str - The string to extract the first sentence from.
 * @returns {string} - The first sentence of the string.
 */
function extractFirstSentence(str) {
  const match = str.match(/^.*?[.?!](\s|$)/);
  return match ? match[0].trim() : str;
}

/**
 * Function to fetch existing issues from the GitHub repository.
 * @returns {Array} - Array of existing issue titles and numbers.
 */
function getExistingIssues(repo, headers) {
  const apiUrl = `https://api.github.com/repos/${repo}/issues?state=all`;
  const response = UrlFetchApp.fetch(apiUrl, { headers });
  const issues = JSON.parse(response.getContentText());

  return issues.map(issue => ({
    number: issue.number, // Ensure we are capturing the issue number correctly
    title: issue.title,
    body: issue.body
  }));
}

/**
 * Helper function to extract issue number from a line containing a GitHub hyperlink.
 * @param {string} line - The line containing the GitHub hyperlink.
 * @returns {string} - The extracted issue number.
 */
function extractIssueNumberFromLine(line) {
  const regex = /https:\/\/github\.com\/\w+\/\w+\/issues\/(\d+)/;
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
 * Test function to manually run and check the script with logging.
 */
function testProcessDocument() {
  const testUrl = 'https://docs.google.com/document/d/1FE1d6oPtxj0Leo4Hi-Cz5IC9CIkua8m4hqDX0nTsNDg/edit';
  const result = processDocument(testUrl);
  Logger.log('Test result: ' + JSON.stringify(result));
}
