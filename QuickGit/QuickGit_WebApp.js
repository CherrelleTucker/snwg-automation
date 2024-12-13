/**
 * WEBAPP.GS - Web Application Interface
 * Contains web app specific functions and request handling
 * Acts as bridge between frontend and core functionality
 */

// =============================================================================
// WEB APP INITIALIZATION
// =============================================================================

/**
 * Handles web app requests and OAuth flow
 * Entry point for the web application
 * @param {Object} e - Event object from web app
 * @returns {HtmlOutput} The HTML page
 */
function doGet(e) {
  // Handle OAuth callback
  if (e.parameter.code) {
    const result = handleOAuthCallback(e.parameter.code, e.parameter.state);
    if (!result.success) {
      return HtmlService.createHtmlOutput('<h3>Authentication Failed</h3><p>Error: ' + result.error + '</p>');
    }
    return HtmlService.createHtmlOutput('<h3>Authentication Successful!</h3><script>setTimeout(function() { window.close(); }, 2000);</script>');
  }
  
  // Check if user is authenticated
  const userProperties = PropertiesService.getUserProperties();
  const isAuthenticated = Boolean(userProperties.getProperty('github_access_token'));
  
 // Pass `isAuthenticated` to the HTML template
  const template = HtmlService.createTemplateFromFile('index');
  template.isAuthenticated = isAuthenticated;

  // Log authentication state for debugging
  console.log('User authentication state:', isAuthenticated);
  
  return template.evaluate()
    .setTitle('QuickGit')
    .setFaviconUrl('https://raw.githubusercontent.com/CherrelleTucker/snwg-automation/348c5c6f451c951c6a4558fec69e75809886aa45/DocToGitHubTransferToolIcon.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* =============================================================================
DOCUMENT PROCESSING ENDPOINTS
=============================================================================
*/

/**
 * Processes document from URL
 * Called by frontend when processing a new document
 * @param {string} docUrl - URL of document to process
 * @returns {Object} Processing results or error
 */
function processDocument(docUrl) {
  try {
    const doc = DocumentApp.openByUrl(docUrl);
    if (!doc) {
      return { error: 'Invalid document URL or no access to document' };
    }

    // Get document metadata
    const docMetadata = {
      title: doc.getName(),
      url: docUrl
    };

    // First use the working parser to get the basic structure
    const content = doc.getBody().getText();
    const basicParsed = parseDocContent(content, docMetadata);
    
    // Now use parseDocContentWithFormatting to get formatted versions
    const formattedParsed = parseDocContentWithFormatting(doc.getBody(), docMetadata);
    
    // Merge the results, preserving the working issue recognition from basicParsed
    // but using the formatted content from formattedParsed
    const mergedResults = {
      newIssues: basicParsed.newIssues.map((issue, index) => ({
        ...issue,
        body: formattedParsed.newIssues[index]?.body || issue.body
      })),
      updateIssues: basicParsed.updateIssues.map((issue, index) => ({
        ...issue,
        comment: formattedParsed.updateIssues[index]?.body || issue.comment
      }))
    };

    // Validate parsed content
    if (!mergedResults.newIssues.length && !mergedResults.updateIssues.length) {
      return { error: 'No new issues or updates found in document' };
    }

    return mergedResults;
  } catch (error) {
    console.error('Error processing document:', error);
    return { error: error.message || 'Failed to process document' };
  }
}

/* =============================================================================
ISSUE MANAGEMENT ENDPOINTS
=============================================================================
*/


/**
 * Processes issues based on UI data
 * @param {Object} data - Data object containing issue information
 * @param {Array} data.newIssues - Array of new issues to create
 * @param {Array} data.updateIssues - Array of issues to update
 * @returns {Object} Results of processing containing created, updated, and error information
 * @throws {Error} If data validation fails
 */
function processIssues(data) {
  try {
    // Input validation
    if (!data) {
      throw new Error('No data provided for processing');
    }

    // Initialize results object
    const results = {
      created: [],
      updated: [],
      errors: [],
      links: [] // Array to store URLs and metadata
    };

    // Process new issues if they exist
    if (data.newIssues && Array.isArray(data.newIssues)) {
      for (const issue of data.newIssues) {
        try {
          // Validate required fields
          if (!issue.repo || !issue.title || !issue.body) {
            throw new Error(`Missing required fields for new issue: ${issue.title}`);
          }

          // Create the issue
          const createdIssue = createIssue(issue.repo, {
            title: issue.title,
            body: issue.body,
            assignees: issue.assignee ? [issue.assignee] : []
          });

          // Store issue number and URL
          results.created.push(createdIssue.number);
          results.links.push({
              type: 'issue',
              number: createdIssue.number,
              url: createdIssue.html_url,
              repo: issue.repo,
              title: issue.title
          });

        } catch (error) {
          results.errors.push(`Failed to create issue "${issue.title}": ${error.message}`);
        }
      }
    }

    // Process update issues if they exist
    if (data.updateIssues && Array.isArray(data.updateIssues)) {
      for (const update of data.updateIssues) {
        try {
          // Validate required fields
          if (!update.repo || !update.issueNumber || !update.comment) {
            throw new Error(`Missing required fields for update issue #${update.issueNumber}`);
          }

          // add the comment
          const updatedIssue = updateIssue(update.repo, update.issueNumber, update.comment);

          // Store issue number and comment url
          results.updated.push(update.issueNumber);
          results.links.push({
            type: 'comment',
            number: update.issueNumber,
            url: updatedIssue.html_url,
            repo: update.repo
          });

        } catch (error) {
          results.errors.push(`Failed to update issue #${update.issueNumber}: ${error.message}`);
        }
      }
    }

    return results;

  } catch (error) {
    throw new Error(`Failed to process issues: ${error.message}`);
  }
}

/**
 * Validates an individual issue object
 * @param {Object} issue - Issue object to validate
 * @param {string} type - Type of issue ('new' or 'update')
 * @returns {boolean} True if valid, false otherwise
 */
function validateIssue(issue, type) {
  if (type === 'new') {
    return issue && issue.repo && issue.title && issue.body;
  } else if (type === 'update') {
    return issue && issue.repo && issue.issueNumber && issue.comment;
  }
  return false;
}

/* =============================================================================
GITHUB DATA ENDPOINTS
=============================================================================
*/

/**
 * Gets repositories for frontend display
 * @returns {Array} List of repositories
 */
function getRepos() {
  try {
    Logger.log('WebApp: Fetching repositories for frontend');
    const allRepos = fetchOrgRepos();
    Logger.log(`WebApp: Returning ${allRepos.length} repositories to frontend`);
    return allRepos.map(repo => ({
      name: repo.name
    }));
  } catch (error) {
    Logger.log('WebApp: Error in getRepos: ' + error.message);
    throw new Error('Failed to fetch repositories: ' + error.message);
  }
}

/**
 * Gets the list of organization members for the frontend
 * @returns {Array} List of members
 */
function getMembers() {
  try {
    return fetchOrgMembers();
  } catch (error) {
    Logger.log('Error in getMembers: ' + error.message);
    throw new Error('Failed to fetch members: ' + error.message);
  }
}

/**
 * Fetches open issues for a repository with recursion protection
 * @param {string} repoName - Repository name
 * @returns {Array} List of open issues
 */
function fetchRepoIssues(repoName) {
  // Add guard for recursive calls
  if (!repoName) {
    throw new Error('Repository name is required');
  }

  try {
    const client = createGitHubClient();
    const url = `${client.baseUrl}/repos/${CONFIG.ORG_NAME}/${repoName}/issues?state=open&per_page=100`;
    
    // Make a single API call with proper error handling
    const response = UrlFetchApp.fetch(url, {
      headers: client.headers,
      muteHttpExceptions: true
    });

    // Check response code first
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      console.error('Error response:', response.getContentText());
      throw new Error(`GitHub API returned status ${responseCode}`);
    }

    // Parse response once
    const issues = JSON.parse(response.getContentText());
    
    // Return simplified issue objects
    return issues.map(issue => ({
      number: issue.number,
      title: issue.title,
      state: issue.state,
      url: issue.html_url
    }));

  } catch (error) {
    console.error('Error in fetchRepoIssues:', error);
    throw new Error(`Failed to fetch issues: ${error.message}`);
  }
}

