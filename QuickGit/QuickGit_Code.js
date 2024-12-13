/**
 * CODE.GS - Core Backend Functionality
 * Contains core functionality, GitHub API interactions, and document processing logic
 * Handles all backend operations and utility functions
 */

// =============================================================================
// CONFIGURATION AND CONSTANTS
// =============================================================================
/**
 * Core configuration object for application settings
 * Contains API endpoints, organization settings, and OAuth configuration
*/

/**
 * Constants and configuration
 */
const CONFIG = {
  GITHUB_API_BASE: 'https://api.github.com',
  ORG_NAME: 'NASA-IMPACT',
  CACHE_DURATION: 21600, // 6 hours in seconds
  GITHUB_AUTH_URL: 'https://github.com/login/oauth/authorize',
  GITHUB_TOKEN_URL: 'https://github.com/login/oauth/access_token',
  OAUTH_SCOPES: ['repo', 'read:org', 'write:discussion'],
  
  get OAUTH_CLIENT_ID() {
    return PropertiesService.getScriptProperties().getProperty('GITHUB_CLIENT_ID');
  },
  get OAUTH_CLIENT_SECRET() {
    return PropertiesService.getScriptProperties().getProperty('GITHUB_CLIENT_SECRET');
  },
  get OAUTH_REDIRECT_URI() {
    return PropertiesService.getScriptProperties().getProperty('REDIRECT_URI');
  }
};

/* 
=============================================================================
OAUTH AND AUTHENTICATION
=============================================================================
*/

/**
 * Verifies OAuth is properly set up
 * @returns {boolean} True if OAuth configuration is valid
 */
function verifyOAuthConfig() {
  const clientId = CONFIG.OAUTH_CLIENT_ID;
  const clientSecret = CONFIG.OAUTH_CLIENT_SECRET;
  const redirectUri = CONFIG.OAUTH_REDIRECT_URI;  // Fetch from CONFIG to keep it consistent

  if (!clientId || !clientSecret || !redirectUri) {
    console.error('Missing OAuth configuration:', {
      hasClientId: Boolean(clientId),
      hasClientSecret: Boolean(clientSecret),
      hasRedirectUri: Boolean(redirectUri)
    });
    return false;
  }
  return true;
}

/**
 * Starts OAuth flow with error checking
 * @returns {Object} Auth URL or error message
 */
function startOAuthFlow() {
  try {
    const state = Utilities.getUuid();
    PropertiesService.getUserProperties().setProperty('oauth_state', state);

    const authUrl = `${CONFIG.GITHUB_AUTH_URL}?` +
      `client_id=${CONFIG.OAUTH_CLIENT_ID}&` +
      `redirect_uri=${encodeURIComponent(CONFIG.OAUTH_REDIRECT_URI)}&` +
      `scope=${encodeURIComponent(CONFIG.OAUTH_SCOPES.join(' '))}&` +
      `state=${state}`;

    console.log('Generated auth URL:', authUrl);

    return {
      success: true,
      authUrl: authUrl
    };
  } catch (error) {
    console.error('OAuth initialization error:', error);
    return {
      success: false,
      error: `Failed to start authentication: ${error.message}`
    };
  }
}


/**
 * Handles the OAuth callback and exchanges code for access token
 * @param {string} code - Authorization code from GitHub
 * @param {string} state - State parameter for verification
 * @returns {Object} Object containing success status and any error message
 */
function handleOAuthCallback(code, state) {
  try {
    // Verify state parameter matches stored value
    const savedState = PropertiesService.getUserProperties().getProperty('oauth_state');
    if (state !== savedState) {
      throw new Error('State parameter mismatch - possible CSRF attempt');
    }

    // Exchange authorization code for access token
    const response = UrlFetchApp.fetch(CONFIG.GITHUB_TOKEN_URL, {
      method: 'post',
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: {
        client_id: CONFIG.OAUTH_CLIENT_ID,
        client_secret: CONFIG.OAUTH_CLIENT_SECRET,
        code: code,
        redirect_uri: CONFIG.OAUTH_REDIRECT_URI
      },
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());
    if (result.error) {
      throw new Error(`GitHub OAuth error: ${result.error_description}`);
    }

    // Store the access token in user properties
    PropertiesService.getUserProperties().setProperty('github_access_token', result.access_token);

    return { success: true };
  } catch (error) {
    console.error('OAuth callback error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Checks if user is authenticated
 * @returns {boolean} Authentication status
 */
function isUserAuthenticated() {
  // Check if the user has authenticated by checking the token
  const token = PropertiesService.getUserProperties().getProperty('github_access_token');
  return Boolean(token);
}

/**
 * Fetches user permissions from GitHub
 * @returns {Object|null} User permissions or null if error
 */
function getPermissions() {
  try {
    const accessToken = PropertiesService.getUserProperties().getProperty('github_access_token');
    if (!accessToken) {
      throw new Error('No access token found');
    }

    // Make a request to GitHub to get the user details and permissions
    const permissionResponse = UrlFetchApp.fetch('https://api.github.com/user', {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });

    return JSON.parse(permissionResponse.getContentText());
  } catch (error) {
    console.error('Failed to fetch permissions:', error);
    return null;
  }
}

/**
 * Verifies GitHub token permissions
 * @returns {Object} Object containing verification results
 */
function verifyGitHubAccess() {
  try {
    const client = createGitHubClient();
    
    // Check user access
    const userUrl = `${client.baseUrl}/user`;
    const userResponse = UrlFetchApp.fetch(userUrl, {
      headers: client.headers,
      muteHttpExceptions: true
    });

    if (userResponse.getResponseCode() !== 200) {
      throw new Error('Invalid GitHub token or token expired');
    }

    // Check organization access
    const orgUrl = `${client.baseUrl}/orgs/${CONFIG.ORG_NAME}`;
    const orgResponse = UrlFetchApp.fetch(orgUrl, {
      headers: client.headers,
      muteHttpExceptions: true
    });

    if (orgResponse.getResponseCode() !== 200) {
      throw new Error(`No access to organization ${CONFIG.ORG_NAME}`);
    }

    return {
      success: true,
      user: JSON.parse(userResponse.getContentText()).login,
      organization: CONFIG.ORG_NAME
    };
  } catch (error) {
    console.error('GitHub access verification failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Clears stored OAuth token
 * Used for logout or token refresh
 */
function clearStoredToken() {
  PropertiesService.getUserProperties().deleteProperty('github_access_token');
  Logger.log('Token cleared - app will require reauthorization');
}
/* =============================================================================
GITHUB API INTERACTIONS
=============================================================================
*/

/**
 * Creates configured GitHub API client
 * @returns {Object} Configured GitHub client with auth headers
 */
function createGitHubClient() {
  const token = PropertiesService.getUserProperties().getProperty('github_access_token');
  
  if (!token) {
    throw new Error('Authentication required. Please authenticate with GitHub first.');
  }
  
  return {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json',
      'User-Agent': 'QuickGit'
    },
    baseUrl: CONFIG.GITHUB_API_BASE,
    validateResponse: function(response) {
      const code = response.getResponseCode();
      if (code !== 200 && code !== 201) {
        const error = JSON.parse(response.getContentText());
        throw new Error(`GitHub API Error: ${error.message} (${code})`);
      }
      return response;
    }
  };
}

/**
 * Fetches repositories for the organization
 * Includes pagination and caching
 * @returns {Array} List of repositories
 */
function fetchOrgRepos() {
  const client = createGitHubClient();
  const cache = CacheService.getScriptCache();
  const cacheKey = `org_repos_${CONFIG.ORG_NAME}`;

  try {
    // Check cache first
    const cachedRepos = cache.get(cacheKey);
    if (cachedRepos) {
      Logger.log('Returning cached repositories');
      return JSON.parse(cachedRepos);
    }

    Logger.log('Cache miss - fetching repositories from GitHub');
    let allRepos = [];
    let nextPage = 1;
    const perPage = 100; // Maximum allowed by GitHub API

    while (true) {
      const url = `${client.baseUrl}/orgs/${CONFIG.ORG_NAME}/repos?per_page=${perPage}&page=${nextPage}&sort=full_name`;
      Logger.log(`Fetching page ${nextPage}`);
      
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: client.headers,
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        throw new Error(JSON.parse(response.getContentText()).message);
      }

      const pageRepos = JSON.parse(response.getContentText());
      if (!pageRepos || pageRepos.length === 0) {
        break;
      }

      // Only keep essential fields
      const simplifiedRepos = pageRepos.map(repo => ({
        name: repo.name,
        full_name: repo.full_name,
        private: repo.private
      }));

      allRepos = allRepos.concat(simplifiedRepos);
      
      // Check if there are more pages
      const linkHeader = response.getHeaders()['Link'];
      if (!linkHeader || !linkHeader.includes('rel="next"')) {
        break;
      }
      
      nextPage++;
      
      // Add a small delay to avoid rate limiting
      Utilities.sleep(100);
    }

    Logger.log(`Total repositories fetched: ${allRepos.length}`);

    // Cache the results
    if (allRepos.length > 0) {
      cache.put(cacheKey, JSON.stringify(allRepos), CONFIG.CACHE_DURATION);
    }

    return allRepos;

  } catch (error) {
    Logger.log(`Error in fetchOrgRepos: ${error.message}`);
    throw new Error('Failed to fetch repositories: ' + error.message);
  }
}

/**
 * Fetches all organization members with pagination and caching
 * @returns {Array<Object>} List of organization members with login and avatar info
 * @throws {Error} If the API request fails or returns non-200 status
 */
function fetchOrgMembers() {
  const client = createGitHubClient();
  const cache = CacheService.getScriptCache();
  const cacheKey = `org_members_${CONFIG.ORG_NAME}`;

  try {
    // Check cache first
    const cachedMembers = cache.get(cacheKey);
    if (cachedMembers) {
      return JSON.parse(cachedMembers);
    }

    // Fetch members with pagination
    let members = [];
    let page = 1;
    const perPage = 100; // GitHub API max per page

    while (true) {
      // Construct URL with pagination params
      const url = `${client.baseUrl}/orgs/${CONFIG.ORG_NAME}/members?` + 
                 `per_page=${perPage}&page=${page}&role=all`;

      // Make API request with error handling
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: client.headers,
        muteHttpExceptions: true
      });

      // Check response status
      const responseCode = response.getResponseCode();
      if (responseCode !== 200) {
        const error = JSON.parse(response.getContentText());
        throw new Error(`GitHub API error: ${error.message} (${responseCode})`);
      }

      // Parse response and check for more pages
      const pageMembers = JSON.parse(response.getContentText());
      if (!pageMembers || !pageMembers.length) {
        break; // No more members to fetch
      }

      // Add members with essential fields only
      members = members.concat(pageMembers.map(member => ({
        login: member.login,
        avatar_url: member.avatar_url
      })));

      page++;
    }

    // Sort members by login name
    members.sort((a, b) => a.login.localeCompare(b.login));

    // Cache results for 6 hours (21600 seconds)
    if (members.length > 0) {
      cache.put(cacheKey, JSON.stringify(members), CONFIG.CACHE_DURATION);
    }

    return members;

  } catch (error) {
    Logger.log(`Error in fetchOrgMembers: ${error.message}`);
    throw new Error(`Failed to fetch organization members: ${error.message}`);
  }
}

/*/**
 * Fetches open issues for a repository with enhanced error handling
 * @param {string} repoName - Repository name
 * @returns {Array} List of open issues
 */
function fetchRepoIssues(repoName) {
  try {
    // First validate the GitHub token
    const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    if (!token) {
      throw new Error('GitHub token not configured. Please set up authentication.');
    }

    // Create GitHub client
    const client = createGitHubClient();
    
    // First verify repository access
    const repoCheckUrl = `${client.baseUrl}/repos/${CONFIG.ORG_NAME}/${repoName}`;
    const repoCheck = UrlFetchApp.fetch(repoCheckUrl, {
      headers: client.headers,
      muteHttpExceptions: true
    });
    
    if (repoCheck.getResponseCode() !== 200) {
      const error = JSON.parse(repoCheck.getContentText());
      if (repoCheck.getResponseCode() === 404) {
        throw new Error(`Repository ${repoName} not found or no access.`);
      } else if (repoCheck.getResponseCode() === 403) {
        throw new Error(`Insufficient permissions to access ${repoName}. Please check your token permissions.`);
      }
      throw new Error(error.message || `Failed to access repository: ${repoCheck.getResponseCode()}`);
    }

    // If we can access the repo, fetch issues
    const issuesUrl = `${client.baseUrl}/repos/${CONFIG.ORG_NAME}/${repoName}/issues?state=open`;
    const response = UrlFetchApp.fetch(issuesUrl, {
      headers: client.headers,
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      const error = JSON.parse(response.getContentText());
      throw new Error(`Failed to fetch issues: ${error.message}`);
    }

    const issues = JSON.parse(response.getContentText());
    
    // Log success for debugging
    console.log(`Successfully fetched ${issues.length} issues from ${repoName}`);
    
    return issues.map(issue => ({
      number: issue.number,
      title: issue.title,
      state: issue.state,
      html_url: issue.html_url
    }));

  } catch (error) {
    console.error(`Error in fetchRepoIssues for ${repoName}:`, error);
    throw new Error(error.message || 'Failed to fetch issues. Please check repository access and permissions.');
  }
}

/* =============================================================================
DOCUMENT PROCESSING
=============================================================================
*/

/**
 * Main processor for Google Doc with rich text preservation
 * @param {string} docUrl - URL of the Google Doc to process
 * @returns {Object} Object containing arrays of new and update issues
 */
function processGoogleDoc(docUrl) {
  try {
    if (!docUrl) {
      throw new Error('Document URL is required');
    }
    
    const doc = DocumentApp.openByUrl(docUrl);
    if (!doc) {
      throw new Error('Could not open document');
    }

    // Get document metadata
    const docMetadata = {
      title: doc.getName(),
      url: docUrl
    };

    const body = doc.getBody();
    return parseDocContentWithFormatting(body, docMetadata);
    
  } catch (error) {
    Logger.log('Error processing document: ' + error.message);
    throw new Error('Failed to process document: ' + error.message);
  }
}

/**
 * Converts a string or Google Doc element to HTML
 * @param {string|Element} element - Content to convert
 * @returns {string} HTML formatted string
 */
function convertElementToHtml(element) {
  if (!element) return '';
  
  try {
    // Handle string inputs directly
    if (typeof element === 'string') {
      return element;
    }
    
    // Handle list items specially
    if (element.getType && element.getType() === DocumentApp.ElementType.LIST_ITEM) {
      return processListItem(element);
    }
    
    // Get text content
    let text;
    if (element.asText) {
      text = element.asText();
    } else if (element.getText) {
      text = element.getText();
    } else {
      // If neither method exists, treat as plain text
      return String(element);
    }
    
    if (!text) return '';
    
    const textStr = typeof text === 'string' ? text : text.getText();
    if (!textStr) return '';
    
    let result = '';
    let i = 0;
    
    // Process text with formatting if it's a Doc element
    if (text.isBold) {  // Check if it has formatting methods
      while (i < textStr.length) {
        const styles = {
          bold: text.isBold(i),
          italic: text.isItalic(i),
          link: text.getLinkUrl(i)
        };
        
        // Find next formatting change
        let j = i + 1;
        while (j < textStr.length) {
          if (styles.bold !== text.isBold(j) ||
              styles.italic !== text.isItalic(j) ||
              styles.link !== text.getLinkUrl(j)) {
            break;
          }
          j++;
        }
        
        // Extract and format text segment
        let segment = textStr.substring(i, j);
        result += applyFormatting(segment, styles);
        i = j;
      }
      return result;
    } else {
      // Return plain text if no formatting methods
      return textStr;
    }
  } catch (error) {
    console.error('Error converting element to HTML:', error);
    // Return the plain text content as fallback
    return element.toString();
  }
}

/**
 * Applies markdown formatting to text segment
 * @param {string} text - Text to format
 * @param {Object} styles - Formatting styles to apply
 * @returns {string} Formatted text
 */
function applyFormatting(text, styles) {
  if (!text) return '';
  
  try {
    let result = text;
    
    if (styles.link) {
      const cleanText = text.trim();
      return `[${cleanText}](${styles.link}) `;
    }
    
    if (styles.bold) {
      result = `**${result}**`;
    }
    if (styles.italic) {
      result = `_${result}_`;
    }
    
    return result;
  } catch (error) {
    console.error('Error applying formatting:', error);
    return text;
  }
}

/**
 * Process list items with proper formatting
 * @param {Element} element - The list item element
 * @returns {string} Formatted text with proper list markers and indentation
 */
function processListItem(element) {
  try {
    const listItem = element.asListItem();
    const nestingLevel = listItem.getNestingLevel() || 0;
    const indent = '  '.repeat(nestingLevel);
    
    // Get text with any inline formatting but without list markers
    const text = element.asText();
    let content = '';
    let i = 0;
    const textStr = text.getText();
    
    while (i < textStr.length) {
      const styles = {
        bold: text.isBold(i),
        italic: text.isItalic(i),
        link: text.getLinkUrl(i)
      };
      
      let j = i + 1;
      while (j < textStr.length) {
        if (styles.bold !== text.isBold(j) ||
            styles.italic !== text.isItalic(j) ||
            styles.link !== text.getLinkUrl(j)) {
          break;
        }
        j++;
      }
      
      const segment = textStr.substring(i, j);
      content += applyFormatting(segment, styles);
      i = j;
    }
    
    return `${indent}* ${content.trim()}`;
  } catch (error) {
    console.error('Error processing list item:', error);
    return element.getText(); // Fallback to plain text
  }
}

/**
 * Parses document content into structured format
 * @param {string} content - Raw document content
 * @param {Object} docMetadata - Document metadata
 * @returns {Object} Parsed content structure
 */
function parseDocContent(content, docMetadata) {
  // Debug log the incoming content
  console.log("Parsing content:", content);

  const results = {
    newIssues: [],
    updateIssues: []
  };

  // Split content into lines and clean up
  const lines = content.split('\n')
    .map(line => line.trim())
    .filter(line => line);
  
  // Debug log the lines
  console.log("Processed lines:", lines);

  let currentIssue = null;
  let contentLines = [];

  // Process each line
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // Debug log each line being processed
    console.log(`Processing line ${i}:`, line);
    
    // Check for issue markers
    const newMatch = line.match(/^new\s+issue:\s*(.+)/i);
    const updateMatch = line.match(/^update\s+issue:\s*(.+)/i);

    // Debug log matches
    if (newMatch) console.log("Found new issue:", newMatch[1]);
    if (updateMatch) console.log("Found update issue:", updateMatch[1]);

    if (newMatch || updateMatch) {
      // Save previous issue if exists
      if (currentIssue) {
        console.log("Finalizing previous issue:", currentIssue);
        finalizeIssue(currentIssue, contentLines, results, docMetadata);
      }

      // Start new issue
      currentIssue = {
        type: newMatch ? 'new' : 'update',
        title: (newMatch || updateMatch)[1].trim(),
        lineNumber: i + 1
      };
      contentLines = [];
      console.log("Started new issue:", currentIssue);
    }
    // If we have a current issue and this isn't the start of another issue
    else if (currentIssue && !hasUpcomingIssue(line)) {
      contentLines.push(line);
      console.log("Added content line:", line);
    }
  }

  // Handle last issue
  if (currentIssue) {
    console.log("Finalizing last issue:", currentIssue);
    finalizeIssue(currentIssue, contentLines, results, docMetadata);
  }

  // Debug log final results
  console.log("Final results:", results);

  return results;
}

/**
 * Parses document content preserving formatting
 * @param {Body} body - Google Doc body element
 * @param {Object} docMetadata - Document metadata for linking
 * @returns {Object} Parsed new and update issues
 */
function parseDocContentWithFormatting(body, docMetadata) {
  try {
    const results = {
      newIssues: [],
      updateIssues: []
    };
    
    let currentIssue = null;
    let contentElements = [];
    
    if (!body) {
      console.log('Null document body provided');
      return results;
    }
    
    const numChildren = body.getNumChildren();
    console.log(`Processing ${numChildren} paragraphs`);

    function isIssueMarker(text) {
      text = text.toLowerCase().trim();
      return text.startsWith('new issue:') || 
             text.startsWith('update issue:') ||
             text.startsWith('update issue -') ||
             text.startsWith('new issue -');
    }

    function isSectionBreak(text) {
      text = text.trim();
      return text.startsWith('[') && text.endsWith(']') || // [images], [section], etc.
             text === '---' || // horizontal rule
             text === ''; // empty line
    }
    
    for (let i = 0; i < numChildren; i++) {
      const child = body.getChild(i);
      if (!child) continue;
      
      let text = '';
      try {
        const textElement = child.asText();
        if (textElement) {
          text = textElement.getText().trim();
          console.log(`Processing line ${i}:`, text);
        }
      } catch (error) {
        console.log(`Error getting text from child at index ${i}:`, error);
        continue;
      }
      
      if (!text) continue;
      
      // Look for issue markers
      const newMatch = text.match(/^new\s+issue:\s*(.+)/i);
      const updateMatch = text.match(/^update\s+issue:\s*(.+)/i);
      
      if (newMatch || updateMatch) {
        console.log("Found issue marker:", text);
        
        // Finalize previous issue if exists
        if (currentIssue) {
          finalizeIssue(currentIssue, contentElements, results, docMetadata);
          console.log("Finalized previous issue");
        }
        
        // Start new issue
        currentIssue = {
          type: newMatch ? 'new' : 'update',
          title: (newMatch || updateMatch)[1].trim(),
          lineNumber: i + 1
        };
        contentElements = []; // Reset content elements
        console.log("Started new issue:", currentIssue);
      } else if (currentIssue) {
        // Check if this line starts a new section or is a different issue marker
        if (isIssueMarker(text) || isSectionBreak(text)) {
          // Finalize current issue before the break
          finalizeIssue(currentIssue, contentElements, results, docMetadata);
          currentIssue = null;
          contentElements = [];
          console.log("Found boundary, finalized issue");
        } else {
          // Only add content if it's not a boundary
          contentElements.push(child);
          console.log("Added content:", text);
        }
      }
    }
    
    // Handle last issue
    if (currentIssue) {
      finalizeIssue(currentIssue, contentElements, results, docMetadata);
      console.log("Finalized last issue");
    }
    
    console.log("Final results:", results);
    return results;
  } catch (error) {
    console.error('Error in parseDocContentWithFormatting:', error);
    throw error;
  }
}

/**
 * Checks if a line starts a new issue
 * @param {string} line - Line to check
 * @returns {boolean} True if line starts a new issue
 */
function isNextIssueLine(line) {
  return /^(new|update)\s+issue:?/i.test(line);
}

/**
 * Checks if any of the upcoming lines starts a new issue
 * @param {string[]} upcomingLines - Array of lines to check
 * @returns {boolean} True if next issue marker found
 */
function isNextIssueLine(upcomingLines) {
  for (const line of upcomingLines) {
    if (line.toLowerCase().match(/^(new|update)\s+issue:/)) {
      return true;
    }
  }
  return false;
}

/**
 * Checks if a line contains an issue marker
 * @param {string} line - Line to check
 * @returns {boolean} True if line contains issue marker
 */
function isIssueMarker(line) {
  line = line.toLowerCase();
  return line.startsWith('new issue:') || line.startsWith('update issue:');
}

/**
 * Checks if a line is an issue marker
 * @param {string} line - Line to check
 * @returns {boolean} True if line starts a new issue
 */
function isIssueLine(line) {
  return /^(new|update)\s+issue:?/i.test(line);
}

/**
 * Checks if any upcoming lines contain issue markers
 * @param {string[]} upcomingLines - Array of lines to check
 * @returns {boolean} True if next issue marker found
 */
function hasUpcomingIssue(upcomingLines) {
  for (const line of upcomingLines) {
    if (line.toLowerCase().match(/^(new|update)\s+issue:/)) {
      return true;
    }
  }
  return false;
}

/* =============================================================================
ISSUE MANAGEMENT
=============================================================================
*/

/**
 * Creates a new issue with proper assignee handling using two-step process if needed
 * @param {string} repo - Repository name
 * @param {Object} issueData - Issue data including title, body, and assignees
 * @returns {Object} Created issue details
 */
function createIssue(repo, issueData) {
  const client = createGitHubClient();
  const url = `${client.baseUrl}/repos/${CONFIG.ORG_NAME}/${repo}/issues`;
  
  try {
    // Format payload ensuring assignees is an array
    const payload = {
      title: issueData.title,
      body: issueData.body,
      assignees: Array.isArray(issueData.assignees) ? issueData.assignees : 
                issueData.assignee ? [issueData.assignee] : 
                []
    };

    Logger.log('Creating issue with payload:', JSON.stringify(payload));

    // First step: Create the issue
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: client.headers,
      muteHttpExceptions: true,
      payload: JSON.stringify(payload)
    });

    if (response.getResponseCode() !== 201) {
      const errorResponse = JSON.parse(response.getContentText());
      throw new Error(`Failed to create issue: ${errorResponse.message}`);
    }

    const createdIssue = JSON.parse(response.getContentText());
    Logger.log('Issue created:', createdIssue.number);

    // Second step: If issue was created but assignees are empty, try direct assignment
    if (createdIssue.number && payload.assignees.length && (!createdIssue.assignees || !createdIssue.assignees.length)) {
      Logger.log('Attempting direct assignment for issue:', createdIssue.number);
      
      const assignUrl = `${url}/${createdIssue.number}/assignees`;
      const assignResponse = UrlFetchApp.fetch(assignUrl, {
        method: 'POST',
        headers: client.headers,
        payload: JSON.stringify({ assignees: payload.assignees }),
        muteHttpExceptions: true
      });

      if (assignResponse.getResponseCode() === 201) {
        const updatedIssue = JSON.parse(assignResponse.getContentText());
        Logger.log('Assignment successful:', updatedIssue.assignees);
        return {
          number: updatedIssue.number,
          html_url: updatedIssue.html_url
        };
      } else {
        Logger.log('Assignment failed, but issue was created');
      }
    }

    return {
      number: createdIssue.number,
      html_url: createdIssue.html_url
    };
    
  } catch (error) {
    Logger.log('Error in createIssue:', error);
    throw error;
  }
}

/**
 * Updates an existing issue in the specified GitHub repository.
 * Uses the GitHub API to post a comment on an existing issue.
 * @param {string} repo - The name of the repository where the issue exists.
 * @param {number} issueNumber - The number of the issue to update.
 * @param {string} comment - The comment to add to the existing issue.
 * @returns {Object} Object containing the updated issue number and the HTML URL of the added comment.
 * @throws {Error} If updating the issue fails or response is missing expected fields.
 */
function updateIssue(repo, issueNumber, comment) {
  // Create the GitHub API client
  const client = createGitHubClient();
  const url = `${client.baseUrl}/repos/${CONFIG.ORG_NAME}/${repo}/issues/${issueNumber}/comments`;

  try {
    // Log the URL and payload for debugging
    Logger.log(`Updating issue at URL: ${url}`);
    Logger.log(`Payload: ${JSON.stringify({ body: comment })}`);

    // Ensure that the payload includes all necessary fields in the correct format
    const payload = {
      body: comment
    };

    // Make the POST request to add a comment to the issue
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: client.headers,
      muteHttpExceptions: true,
      payload: JSON.stringify(payload)
    });

    // Check if the request was successful (HTTP 201 Created)
    if (response.getResponseCode() === 201) {
      // Parse the response to get the updated issue details
      const updatedIssue = JSON.parse(response.getContentText());

      // Validate that the html_url is present in the response
      if (updatedIssue.html_url) {
        Logger.log(`Successfully updated issue with comment: ${updatedIssue.html_url}`);
        return {
          number: issueNumber,
          html_url: updatedIssue.html_url
        };
      } else {
        throw new Error("html_url missing in response for updated issue comment.");
      }
    } else {
      // If not successful, log full response details for debugging
      Logger.log(`Full response for failed issue update: ${response.getContentText()}`);
      const errorResponse = JSON.parse(response.getContentText());
      throw new Error(`Failed to update issue #${issueNumber}: ${errorResponse.message}. Ensure the issue number is correct and the GitHub token has the appropriate permissions.`);
    }
  } catch (error) {
    // Log any errors that occur during issue updating
    Logger.log(`Error in updateIssue: ${error.message}`);
    throw error;
  }
}

/**
 * Processes confirmed issues with improved assignee handling
 * @param {Object} data - Object containing new issues and updates
 * @returns {Object} Results including links to created/updated items
 */
function processConfirmedIssues(data) {
  const results = {
    created: [],
    updated: [],
    errors: [],
    links: []
  };

  try {
    if (data.newIssues?.length) {
      Logger.log('Processing new issues:', data.newIssues);
      
      for (const issue of data.newIssues) {
        try {
          // Log the incoming issue data
          Logger.log('Processing issue:', issue);
          
          if (!issue.repo) {
            throw new Error('Repository is required');
          }

          // Create issue with explicit assignee handling
          const issueData = {
            title: issue.title,
            body: issue.body,
            assignees: issue.assignee ? [issue.assignee] : [] // Ensure assignee is in array format
          };

          Logger.log('Prepared issue data:', issueData);
          const createdIssue = createIssue(issue.repo, issueData);
          Logger.log('Created issue response:', createdIssue);

          results.created.push(createdIssue.number);
          results.links.push({
            type: 'issue',
            number: createdIssue.number,
            url: createdIssue.html_url,
            repo: issue.repo,
            title: issue.title,
            assignees: createdIssue.assignees // Track assignees in results
          });
        } catch (error) {
          Logger.log('Error creating issue:', error);
          results.errors.push(`Failed to create issue "${issue.title}": ${error.message}`);
        }
      }
    }

    // Handle update issues (unchanged)
    if (data.updateIssues?.length) {
      for (const update of data.updateIssues) {
        try {
          const updatedIssue = updateIssue(update.repo, update.issueNumber, update.comment);
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
    Logger.log('Error in processConfirmedIssues:', error);
    throw new Error(`Failed to process issues: ${error.message}`);
  }
}

/**
 * Finalizes and formats an issue before adding to results
 * @param {Object} issue - Current issue being processed
 * @param {Array<Element>} contentElements - Additional content elements
 * @param {Object} results - Results object to update
 * @param {Object} docMetadata - Document metadata for linking
 */
function finalizeIssue(issue, contentElements, results, docMetadata) {
  try {
    // Skip if missing required data
    if (!issue || !issue.type) {
      console.error('Invalid issue data provided to finalizeIssue');
      return;
    }

    // Format content, even if there are no elements (will just get footer)
    const formattedContent = finalizeIssueContent(contentElements || [], docMetadata);

    if (issue.type === 'new') {
      // For new issues, use title and body structure
      results.newIssues.push({
        title: issue.title,
        body: formattedContent,
        lineNumber: issue.lineNumber
      });
    } else {
      // For update issues, use title and comment structure
      // For update issues, the title content becomes part of the comment
      const commentContent = issue.title + '\n\n' + formattedContent;
      
      results.updateIssues.push({
        title: null, // We don't need a separate title for updates
        comment: commentContent,
        lineNumber: issue.lineNumber
      });
    }
  } catch (error) {
    console.error('Error in finalizeIssue:', error);
  }
}

/**
 * Helper function to format issue content with proper formatting and footer
 * @param {Array} contentElements - Content elements
 * @param {Object} docMetadata - Document metadata
 * @param {Object} issue - The current issue being processed
 * @returns {string} Formatted issue content
 */
function finalizeIssueContent(contentElements, docMetadata, issue) {
  try {
    // Handle content elements
    const formattedContent = (contentElements || [])
      .map(element => {
        if (!element) return '';
        return convertElementToHtml(element);
      })
      .filter(content => content && content.trim())
      .join('\n');

    // For update issues, include the comment line as part of the content
    let content = formattedContent;
    if (issue && issue.type === 'update' && issue.title) {
      content = `${issue.title}\n${content}`;
    }

    // Add attribution footer with proper markdown
    const footer = docMetadata ?
      `\n---\n*Generated by [QuickGit](https://script.google.com/a/macros/nasa.gov/s/AKfycbwbJ8OCr9TxiPE9caMtlhDPTAKIe0QsMY5bgaNO2N2heVqXF8ctnE0_k1Zu1bFmSLm1DA/exec) from [${docMetadata.title}](${docMetadata.url})*` :
      '';

    return content + footer;
  } catch (error) {
    console.error('Error in finalizeIssueContent:', error);
    return ''; // Return empty string on error
  }
}

/**
 * Gathers issue data from the UI with improved assignee handling
 * @returns {Object|null} Collected issue data or null if validation fails
 */
function gatherIssueData() {
  const newIssues = [];
  const updateIssues = [];
  
  try {
    // Gather new issues
    const newIssueElements = document.querySelectorAll('.issue-item[data-issue-type="new"]');
    newIssueElements.forEach(item => {
      const index = parseInt(item.dataset.issueIndex);
      const repoInput = item.querySelector('.repo-search');
      const assigneeInput = item.querySelector('.assignee-search');
      
      if (repoInput?.value) {
        // Log the gathered data
        Logger.log('Gathering data for new issue:', {
          title: documentContent.newIssues[index].title,
          assignee: assigneeInput?.value
        });

        newIssues.push({
          title: documentContent.newIssues[index].title,
          body: documentContent.newIssues[index].body,
          repo: repoInput.value,
          assignee: assigneeInput?.value || null // Explicitly handle empty assignee
        });
      }
    });
    
    // Gather update issues (unchanged)
    const updateIssueElements = document.querySelectorAll('.issue-item[data-issue-type="update"]');
    updateIssueElements.forEach(item => {
      const index = parseInt(item.dataset.issueIndex);
      const repoInput = item.querySelector('.repo-search');
      const issueInput = item.querySelector('.issue-search');
      
      if (repoInput?.value && issueInput?.value) {
        const issueNumber = issueInput.value.match(/#(\d+):/)?.[1];
        if (issueNumber) {
          updateIssues.push({
            repo: repoInput.value,
            issueNumber: parseInt(issueNumber),
            comment: documentContent.updateIssues[index].comment
          });
        }
      }
    });
    
    return {
      newIssues,
      updateIssues
    };
  } catch (error) {
    console.error('Error gathering issue data:', error);
    return null;
  }
}

/* =============================================================================
UTILITY FUNCTIONS
=============================================================================
*/

/**
 * Formats issue content with proper structure for display
 * @param {string[]} contentLines - Raw content lines
 * @returns {string} Formatted content
 */
function formatIssueContent(contentLines) {
  // Filter out empty lines and format links
  const formattedLines = contentLines
    .filter(line => line.trim())
    .map(line => {
      // Convert markdown links to proper format
      line = line.replace(/\[([^\]]+)\]\(([^\)]+)\)/g, '$1 ($2)');
      return line;
    });

  return formattedLines.join('\n');
}

/**
 * Separates title and description from issue content
 * @param {string} content - Raw issue content
 * @returns {Object} Separated title and description
 */
function separateTitleAndDescription(content) {
  let title, description;
  
  // Split on first hyphen or semicolon
  const separatorMatch = content.match(/^([^-;]+)[-;](.+)$/);
  
  if (separatorMatch) {
    title = separatorMatch[1].trim();
    description = separatorMatch[2].trim();
  } else {
    title = content.trim();
    description = '';
  }
  
  return { title, description };
}

/**
 * @param {string} type - The type of issue action (e.g., 'New Issue', 'Updated Issue').
 * @param {string} title - The title or identifier for the issue.
 * @param {string} url - The URL of the GitHub issue or comment.
 * @returns {string} HTML string for displaying the link.
 */
function generateIssueLinkHtml(type, title, url) {
  // Hide previous banners or messages before generating new ones
  hideRepositoryBanner();
  return `<div><strong>${type}:</strong> <a href="${url}" target="_blank">${title}</a></div>`;
}

/**
 * Gets document metadata from URL
 * @param {string} docUrl - Document URL
 * @returns {Object} Document metadata
 */
function getDocumentMetadata(docUrl) {
  try {
    const doc = DocumentApp.openByUrl(docUrl);
    return {
      title: doc.getName(),
      url: docUrl
    };
  } catch (error) {
    Logger.log(`Error getting document metadata: ${error.message}`);
    return null;
  }
}

// Add any additional utility functions here...

/* =============================================================================
HELPER FUNCTIONS
=============================================================================
*/

/**
 * Helper function to include HTML files
 * @param {string} filename - Name of the HTML file to include
 * @returns {string} The evaluated HTML content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Clears previous issues and comments.
 */
function clearPreviousIssuesAndComments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GitHub Issues');
  if (sheet) {
    sheet.clear();
  }
}

/*
 * Hides the repository applied banner.
 */
function hideRepositoryBanner() {
  const ui = SpreadsheetApp.getUi();
  const bannerElement = ui.alert('Repository applied successfully!', ui.ButtonSet.OK);
  Utilities.sleep(500);
}



