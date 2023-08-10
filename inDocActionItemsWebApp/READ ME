# Google Apps Script Web App README

This repository contains code for a Google Apps Script-based web application designed to process Google Documents and extract specific action items. The application comprises of three main parts: a primary code (`code.js`), a web app-specific server-side script (`webAppCode.js`), and a client-side web page (`Page.html`).

## Table of Contents

- [Primary Script (`code.js`)](#primary-script-codejs)
- [Web App Server-side Script (`webAppCode.js`)](#web-app-server-side-script-webappcodejs)
  - [`doGet()`](#doget)
  - [`processDocumentId(input)`](#processdocumentidinput)
- [Client-side Web Page (`Page.html`)](#client-side-web-page-pagehtml)
  - [Key Elements](#key-elements)
  - [JavaScript](#javascript)
- [Instructions for Usage](#instructions-for-usage)

## Primary Script (`code.js`)

This script contains the primary functions and logic that drive the functionality of the web application. *(Note: You might need to provide specific details or functions contained in this script here for a clearer understanding.)*

## Web App Server-side Script (`webAppCode.js`)

### `doGet()`
- **Description:** Serves the HTML output to the user when the web app is accessed.
- **Returns:** The HTML page to display.

### `processDocumentId(input)`
- **Description:** Processes an input string to extract a Google Document ID. If the input is a full URL, the function extracts the ID part. Otherwise, it assumes the input is already a document ID.
- **Parameters:** 
  - `input`: The input string which can be a full Google Document URL or just the document ID.
- **Returns:** 'Success' or an error message.

## Client-side Web Page (`Page.html`)

This is an HTML page that provides a user interface for inputting a Google Document's URL or ID. It also displays the instructions for using the tool and an example image.

### Key Elements:

- **Input Field:** For users to enter a Google Document URL or ID.
- **Button:** Triggers the process to extract and list action items.
- **Message Box:** Displays responses from the server-side script.
- **Example Image:** Shows a sample Google Document. If the image fails to load, a link redirects the user to an actual Google Document.

### JavaScript:

#### `submitId()`
- **Description:** Collects the input from the user and communicates with the server-side script. On success, it updates the message box with a response.
  
---

## Instructions for Usage:

1. Access the web app.
2. Enter either a full Google Document URL or just the Document ID in the provided input field.
3. Click on the "Collect Actions" button.
4. Review the message box for the outcome.
5. Follow the displayed instructions and format specifications to ensure the Google Document can be successfully processed.

