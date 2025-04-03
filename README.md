# Exa API Integration for Google Sheets

## Description

This Google Apps Script allows you to leverage the power of the Exa API directly within your Google Sheets. Search the web, find similar content, and retrieve contents based on URLs without leaving your spreadsheet.

## Features

*   **Set API Key Securely:** Use a sidebar interface to save your Exa API key as a user property.
*   **Custom Sheet Functions:** Directly call Exa API endpoints using simple functions in your sheet cells (e.g., `=EXA_SEARCH(...)`). *(Note: Actual function names might differ based on implementation).*
*   **Seamless Integration:** Access Exa's capabilities within your existing spreadsheet workflows.

## Setup & Installation

1.  **Open your Google Sheet:** Go to the Google Sheet where you want to use this script.
2.  **Open the Script Editor:** Click on "Extensions" > "Apps Script".
3.  **Copy the Code:**
    *   Copy the contents of `Code.gs` (or your main `.gs` file) from this project and paste it into the `Code.gs` file in the Apps Script editor. Overwrite any existing template code.
    *   Create a new HTML file in the editor (File > New > HTML file). Name it `Sidebar.html` (ensure the name matches exactly, including capitalization).
    *   Copy the contents of the `Sidebar.html` file from this project and paste it into the newly created `Sidebar.html` file in the editor.
4.  **Save the Project:** Click the floppy disk icon (Save project) and give your script project a name (e.g., "Exa API Integration").
5.  **Refresh Your Sheet:** Close the Apps Script editor tab and refresh your Google Sheet browser tab.
6.  **Authorize the Script:**
    *   After refreshing, a new custom menu item should appear (e.g., "Exa Tools" - *this name might vary based on your `onOpen` function in `Code.gs`*). Click on it, then select the option to set the API key (e.g., "Set API Key").
    *   A dialog box will pop up asking for authorization. Review the permissions requested (it will likely need access to external services, script properties, and the current spreadsheet) and click "Allow". You might need to go through an "Advanced" > "Go to (project name)" flow if Google warns it's an unverified app.
7.  **Set Your API Key:**
    *   Once authorized, the sidebar should open.
    *   Go to [https://exa.ai/](https://exa.ai/) to get your API key if you don't have one.
    *   Paste your Exa API key into the input field in the sidebar and click "Save API Key".
    *   You should see a success message "✅ API Key saved!".

## Usage

1.  **Set API Key:** If you haven't already, use the custom menu (e.g., "Exa Tools" > "Set API Key") to open the sidebar and save your key. You only need to do this once per user unless your key changes.
2.  **Use Custom Functions:** In any cell within your Google Sheet, you can now use the custom functions defined in the script. Examples might include:
    *   `=EXA_SEARCH("your query here")` - To perform a standard Exa search.
    *   `=EXA_FINDSIMILAR("https://example.com")` - To find pages similar to the provided URL.
    *   `=EXA_CONTENTS("result_id_or_url")` - To retrieve the contents of a specific result.

3.  **Function Arguments:** Refer to your `Code.gs` file or documentation within it for the specific arguments accepted by each custom function (e.g., number of results, include domains, start/end dates, etc.).

## Configuration

*   **Exa API Key:** This script requires a valid Exa API key to function. The key is stored securely using Google Apps Script's User Properties service, scoped to the user who saves it.

---