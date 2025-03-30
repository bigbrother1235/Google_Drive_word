# GoogleDrive_final
# Google Drive Integration for Word

## Overview

This add-in for Microsoft Word allows users to access, browse, and insert content from their Google Drive directly within Word. Users can navigate their Google Drive folders, search for files, preview content, and insert it into their current document without leaving Word.

## Key Features

- **Google Drive Authentication**: Securely connect to your Google Drive account
- **File Navigation**: Browse folders and files within your Google Drive
- **Search Functionality**: Quickly find files by name
- **File Preview**: Preview various file types including:
  - PDF documents (with text selection capabilities)
  - Images
  - Google Docs
  - Word documents
  - Text files
- **Content Insertion**: Insert selected file content directly into your Word document
- **File Download**: Download Google Drive files directly from the add-in

## Requirements

- Microsoft Word 2016 or later
- Internet connection (for Google Drive access)
- Google account
- Platform supporting Office Add-ins (Windows, Mac, Web)

## Development Setup

This project is developed using VS Code's Office Add-in extension. To set up the development environment:

1. **Install necessary tools**:
   - [Visual Studio Code](https://code.visualstudio.com/)
   - [Office Yeoman Generator](https://github.com/OfficeDev/generator-office)
   - [Node.js](https://nodejs.org/) (LTS version)
   - [Python 3.6+](https://www.python.org/downloads/)

2. **Backend setup**:
   - Create a virtual environment in the `backend` directory:
     ```
     cd backend
     python -m venv venv
     # Windows
     venv\Scripts\activate
     # Linux/Mac
     source venv/bin/activate
     ```
   - Install dependencies:
     ```
     pip install -r requirements.txt
     ```
   - **Important**: Place your own `client_secret.json` file in the `backend` directory (obtained from Google Cloud Console)

3. **Google API setup**:
   - Create a project in the [Google Cloud Console](https://console.cloud.google.com/)
   - Enable the Google Drive API
   - Create OAuth credentials (Web application type)
   - Set the redirect URI to: `https://localhost:5001/taskpane.html`
   - Download the credentials JSON file and rename it to `client_secret.json`

## Project Structure

```
project/
├── frontend/              # Frontend files
│   ├── taskpane.html      # Main UI interface
│   ├── taskpane.css       # Styles
│   └── taskpane.js        # Scripts (merged from multiple files)
│
├── backend/               # Python backend
│   ├── app.py             # Flask main application
│   ├── auth.py            # Google authorization handler
│   ├── config.py          # Configuration file
│   ├── files.py           # File operations
│   ├── documents.py       # Document processing
│   ├── utils.py           # Utility functions
│   ├── requirements.txt   # Python dependencies
│   └── client_secret.json # Google API credentials (to be added by user)
```

## Running the Add-in

1. **Start the backend server**:
   ```
   cd backend
   # Activate virtual environment (as shown above)
   python app.py
   ```
   The backend service will run at `https://localhost:5001`.

2. **Start the frontend server**:
   ```
   cd frontend
   npm start
   ```

3. **Load the add-in**:
   - The project already includes a Manifest file; no need to create a new one

## Usage Guide

1. After loading the add-in, click the "Authorize Google Drive" button
2. Complete the Google account authorization flow
3. Use the breadcrumb navigation to browse your Google Drive folders
4. Click on files to preview their content
5. In the preview interface, you can:
   - Download the file
   - Open the file in Google Drive
   - Insert the file content into your current Word document
6. Use the search box to quickly find specific files

## Known Issues

1. **Merged JavaScript file**: `taskpane.js` is merged from multiple different JavaScript files, which may lead to suboptimal code organization. In future versions, we plan to refactor it into a more modular structure.

2. **Google API credentials**: Users must provide their own `client_secret.json` file, which requires setting up a project and API permissions in the Google Cloud Console.

3. **Certificate warnings**: Since the development environment uses self-signed certificates, browsers may display security warnings. This is normal, and you can safely proceed during development.

4. **Google Docs loading errors**: For Google Docs files, loading errors may occasionally occur. When encountering this issue, you can directly click the "Open in Google Drive" button to view the file content in Google Drive.

## Security Considerations

- This add-in only requests read-only permissions for Google Drive
- Your Google credentials are not stored on the backend server
- Authorization tokens are only temporarily stored in the browser's local storage

---

If you have any questions or suggestions, please submit an issue or contribute code to improve this project.
