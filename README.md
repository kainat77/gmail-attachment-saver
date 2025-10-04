# Gmail Attachment Saver

A Google Apps Script add-on that allows you to save Gmail attachments directly to Google Drive with advanced folder navigation, quick access shortcuts, and file renaming options.

## Features

- **Quick Access Folders** - Save up to 10 favorite folders for one-click access
- **Smart Folder Browsing** - Navigate through My Drive and Starred folders
- **Search Functionality** - Find any folder by name
- **Batch File Renaming** - Rename files before saving with automatic timestamp options
- **File Size Display** - See attachment sizes before saving
- **Recent Folders** - Quick access to recently modified folders

## Screenshots

The add-on provides:
- A clean interface with folder icons and visual hierarchy
- Search functionality with real-time results
- Folder browsing with breadcrumb navigation
- Success confirmation with file links

## Installation

### Step 1: Create Apps Script Project

1. Go to [Google Apps Script](https://script.google.com)
2. Click **New Project**
3. Name your project (e.g., "Gmail Attachment Saver")

### Step 2: Add the Code

1. Delete the default `function myFunction()` code
2. Copy the entire code from `Code.gs` in this repository
3. Paste it into the script editor

### Step 3: Create Manifest File

1. In the Apps Script editor, click the **Project Settings** (gear icon)
2. Check **"Show 'appsscript.json' manifest file in editor"**
3. Go back to the **Editor** tab
4. Click on `appsscript.json` in the file list
5. Replace the content with the code from `appsscript.json` in this repository

### Step 4: Enable Drive API

1. In the Apps Script editor, click **Services** (+ icon on left sidebar)
2. Find **Google Drive API**
3. Select version **v2**
4. Click **Add**

### Step 5: Deploy the Add-on

1. Click **Deploy** > **Test deployments**
2. Click **Install**
3. Select your Gmail account
4. Review and authorize the required permissions:
   - Read email messages when the add-on is running
   - View and manage Google Drive files
5. Click **Allow**

### Step 6: Authorize Scopes (First Time)

1. In the Apps Script editor, select the function `authorizeScopes` from the dropdown
2. Click **Run**
3. Review permissions and click **Allow**
4. Check the **Execution log** to confirm successful authorization

## Usage

### Opening the Add-on

1. Open any email with attachments in Gmail
2. Click the add-on icon in the right sidebar
3. The add-on will display all attachments

### Quick Access Setup

1. Click **"Set Up Quick Access"** or **"Manage Quick Access"**
2. Browse to a folder or enter a folder ID
3. Click **"Add to Quick Access"** (⭐ button)
4. The folder will appear in Quick Access for future use

### Saving Attachments

1. Choose a destination folder:
   - Click a Quick Access folder
   - Browse My Drive or Starred folders
   - Search by folder name
   - Enter a folder ID manually
   - Create a new folder

2. Optionally rename files:
   - Toggle "Add timestamp" to append date/time to filenames
   - Edit individual filenames
   - Leave blank to skip specific files

3. Click **"Save to Drive"**

4. View results with links to open saved files

## Features Explained

### Quick Access Folders
- Save up to 10 frequently used folders
- One-click access for faster workflow
- Manage from the Quick Access screen

### Folder Browsing
- **My Drive** - Navigate your entire Drive hierarchy
- **Starred** - Access starred folders
- **Recent** - See recently modified folders (in Advanced section)
- Navigate with breadcrumb paths

### File Renaming
- **Timestamp option** - Adds current date/time to prevent duplicates
- **Custom names** - Rename each file individually
- **Skip files** - Leave filename blank to skip saving that file

### Search
- Search for any folder by name
- Results show folder path
- Options to select or browse into folders

## Permissions Explained

This add-on requires the following permissions:

- **Gmail messages (when running)** - To read attachments from the current email
- **Google Drive** - To save files and browse folders
- **External requests** - For Drive API calls

Your data is never stored externally. All operations happen within your Google account.

## Troubleshooting

### "No attachments in this email"
- The email must have at least one attachment
- Some inline images may not be detected as attachments

### "Invalid folder ID or no access"
- Verify the folder ID is correct
- Ensure you have write access to the folder
- Shared folders must be added to "My Drive" first

### "Maximum 10 folders"
- Quick Access is limited to 10 folders
- Remove a folder before adding a new one

### Large files taking too long
- Files over 10 MB may take time to upload
- The add-on will show a warning for large files
- Consider saving files in smaller batches

### Authorization issues
1. Run the `authorizeScopes()` function manually
2. Check that Drive API is enabled in Services
3. Reauthorize the add-on from Deploy > Test deployments

## Uninstallation

1. Go to [Gmail Settings](https://mail.google.com/mail/u/0/#settings/addons)
2. Find "Gmail Attachment Saver" under "Developer add-ons"
3. Click **Remove**

Or:

1. Go to [Apps Script](https://script.google.com)
2. Open the project
3. Click **Deploy** > **Test deployments**
4. Click **Uninstall**

## Technical Details

- **Platform**: Google Apps Script
- **APIs Used**: Gmail API, Google Drive API v2
- **Storage**: User Properties Service (for Quick Access folders)
- **Cache**: Cache Service (for recent folders)
- **Execution Time**: Optimized for typical use; large batches may approach 6-minute limit

## Privacy & Security

- No data leaves your Google account
- No external servers or databases
- All processing happens in your Apps Script environment
- Quick Access preferences stored in your User Properties
- Open source - you can review all code

## Contributing

Contributions are welcome! Feel free to:
- Report bugs via Issues
- Suggest features
- Submit pull requests
- Improve documentation

## License

MIT License - Feel free to modify and distribute

## Support

For issues or questions:
1. Check the Troubleshooting section above
2. Review the code and comments
3. Open an issue on GitHub

## Version History

**v1.0.0** - Initial release
- Quick Access folders
- Folder browsing and search
- File renaming with timestamps
- File size display
- Material Design UI

---

**Note**: This is an unofficial add-on and is not affiliated with or endorsed by Google.
