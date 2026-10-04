# Recording Applications to Google Sheets

The form posts every submission to a **Google Apps Script Web App**, which writes
a row into your Google Sheet (resumes are saved into a Drive folder and linked
in the sheet).

## Step 1 — Create the Google Sheet

1. Go to https://sheets.new
2. Name it e.g. `MLP Job Applications`
3. Note the **Spreadsheet ID** — the long string in the URL between
   `/d/` and `/edit` (you'll paste it into the script below).
   Example: `1AbC...xYz` from
   `https://docs.google.com/spreadsheets/d/1AbC...xYz/edit`

## Step 2 — Add the Apps Script

1. In the sheet: **Extensions → Apps Script**
2. Delete the sample code, paste the script below
3. Replace `PASTE_SPREADSHEET_ID_HERE` with your Spreadsheet ID
4. Save (Ctrl+S), name the project e.g. `MLP Applications API`

```javascript
var SPREADSHEET_ID = 'PASTE_SPREADSHEET_ID_HERE';

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);

    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    var sheet = ss.getSheetByName('Applications') || ss.insertSheet('Applications');
    var questions = ss.getSheetByName('Questions') || ss.insertSheet('Questions');
    var subFolder = getFolder_(ss.getId(), 'Resumes');

    // Resume
    var resumeLink = '';
    if (data.resumeName && data.resumeBase64) {
      var blob = Utilities.newBlob(Utilities.base64Decode(data.resumeBase64),
        data.resumeType || 'application/octet-stream', data.resumeName);
      var file = subFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      resumeLink = file.getUrl();
    }

    // Headers (auto-create)
    var keys = Object.keys(data).filter(function (k) {
      return k !== 'resumeBase64' && k !== 'resumeType';
    });
    keys = keys.concat(resumeLink ? ['resumeLink'] : ['resumeLink']);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(keys.concat(['ReceivedAt']));
    }
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var row = headers.map(function (h) {
      if (h === 'resumeLink') return resumeLink;
      if (h === 'ReceivedAt') return new Date();
      return data[h] !== undefined ? data[h] : '';
    });
    sheet.appendRow(row);

    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getFolder_(parentId, name) {
  var parent = DriveApp.getFileById(parentId);
  var it = parent.getParents().next().getFoldersByName(name);
  return it.hasNext() ? it.next()
    : parent.getParents().next().createFolder(name);
}
```

## Step 3 — Deploy the Web App

1. In the Apps Script editor: **Deploy → New deployment**
2. Gear icon (Select type) → **Web app**
3. Description: anything; **Execute as: Me**;
   **Who has access: Anyone** (important — applicants don't log in)
4. Click **Deploy** → authorize the permissions when prompted
   (it will warn "unverified app" → **Advanced → Go to project (unsafe)** — this is your own script)
5. Copy the **Web app URL** (ends in `/exec`)

## Step 4 — Connect the form

1. Open the form's `index.html`
2. Find the line near the top of the `<script>`:
   ```js
   var APP_SCRIPT_URL = 'PASTE_YOUR_GOOGLE_APPS_SCRIPT_URL_HERE';
   ```
3. Replace `PASTE_YOUR_GOOGLE_APPS_SCRIPT_URL_HERE` with your `/exec` URL
4. Commit & push (or ask me to do it)

Test: open the live form, fill it out, submit, and check the `Applications`
tab in your sheet — a new row should appear.
