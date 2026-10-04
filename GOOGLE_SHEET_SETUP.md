# Recording Applications to Google Sheets

The form posts every submission to a **Google Apps Script Web App**, which writes
a row into your Google Sheet («Applications» tab). Resumes are saved to a Drive
folder named «Resumes» and linked in the `resumeLink` column.

## One-time setup (Steps 1–3)

### Step 1 — Create the Google Sheet

1. Go to https://sheets.new and name it e.g. `MLP Job Applications`.

### Step 2 — Add the script

1. In the sheet: **Extensions → Apps Script**
2. Delete any sample code and paste the script below (it must be created from
   inside the sheet, so `getActiveSpreadsheet()` points at it).
3. Save (Ctrl+S).

```javascript
function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);

    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Script must be created from inside the Google Sheet (Extensions > Apps Script).');

    var sheet = ss.getSheetByName('Applications') || ss.insertSheet('Applications');
    sheet.setFrozenRows(1);

    // Save resume to Drive folder "Resumes", share as view link
    var resumeLink = '';
    if (data.resumeName && data.resumeBase64) {
      var it = DriveApp.getFoldersByName('Resumes');
      var folder = it.hasNext() ? it.next() : DriveApp.createFolder('Resumes');
      var safeName = ((data.firstName || 'applicant') + '-' + (data.lastName || '')).replace(/\s+/g, '_');
      var ext = (String(data.resumeName).match(/\.[^.]+$/) || [''])[0];
      var blob = Utilities.newBlob(Utilities.base64Decode(data.resumeBase64),
        data.resumeType || 'application/octet-stream',
        safeName + '-resume' + ext);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      resumeLink = file.getUrl();
    }

    // Columns = every field the form sent (minus raw file data), always plus resumeLink & ReceivedAt
    var columns = Object.keys(data).filter(function (k) {
      return k !== 'resumeBase64' && k !== 'resumeType';
    }).concat(['resumeLink', 'ReceivedAt']);

    var lastCol = sheet.getLastColumn();
    var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

    // Auto-add any columns that don't exist yet (e.g., new questions added later)
    var additions = columns.filter(function (c) { return headers.indexOf(c) === -1; });
    if (additions.length) {
      if (lastCol === 0) {
        sheet.getRange(1, 1, 1, additions.length).setValues([additions]);
      } else {
        sheet.getRange(1, lastCol + 1, 1, additions.length).setValues([additions]);
      }
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }

    var row = headers.map(function (h) {
      if (h === 'ReceivedAt') return new Date();
      if (h === 'resumeLink') return resumeLink;
      var v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    if (row.length) sheet.appendRow(row);

    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

### Step 3 — Deploy

1. **Deploy → New deployment** → gear icon → **Web app**
2. Execute as: **Me**  |  Who has access: **Anyone**
3. Deploy → authorize (the "unverified app" warning is your own script:
   **Advanced → Go to project (unsafe)**).
4. Copy the Web App URL (ends in `/exec`) and put it in `index.html` as
   `APP_SCRIPT_URL` in the `<script>` section, then commit & push.

## Updating the script after a change (important!)

Deployments snapshot the code. After editing the script you MUST publish it:

**Deploy → Manage deployments → ✏️ (edit) → Version: New version → Deploy**

The URL stays the same. If you skip this, the old code keeps running.

## Migrating an existing sheet (already-deployed case)

If columns were created by an earlier test (e.g., only firstName/lastName):
1. Right-click the **Applications** tab → **Delete**
   (or just delete the header row 1)
2. Redeploy a new version (above)
3. Submit another test applicant — all columns (email, phone, position,
   A1–A10, LQ/LO answers, resumeLink, etc.) will be created automatically.

## Verify a resume link

Attach a PDF when test-submitting. After the row appears, open `resumeLink` —
it points to the uploaded file in the Drive «Resumes» folder (view access via
link for recruiters).
