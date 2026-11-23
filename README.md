# Finance Tracker

Finance Tracker runs on [GitHub Pages](https://dtejada3.github.io/FinTracker/home/). Its Google Drive mode stores each person's ledger and receipt files in their own Drive. The Apps Script instructions below apply to the legacy connection and to exporting existing receipt images for migration.

## Enable Google Drive sign-in

1. In the same Google Cloud project, enable the **Google Drive API** and the **Google Sheets API**. Configure the Google Auth Platform consent screen and add the accounts that will test the app to its test audience.
2. Create a **Web application** OAuth client. Add `https://dtejada3.github.io` under **Authorized JavaScript origins**. Copy its public client ID into `webClientId` in [finance-google.json](finance-google.json). The browser requests `https://www.googleapis.com/auth/drive.file` and keeps its short-lived access token in memory. [Google's client setup guide](https://developers.google.com/workspace/guides/create-credentials)
3. For the Android APK, create an **Android** OAuth client in that project for `com.example.financetracker` using the SHA-1 of the certificate that signs the APK. A debug signing certificate and a release signing certificate typically require separate Android clients. Build and install the updated wrapper for native authorization.
4. If the legacy ledger has uploaded photos, deploy the updated [Code.gs](apps-script/Code.gs) to the existing Apps Script deployment first. Its export endpoint reads images from the tracker’s original Drive folder so the migration can copy them into the user's Drive.
5. Set `enabled` to `true` in `finance-google.json` and publish to GitHub Pages. Test with the same Google account in a browser and the APK before using the migration with live records.

### Move an existing ledger

Open the app on a device that still has the legacy connection. Choose **Sign in and migrate** and select the Google account that should own the ledger. The app copies the Sheet and pending device edits, checks the copy, then clears the old connection on that device. Repeat on the other devices using the same account. Keep your original Sheet as a reference while you verify the import.

The app looks for **Financial Tracker Account** and **Finance Tracker Ledger** among the files its Google Drive permission can access. The spreadsheet's `Ledger` tab stores individual changes so simultaneous device writes can be read back without replacing the entire Sheet.

## Set up the legacy Apps Script web app

1. Open the target spreadsheet and choose **Extensions → Apps Script**. Paste [Code.gs](apps-script/Code.gs) into the script file. Create an HTML file named `index` and paste [index.html](apps-script/index.html) into it. If the project already has an HTML file that serves the tracker, update that file.
2. In **Project Settings → Script Properties**, add `API_KEY` with a new, long random value. Set the project time zone to match your spreadsheet. The bundled [manifest](apps-script/appsscript.json) uses `America/Chicago`.
3. Select `setup` in the Apps Script editor, click **Run**, and authorize the requested Google permissions. Setup registers the bound spreadsheet and creates the tracker tabs. Use a dedicated spreadsheet for a new installation.
4. Choose **Deploy → New deployment → Web app**. Set **Execute as** to **Me** and select **Only myself** for personal use. For an existing deployment, edit it under **Deploy → Manage deployments** and select **New version** to keep its URL.
5. Open the deployed URL ending in `/exec` and enter the access key. Create a test transaction and reload to confirm it syncs.

Receipt images uploaded by the app are stored on Drive and shared with anyone who has the individual image link.

## Update with clasp

Enable the Apps Script API in your [Apps Script settings](https://script.google.com/home/usersettings). Copy `.clasp.example.json` to `.clasp.json` and set `scriptId` to the **Script ID** shown under **Project Settings** in the bound project. Adjust the local manifest to match the target project’s time zone and settings.

```sh
npm install -g @google/clasp
clasp login
clasp show-file-status
clasp push
```

`clasp push` replaces the project’s source files. Back up any existing files and reconcile the manifest before using it on an existing project. Then finish the Script Properties, setup, and deployment steps above. [Google’s clasp documentation](https://github.com/google/clasp)
