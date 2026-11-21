# Finance Tracker

Finance Tracker is a Google Apps Script web app backed by a Google Sheet. Apps Script serves the interface and handles sync, receipt images, and PIN reset emails. GitHub stores the source. Open the Apps Script web app URL to use the tracker.

## Set up the web app

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
