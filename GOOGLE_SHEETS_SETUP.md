# SPC Inventory → Google Sheets archive setup

The five native Google Sheets files are already created inside the user's Google Drive **SPC INVENTORY** folder:
https://drive.google.com/drive/folders/1qfvthREBHLxX-INm3zk81QwEc2B-fXJN

| Location | Google Sheet |
|---|---|
| Jesup | https://docs.google.com/spreadsheets/d/1g0uOU0kVxzIGZJLX_9nCQE6NVLOx7RTGq8qY2gC7ykQ/edit |
| Richmond Hill | https://docs.google.com/spreadsheets/d/1zNDwUNFh-u6r7uplolt295j2ARCdmLKcobbrwhtjSTQ/edit |
| Ogeechee | https://docs.google.com/spreadsheets/d/1G_-Zjr6ITDZdrnAQ2QE0w8nIqZa6L5RjYO150eKEGcs/edit |
| Claxton | https://docs.google.com/spreadsheets/d/1J7tw4_XpltoDRaIMAGG_7ZXt24Envd8JeRc86GxMih0/edit |
| Blackshear | https://docs.google.com/spreadsheets/d/1-mYDzbtseQikCCqkme4NByKoKeZACw8a6anp3WKxPEs/edit |

## One-time Google authorization (required)
The public static Netlify website cannot directly use ChatGPT's Google Drive connection. A separate Google Apps Script Web App must be deployed **by the account authorized to edit these five files**.

1. Open https://script.google.com/home/projects/create while signed in to Google.
2. Rename the Apps Script project, for example **SPC Inventory Archive**.
3. Open the repository file [google-apps-script.gs](./google-apps-script.gs), copy its entire code, and replace the default Code.gs code with it. Save.
4. Open **Project settings** (gear icon), scroll to **Script properties**, add property name `SPC_SYNC_KEY` with a new private, random string of **at least 24 characters**. Do not publish the value or put it in the GitHub repo.
5. Click **Deploy → New deployment → Web app**. Set **Execute as: Me** and **Who has access: Anyone**; authorize the Google permissions to edit the five inventory sheets. Copy the deployment URL ending in `/exec`.
6. Open https://spc-inventory-app.netlify.app. Expand **Google Sheets archive**, paste the **Web App URL** and **the same private sync key**, then tap **Save connection settings**.
7. Save an actual inventory count, then open that location's Google Sheet and verify a new tab named for the date (e.g. `11/17`) appears. Future same-day counts become `11/17 (2)`, etc.
8. If an archived count does not appear, use **Retry saved counts**. Replays carry the original count ID and will not create duplicate tabs when the server has already archived the same count.

## Important behavior

- A new inventory is **always saved locally first**. Google upload is attempted after the local save.
- The browser sends the Apps Script POST via `no-cors`; a successful network request **cannot prove that Google Sheets accepted the data**. The app intentionally labels submissions as **unverified** and retains IDs for safe retries. Verify the dated tab in the actual spreadsheet.
- A tab displays the item catalog, counted cases (including decimals), supplier SKUs, case pack, category, and blank PAR levels. The suggested-order formula only calculates when a PAR value is entered. PAR values carry forward from the prior dated tab when available.
- The `START HERE` tab is removed after the first successful archived inventory for each location.
- The integration URL and sync key are stored **in that device's browser**, not GitHub. Set up each device separately. Don't use untrusted/shared devices.
- Cloud write authorization is based on a shared secret in the browser; this is an owner-managed prototype, **not enterprise-grade account authentication**. Keep the URL and key private. Google Apps Script access as “Anyone” is required for this static browser approach, so consider a more robust authenticated backend for staff accounts.
- Any existing counts saved before this feature are **not automatically uploaded**. Counts newly completed with this feature are queued and submitted when configured.
- If you revise the script later, redeploy a **new version** through **Deploy → Manage deployments → Edit**. Updating GitHub alone does not update the live Apps Script deployment.
