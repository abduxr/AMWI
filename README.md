## Sheet Automation (MW Email Digest)

Sends you a daily email listing **new or updated** MW rows (filtered for your name) from an Excel sheet.

If your sheet is “dynamic in Chrome” (Google Sheets), this supports reading **directly from Google Sheets** so you always get the latest data.
If your sheet is an Excel file hosted on **SharePoint/OneDrive** (like `cisco-my.sharepoint.com`), this supports downloading it via **Microsoft Graph**.

### Setup

- **Install Python**: Python 3.10+ recommended.
- **Create venv + install deps**

```bash
cd "/Users/abdun/Documents/Sheet_Automation"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Configure

1. Copy the config template:

```bash
cp config.example.toml config.toml
```

2. Edit `config.toml`:
   - **source.type**: choose `excel` or `google_sheets`
   - If `excel`:
     - **excel.path**: absolute path to your Excel file
   - **filter.assignee_name**: must match the name in the sheet (for you: `N Abdullah`)
   - **filter.col_* names**: update to match your Excel column headers (if different)
   - **smtp**: put your SMTP host/user/password (for Gmail, use an **App Password**)

### If your sheet is Google Sheets (Chrome)

1. Set in `config.toml`:
   - `source.type = "google_sheets"`
   - `google_sheets.spreadsheet_id = "..."` (from the URL)
   - `google_sheets.service_account_json = "/path/to/service-account.json"`
   - optionally `google_sheets.worksheet = "TabName"`

2. Create a Google Cloud **service account** and download its JSON key.

3. Share the Google Sheet with the **service account email** (ends with `iam.gserviceaccount.com`) as a viewer/editor.

### If your sheet is SharePoint / OneDrive (Chrome link)

Your link looks like a SharePoint-hosted Excel file. To automate it daily, we download the workbook using Microsoft Graph.

1. In `config.toml`, set:
   - `source.type = "sharepoint"`
   - `sharepoint.share_url = "https://cisco-my.sharepoint.com/:x:/p/..."`
   - `sharepoint.tenant_id`, `sharepoint.client_id`, `sharepoint.client_secret`

2. Create an Azure AD App Registration (one-time):
   - Add **Application** permission: `Sites.Read.All`
   - Click **Grant admin consent**
   - Create a **Client secret**

3. Run the script once to test.

### If you can’t create Azure apps (employee / personal)

Use `sharepoint_browser` mode. This downloads the workbook using your normal SharePoint login in a browser.

1. In `config.toml`:
   - `source.type = "sharepoint_browser"`
   - `sharepoint_browser.share_url = "https://cisco-my.sharepoint.com/:x:/p/..."`
   - `sharepoint_browser.storage_state_path = "/Users/abdun/Documents/Sheet_Automation/sharepoint.storage.json"`

2. Install Playwright’s browser (one-time):

```bash
source .venv/bin/activate
playwright install chromium
```

3. One-time login (this opens a real browser window; you sign in as usual):

```bash
source .venv/bin/activate
python sheet_notifier.py --config config.toml --init-sharepoint-auth
```

4. Now normal runs will download headlessly:

```bash
source .venv/bin/activate
python sheet_notifier.py --config config.toml
```

### Run once (test)

```bash
source .venv/bin/activate
python sheet_notifier.py --config config.toml --send-if-no-changes
```

It creates/updates `.state.json` to track what changed between runs.

### Schedule daily at 6:10 AM IST (macOS launchd)

1. Create a LaunchAgent file at `~/Library/LaunchAgents/com.abdullah.sheet-notifier.plist` with content:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.abdullah.sheet-notifier</string>

  <key>ProgramArguments</key>
  <array>
    <string>/Users/abdun/Documents/Sheet_Automation/.venv/bin/python</string>
    <string>/Users/abdun/Documents/Sheet_Automation/sheet_notifier.py</string>
    <string>--config</string>
    <string>/Users/abdun/Documents/Sheet_Automation/config.toml</string>
  </array>

  <key>StartCalendarInterval</key>
  <dict>
    <key>Hour</key><integer>6</integer>
    <key>Minute</key><integer>10</integer>
  </dict>

  <key>StandardOutPath</key>
  <string>/Users/abdun/Documents/Sheet_Automation/launchd.out.log</string>
  <key>StandardErrorPath</key>
  <string>/Users/abdun/Documents/Sheet_Automation/launchd.err.log</string>

  <key>RunAtLoad</key>
  <false/>
</dict>
</plist>
```

2. Load it:

```bash
launchctl load -w ~/Library/LaunchAgents/com.abdullah.sheet-notifier.plist
```

If your Mac time zone is set to **India Standard Time**, it will run at 6:10 AM IST.

