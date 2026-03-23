# Google API Setup

Run this once before using `/weekly-task` or `/weekly-report`.

## 1. Enable APIs

In [Google Cloud Console](https://console.cloud.google.com/):
1. Select (or create) your project
2. Go to **APIs & Services → Library**
3. Enable **Google Docs API**
4. Enable **Google Drive API**

## 2. Create a Service Account

1. Go to **APIs & Services → Credentials**
2. Click **Create Credentials → Service Account**
3. Name it (e.g., `weekly-report-bot`), click **Create and Continue**
4. Skip optional role/user steps, click **Done**
5. Click the service account → **Keys → Add Key → Create new key → JSON**
6. Download the JSON file, save it somewhere safe (e.g., `~/.config/weekly-report-credentials.json`)

## 3. Share your Drive folder with the service account

The target folder can be a personal Drive folder or a **Shared Drive** folder — both are supported.

1. Create (or identify) the Drive folder where reports should be saved
2. Right-click the folder → **Share**
3. Enter the service account email (found in the JSON file under `"client_email"`)
4. Grant **Content manager** (or **Editor**) access
5. Copy the folder ID from the URL:
   - Personal Drive: `https://drive.google.com/drive/folders/{FOLDER_ID}`
   - Shared Drive: `https://drive.google.com/drive/u/0/folders/{FOLDER_ID}`

> **Shared Drive note:** The helper scripts use `supportsAllDrives=True` in all Drive API calls, so Shared Drive folders work without any extra configuration.

## 4. Set environment variables

Add to your shell profile (`~/.zshrc` or `~/.bashrc`):

```bash
export GOOGLE_APPLICATION_CREDENTIALS="/path/to/your-credentials.json"
export WEEKLY_REPORT_FOLDER_ID="your_drive_folder_id_here"
```

Reload: `source ~/.zshrc`

## 5. Install dependencies

```bash
pip install google-api-python-client google-auth
npm install -g pptxgenjs
```

## Verify

```bash
echo "Hello from weekly-report setup test" > /tmp/test-doc.txt && \
python3 .claude/skills/_shared/helpers/google_docs.py create \
  --title "Test Doc" \
  --content /tmp/test-doc.txt \
  --folder-id $WEEKLY_REPORT_FOLDER_ID
```

Expected: prints a `https://docs.google.com/document/d/...` URL and the doc appears in your Drive folder.
