---
name: weekly-report
description: Reads a filled-in weekly task Google Doc, generates a styled .pptx presentation with pptxgenjs, and uploads it to Google Drive.
disable-model-invocation: false
---

Generates a weekly report `.pptx` from a completed draft Google Doc created by `/weekly-task`, then uploads it to Google Drive.

## Prerequisites

- Google API configured: see `../weekly-task/google-api-setup.md`
- Environment variables set: `GOOGLE_APPLICATION_CREDENTIALS`, `WEEKLY_REPORT_FOLDER_ID`
  - If these are missing, run `source ~/.zshrc` first, then verify with `echo $GOOGLE_APPLICATION_CREDENTIALS`
- Python dependencies are managed automatically via `uv` (no manual install needed)
- Node.js dependency: `npm install -g pptxgenjs`

## Arguments

```
$ARGUMENTS — Google Doc URL (from /weekly-task output)
Example: /weekly-report https://docs.google.com/document/d/1BxiMVs0XRA5...
```

## Step 1 — Read the Google Doc

Run (substituting the skill's base directory for `{SKILL_DIR}`):

```bash
uv run {SKILL_DIR}/helpers/google_docs.py read \
  --url "$ARGUMENTS"
```

If the command exits non-zero, stop and print:
```
Failed to read Google Doc — check permissions and GOOGLE_APPLICATION_CREDENTIALS.
```

Store the output as the doc content for the next step.

## Step 2 — Parse content and build slides data

Analyze the doc content and produce a `/tmp/slides-{week}.json` file following this schema exactly:

```json
{
  "week": "YYYY-WXX",
  "executive_summary": {
    "bullets": ["string — key achievement bullets (max 6); leave empty to auto-derive from completed tasks"]
  },
  "members": [
    {
      "name": "string — member name (lowercase)",
      "role": "string — e.g. Developer",
      "last_week": [
        {
          "project": "string — e.g. NuFi",
          "task": "string — task description",
          "status": "Completed | N% | Blocked",
          "note": "string — optional note or link, empty string if none"
        }
      ],
      "output_summary": [
        {
          "title": "string — achievement title",
          "url": "string — link URL, empty string if none"
        }
      ],
      "this_week": [
        {
          "project": "string",
          "task": "string",
          "est_days": "number",
          "due_date": "YYYY/MM/DD — compute from today + est_days, skip weekends"
        }
      ],
      "issues": ["string — list of issue/blocker descriptions from Issues section; empty array if none"]
    }
  ],
  "team_analysis": {
    "total_completed": "number — count of Completed items across all members",
    "total_planned": "number — count of this_week tasks across all members",
    "risks": [
      {
        "member": "string",
        "description": "string — reason this member is a risk item"
      }
    ],
    "flags": [
      {
        "type": "long_task | blocked",
        "member": "string",
        "task": "string — task name if applicable",
        "est_days": "number — if long_task",
        "note": "string"
      }
    ]
  }
}
```

**Rules for building slides data:**
- `executive_summary.bullets`: pull from "Demos & Links" section highlights or leave empty (JS auto-derives top completed tasks)
- `last_week[].status`: use "Completed" for done tasks; "N%" (e.g. "70%") for in-progress; "Blocked" if blocked
- `risks`: include any member with non-empty Issues section OR Blocked status tasks
- `flags`: include tasks with `est_days > 5`
- `due_date`: calculate from the date `/weekly-report` is run + `est_days` business days

Write the result to `/tmp/slides-{week}.json`.

## Step 3 — Generate .pptx

Extract the week label from the JSON (field `week`). Run (substituting the skill's base directory for `{SKILL_DIR}`):

```bash
node {SKILL_DIR}/helpers/generate_slides.js \
  --data /tmp/slides-{week}.json \
  --output /tmp/weekly-report-{week}.pptx
```

If the command exits non-zero, stop and print:
```
Failed to generate .pptx — check Node.js and pptxgenjs installation.
```

## Step 4 — Upload to Google Drive

```bash
uv run {SKILL_DIR}/helpers/upload_to_drive.py \
  --file /tmp/weekly-report-{week}.pptx \
  --title "Weekly Report {week}" \
  --folder-id $WEEKLY_REPORT_FOLDER_ID
```

If the command exits non-zero, stop and print:
```
Upload failed — check GOOGLE_APPLICATION_CREDENTIALS and WEEKLY_REPORT_FOLDER_ID.
```

## Step 5 — Output result

```
Report created: {Google Drive URL}
```

Clean up temp files: `/tmp/slides-{week}.json`, `/tmp/weekly-report-{week}.pptx`, `/tmp/weekly-report-draft-{week}.md` (if present).
