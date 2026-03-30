---
name: weekly-task
description: Fetches GitLab MR/issue activity for given users over a date range, groups tasks with Claude analysis, and creates a Google Doc draft for team editing.
disable-model-invocation: false
---

Generates a weekly task draft Google Doc by fetching GitLab activity per user and grouping related work into human-readable task entries.

## Prerequisites

- `glab` CLI authenticated to the project
- Google API configured: see `google-api-setup.md` (same directory as this skill)
- Environment variables set: `GOOGLE_APPLICATION_CREDENTIALS`, `WEEKLY_REPORT_FOLDER_ID`
  - If these are missing, run `source ~/.zshrc` first, then verify with `echo $GOOGLE_APPLICATION_CREDENTIALS`
- Python dependencies are managed automatically via `uv` (no manual install needed)

> **Known limitation:** GitLab API queries are capped at 100 results per request. Weeks with > 100 MRs or issues per user will be silently truncated.

## Arguments

```
$ARGUMENTS format:
  @user1 @user2 [@user3 ...]
  [--this-week]
  [--start YYYY-MM-DD --end YYYY-MM-DD]
```

## Step 1 — Parse arguments and resolve date range

Parse `$ARGUMENTS`:
- Extract usernames (tokens starting with `@`, strip the `@`)
- Detect flags: `--this-week`, `--start`, `--end`

Compute date range:

| Flag | start_date | end_date |
|------|------------|----------|
| (none) | Last Monday (ISO) | Last Sunday (ISO) |
| `--this-week` | This Monday (ISO) | Today (ISO) |
| `--start X --end Y` | X | Y |

Use ISO 8601 format (`YYYY-MM-DDT00:00:00Z`) for GitLab API calls.

Compute ISO week label for the Doc title: `YYYY-WXX` (e.g., `2026-W12`).

## Step 2 — Fetch GitLab data per user

For each username, run the following and apply **client-side date filtering** after fetching:

```bash
# Merged MRs — fetch with server-side hint, filter client-side by merged_at
glab api "projects/:fullpath/merge_requests?author_username={user}&state=merged&merged_after={start_date}T00:00:00Z&merged_before={end_date}T23:59:59Z&per_page=100" 2>/dev/null

# Open MRs (in-progress — no date filter needed)
glab api "projects/:fullpath/merge_requests?author_username={user}&state=opened&per_page=100" 2>/dev/null

# Closed issues — fetch with server-side hint, filter client-side by closed_at
glab api "projects/:fullpath/issues?assignee_username={user}&state=closed&closed_after={start_date}T00:00:00Z&closed_before={end_date}T23:59:59Z&per_page=100" 2>/dev/null
```

> **Note:** The `glab` CLI may append a version update notice as plain text after the JSON output (e.g., "A new version of glab has been released: v1.51.0 → v1.89.0"). Parse only the JSON array — extract the substring from the first `[` to the last `]` before passing to JSON parsing.

> **Important — client-side filtering:** The `merged_after`/`merged_before` and `closed_after`/`closed_before` parameters may not filter correctly depending on the GitLab version. Always apply client-side filtering after parsing:
> - For merged MRs: keep only items where `merged_at[:10]` is between `start_date` and `end_date` (inclusive)
> - For closed issues: keep only items where `closed_at[:10]` is between `start_date` and `end_date` (inclusive)

Collect for each user: merged MRs (with `iid`, `title`, `description`, `web_url`, `labels`, `merged_at`), open MRs (same fields), closed issues (with `iid`, `title`, `web_url`, `labels`, `closed_at`).

## Step 3 — Group and summarize per user

For each user, analyze the collected MRs and issues and group them into feature/task units:

**Grouping rules (apply in order):**
1. MRs linked to the same spec issue (`Related #N` or `Fixes #N` in description) → one group
2. MRs with the same branch prefix (e.g., `model-registry-chunk1`, `model-registry-chunk2`) → one group
3. MRs with the same top-level component in the title (e.g., `feat: model registry ...`) → one group
4. Solo MRs with no clear grouping → one entry each

**Output format per group:**
```
NuFi: {Human-readable feature/task description}
  - {One-line detail} ([#MR_IID](web_url)[, #MR_IID](web_url)])
```

For in-progress (open) MRs, append ` (진행중)` to the description.

**Example:**
```
NuFi: Model Registry 개발
  - 스키마 정의 및 파일 업로드 API 구현 ([#198](url), [#201](url))
  - 목록 조회 버그 수정 ([#195](url))
NuFi: RNGD 발표 준비 (진행중)
  - 발표 자료 작성 ([#210](url))
```

## Step 4 — Generate draft markdown

Assemble the draft using this exact template for each user, in the order they were provided as arguments.

Use the **username** (the argument without `@`) as `{name}` — do NOT look up the GitLab display name.

```markdown
## {username} | Developer

### Done
{grouped task entries from Step 3}

### Demos & Links
<!-- Fill in: links to demos, Figma, screenshots, API docs, etc. -->

### Issues
<!-- Fill in: blockers, risks, or anything that needs attention -->

### Next Week
<!-- Fill in: planned tasks for next week -->

---
```

## Step 5 — Write draft to a temp file and create Google Doc

Write the assembled markdown to `/tmp/weekly-report-draft-{YYYY-WXX}.md`, then run:

```bash
uv run ${CLAUDE_SKILL_DIR}/helpers/google_docs.py create \
  --title "Weekly Report {YYYY-WXX} Draft" \
  --content /tmp/weekly-report-draft-{YYYY-WXX}.md \
  --folder-id $WEEKLY_REPORT_FOLDER_ID
```

If the command exits non-zero, stop and print:
```
Google Docs API call failed — check GOOGLE_APPLICATION_CREDENTIALS and WEEKLY_REPORT_FOLDER_ID.
```

## Step 6 — Output result

```
Draft created: {Google Doc URL}

Please fill in your section in the Doc:
  - Demos & Links
  - Issues
  - Next Week

When ready, run:
  /weekly-report {Google Doc URL}
```
