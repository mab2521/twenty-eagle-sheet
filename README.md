# Twenty Eagle — Case Review Sheet (Google Apps Script)

The Google Sheet side of the Twenty Eagle case pipeline. It receives case data from the CourtPASS scraper, gives reviewers a place to check and annotate each case, and turns approved cases into **WordPress drafts**. It never publishes anything; a person publishes each draft from WordPress.

## Where this fits

```
CourtPASS scraper (Python)  ──►  this Sheet  ──►  WordPress draft  ──►  published page
   courtpass_scraper.py          human review      human review          (done in WP)
   + sheet_sync.py               + "Request JL?"
```

1. **Scraper → Sheet.** `sheet_sync.py` POSTs scraped cases to this script's web app URL. `doPost()` adds new cases as rows and refreshes existing ones (matched on `Case key`).
2. **Human review.** A reviewer reads the row, writes the blog summary, adds links, ticks **Request JL?** if a jurisdictional letter is worth requesting, enters initials, and ticks **Approved**.
3. **Sheet → WordPress.** *Twenty Eagle → Create WordPress drafts for approved rows* creates a draft for every approved row that doesn't have a WP post ID yet.
4. **Human publish.** The reviewer opens the **WP edit link**, checks the draft, and publishes it in WordPress.

## Repository layout

```
src/
  Code.gs            all script logic (menu, sheet setup, sync endpoint, WordPress drafts)
  appsscript.json    project manifest (V8 runtime, New York time zone, web app settings)
.clasp.json.example  template for linking a local copy to your Apps Script project
.gitignore
README.md
```

## The sheet

`Set up sheet` creates a **Cases** tab. Headers are color-coded by who owns the column:

| Color | Owner | Columns | Behavior |
|---|---|---|---|
| Blue | Scraper | Case key, title, decision/argument dates, APL no., category, CourtPASS issues, JI status and docket entries, files available/unavailable, link check, last synced | Overwritten on every sync. Edits trigger a warning. |
| Yellow | Reviewer | Blog summary, decision/video/transcript URLs, Request JL?, reviewer notes, reviewed by, Approved | Never overwritten. The summary and URL columns may be pre-filled by the sync only while blank. |
| Green | Script | WP post ID, WP edit link, WP status | Written when a draft is created. |

The script finds columns by header text, so columns can be reordered freely. Don't rename headers. Re-running `Set up sheet` is safe; it adds any missing columns and leaves data alone.

## Setup

### 1. Add the script to a Google Sheet

**Option A — copy and paste (simplest)**

1. In the Google Sheet, open **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with `src/Code.gs` from this repo.
3. Click the gear icon (**Project Settings**), tick *Show "appsscript.json" manifest file in editor*, and replace that file with `src/appsscript.json`.
4. Save, reload the Sheet, and choose **Twenty Eagle → Set up sheet**. Approve the permission prompt the first time.

**Option B — clasp (keeps the repo and the script in sync)**

```bash
npm install -g @google/clasp
clasp login
cp .clasp.json.example .clasp.json   # then paste your Script ID
clasp push                            # repo → Apps Script
clasp pull                            # Apps Script → repo (if edited in the browser)
```

The Script ID is under **Project Settings → IDs** in the Apps Script editor. `.clasp.json` is git-ignored because each Sheet has its own ID.

### 2. Add Script Properties

In **Project Settings → Script Properties**, add:

| Property | Value |
|---|---|
| `SYNC_TOKEN` | Any long random string. The same value goes in `SHEET_SYNC_TOKEN` on the scraper side. |
| `WP_URL` | The WordPress site address, e.g. `https://example.com` (no trailing slash). |
| `WP_USER` | WordPress username. |
| `WP_APP_PASSWORD` | Created in WordPress under **Users → Profile → Application Passwords**. |
| `WP_POST_TYPE` | Optional. Set to `pages` if each case is a WordPress Page; leave out for Posts. |

**Never put these values in the code or commit them to this repo.**

### 3. Deploy the sync endpoint

1. In Apps Script, choose **Deploy → New deployment → Web app**.
2. *Execute as:* **Me**. *Who has access:* **Anyone** (the `SYNC_TOKEN` check is what keeps strangers out).
3. Copy the web app URL into the scraper's configuration.
4. Open the URL in a browser. You should see `Twenty Eagle sync endpoint is live`.

After editing the code, use **Deploy → Manage deployments → Edit → New version** so the existing URL serves the new code. A brand-new deployment gets a new URL.

### 4. Test the WordPress connection

Choose **Twenty Eagle → Test WordPress connection**. It reports which account it's connected as, where the REST API was found, and whether drafts will be Posts or Pages.

## Sync payload

`doPost()` expects JSON like this:

```json
{
  "token": "<SYNC_TOKEN>",
  "cases": [
    {
      "case_key": "APL-2024-00123",
      "title": "Matter of Clarke v. Board of Education",
      "decision_date": "10/17/2024",
      "argument_date": "09/11/2024",
      "issues_courtpass": "Whether petitioners ... had standing ...",
      "ji_status": "None",
      "blog_summary": "optional pre-fill",
      "decision_url": "optional pre-fill"
    }
  ]
}
```

Keys match the `key` values in the `COLUMNS` list at the top of `Code.gs`. Unknown keys are ignored; missing scraper keys are written as blank. It returns `{ "ok": true, "added": [...], "updated": [...] }` or `{ "ok": false, "error": "..." }`.

## What a draft looks like

Drafts copy the layout of the existing case pages: the blog summary in a wide left column, and a bordered box on the right with **Decision:** (linked to the decision URL) and **Argument:** followed by `vid.` and `tr.` links when those URLs are filled in. Dates are converted to the blog's style (`10/17/2024` → `2024.10.17`). Everything is built from standard WordPress blocks, so it's editable in the block editor. The "Select one…" dropdown on live pages is not generated.

A draft is only created when the row has a title, a blog summary, and reviewer initials. Problems are written to **WP status** as `ERROR: ...` and the other rows continue.

## Safety notes

- Post status is hard-coded to `draft`.
- Scraped text beginning with `=` or `+` is stored as text, so it can't run as a formula.
- The sync uses a script lock, so overlapping syncs don't write over each other.
- Sites hosted on WordPress.com are detected and reported; they don't accept Application Passwords this way.

## Troubleshooting

| Message | Fix |
|---|---|
| `Bad or missing token` | `SYNC_TOKEN` in Script Properties doesn't match the scraper's `SHEET_SYNC_TOKEN`. |
| `Missing column "..."` | Run **Twenty Eagle → Set up sheet**. Check no header was renamed. |
| `Could not find the WordPress REST API` | Check `WP_URL`, and that no security plugin has disabled the REST API. |
| `WordPress 401` | Wrong `WP_USER` or `WP_APP_PASSWORD`, or the host strips the Authorization header. |
| Sync seems to run old code | Deploy a **new version** of the existing deployment (see step 3). |
