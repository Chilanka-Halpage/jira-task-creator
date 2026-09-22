# Jira Task Creator

A small web app that turns a sprint-backlog Excel sheet into Jira tasks.
Because the sheet's `Project` / `Component` / `Assignee` values are free text
that doesn't match Jira's real project keys, components or users, the app has
a one-time **metadata mapping** step: you tell it once "Excel value X = Jira
project/component/user Y", it's remembered, and every future import reuses
it (only brand-new values need mapping).

Runs two ways:
- **Locally**, for yourself — `run.bat`/`run.sh` remember your Jira
  connection and mapping across restarts (saved to `./data` on your own machine).
- **Deployed for a team to share** — see [Deploying for others to use](#deploying-for-others-to-use-free)
  below. Each visitor's Jira credentials and mapping stay private to their own
  browser session; nothing is written to shared server disk.

## Requirements

- Python 3.9+ (already installed on this machine)
- A Jira Cloud site + an [API token](https://id.atlassian.com/manage-profile/security/api-tokens)

## Run it

**Windows:** double-click `run.bat` (or run it from a terminal).
**Mac/Linux:** `./run.sh`

Either script creates a virtual environment, installs dependencies, and
launches the app in your browser at `http://localhost:8501`. Subsequent runs
reuse the same environment, so startup is fast.

Manual alternative:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## How to use it

1. **Connect to Jira** — enter your site URL (`https://yourcompany.atlassian.net`),
   your account email, and an API token. Click *Save & Test Connection*.
2. **Upload Sheet** — upload the `.xlsx` file and pick the sheet to import
   (it defaults to a sheet named like "Sprint Backlog"). Click *Parse sheet*.
3. **Configure Mapping**
   - **Project mapping** — map each distinct Excel `Project` value to a real Jira project.
   - **Component mapping** — for each Jira project actually used in the current
     sheet, pick the default issue type (e.g. "Task") and map the Excel
     `Component` values seen under that project to real Jira components.
     (Projects mapped previously but not present in this sheet aren't shown.)
   - **Assignee mapping** — map each distinct Excel `Assignee` name to a real Jira user.
   - **Task title format** — per project, build the Jira task title from literal
     text plus Excel columns as `{Placeholder}` tokens (pick a column and click
     *Insert*, or type placeholders directly), with a live preview. Example:
     `[NTP-SAAS][{Component}] {Description}` →
     `[NTP-SAAS][AT] Production/UAT Support - AT - Resource 1`. Leave it as the
     default `{Description}` for plain behavior.
   - **Task labels** — per project, pick labels to attach to every created
     task from a dropdown of labels already used in your Jira site, or type a
     new one and press enter to add it. No labels are added unless you
     configure some here.
   - **Story linking** — pick the Jira **link type** used to connect a new task to its
     `Jira Story` reference (e.g. "relates to"). Whenever a row's `Jira Story`
     column holds a real-looking issue key (like `NTPSFC-16683`), the created
     task is linked to that story automatically.
   - **Additional field mappings** — map any other Jira field, per project,
     either:
     - **from an Excel column** (Capitalization, Priority, Status, Comments,
       Section — anything not already handled above); for a dropdown-style
       Jira field the matching option is selected automatically (e.g. `75` →
       the `75%` option), for a plain text/number field the value is set
       as-is; or
     - **to a fixed value used on every task** — for a field Jira's create
       screen *requires* that has no corresponding Excel column at all (e.g.
       a mandatory "Payment Category" dropdown). Without this, task creation
       fails outright with an error like `"Payment Category is required"`.
     Already-mapped fields are excluded from the picker as you add more.
   - Mappings persist for as long as your session lasts (across a local run,
     that's until you close `run.bat`; on a shared deployment, that's until
     you close the browser tab) — re-uploading a new sprint sheet later will
     only ask you to map values it hasn't seen before, within that same
     session. Use **Export mapping config** / **Import mapping config** at
     the top of this tab to save your mapping to a file and reload it later
     or on another device — this works the same way in both modes.
4. **Review & Create** — check the resolved table (skips rows that already
   have a JIRA ID or a status like *Done*/*Discarded* by default). Untick
   **Create?** on any row in the table to skip it individually, even if it's
   otherwise ready — nothing is created for an unticked row regardless of the
   filters above. Use **☑️ Select All** / **⬜ Unselect All** above the table
   to quickly tick or clear every row at once. If any project has a
   **fixed-value** field mapped, an
   editable **Override Fixed Field Values** grid appears — edit any cell to
   use a different value for just that one task, without changing the
   project-wide default. Tick the confirmation box, and click *Create Jira
   tasks*. Results (created issue keys / errors) are shown and can be
   downloaded as CSV, and you can download the original Excel with the new
   JIRA IDs filled in. If any tasks failed, a **🔁 Retry failed task(s)**
   button appears — click it to re-attempt only the failed rows (e.g. after
   fixing a mapping issue); successfully created tasks are left untouched.

## Dark mode

Toggle **🌙 Dark mode** at the bottom of the sidebar. It's a session-only
preference (resets to light on next launch). The data tables' grid contents
are rendered on a canvas internally and may keep a light background even
with dark mode on — a Streamlit rendering limitation, not a bug.

## Stopping the app

Only available when running locally via `run.bat`/`run.sh` (hidden on a
shared deployment, since it would take the app down for every other visitor
at once). Click **🛑 Stop App** at the bottom of the sidebar and confirm —
this shuts down the local server immediately (equivalent to closing the
terminal window running `run.bat`). It also tries to close the browser tab
for you, but browsers only allow a page to close itself if a script opened it
in the first place — for a normally-opened tab this will usually be silently
blocked, so you may still need to close the tab by hand. Run `run.bat` again
to restart.

## Deploying for others to use (free)

**GitHub Pages won't work for this** — it only serves static files, and this
app needs a Python backend (to parse Excel, call Jira's API, etc.) plus Jira
Cloud blocks direct browser-to-API calls from other domains (CORS), so a
static page could never reach Jira anyway.

**[Streamlit Community Cloud](https://share.streamlit.io)** is the free,
purpose-built host for apps like this one:

1. Push this repo to GitHub (public or private both work):
   ```bash
   git remote add origin https://github.com/<you>/<repo-name>.git
   git push -u origin master
   ```
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click **New app**, pick this repo/branch, and set the main file
   path to `app.py`. Deploy.
3. **Do not** set `JIRA_APP_PERSIST_LOCAL` or `JIRA_APP_ALLOW_STOP` as
   secrets/environment variables on the deployed app — leaving them unset is
   what keeps every visitor's Jira credentials private to their own session
   (see below) and keeps the Stop button hidden.

Each visitor connects with their own Jira URL/email/API token, right there in
the app — nothing needs to be pre-configured before sharing the link.

On a shared deployment, the sidebar also shows a **🔒 Forget my Jira
credentials** button — click it to immediately wipe your token from that
session's memory the moment you're done, rather than waiting for Streamlit's
own idle timeout to clear it. Your uploaded sheet and mapping are kept; only
the Jira connection (and anything fetched using it — projects, components,
etc.) is cleared, and you'll need to reconnect to keep using the app.

## Notes

- **Local runs** (`run.bat`/`run.sh`) store your Jira URL/email/API token in
  plain text in `data/connection.json` so you don't have to re-enter them
  every time — keep that folder private and don't commit it to source control
  (already covered by `.gitignore`). **Shared deployments** (Streamlit
  Community Cloud) never write this to disk at all — it lives only in that
  visitor's browser session and is gone when they close the tab, so different
  visitors never see or overwrite each other's credentials. The Connect tab
  never re-displays a saved token in the field either way (it stays blank
  with a "saved" placeholder) so the real secret isn't sitting in the page
  each time you open the tab — leave it blank to keep the saved token, or
  type a new one to replace it.
- New issues are created with whatever status is the default for their
  workflow; the sheet's `Status` column is only used as an *exclude* filter
  during import, it does not transition the created issue.
- If a project's Jira "Create issue" screen doesn't have a field the app
  tries to set (e.g. Components), the app automatically drops that field and
  retries rather than failing the whole row — check the *note* column in the
  results table.
- Story linking only fires when `Jira Story` looks like a real issue key
  (letters, then `-`, then digits — e.g. `DKALKC-3301`). Values like `N/A` are
  ignored. If the key doesn't actually exist in Jira, the task is still
  created but the link attempt fails and is reported in the *note* column.
- Additional field mappings need a default issue type to already be set for
  that project (Component mapping), since Jira's available fields are scoped
  to a project *and* issue type. If a sheet's value has no matching dropdown
  option in Jira, the task is still created — it's just reported in the *note*
  column.
