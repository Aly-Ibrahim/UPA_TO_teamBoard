# TeamBoard — free Trello + ClickUp on Google Sites

**What you get:** Kanban boards with drag & drop (Trello), plus List, Calendar, Dashboard and "My Work" views (ClickUp). Tasks have assignees, due dates, priorities, tags, checklists and comments. People get an email when a task is assigned to them. It runs on Google's free tools: **Sheets** stores the data, **Apps Script** runs the app, and **Sites** is the team homepage.

Cost: **$0**. A free Gmail account or Google Workspace is enough for a team of 10.

---

## Step 1 — Create the database (2 min)

1. Go to **sheets.new** and name the sheet **TeamBoard Data**.
2. Click **Extensions → Apps Script**. The script editor opens in a new tab.

## Step 2 — Paste the code (3 min)

1. In the editor, select everything in `Code.gs`, delete it, and paste the contents of **Code.gs** from this folder.
2. Click **+** next to *Files* → **HTML**, and name it exactly `Index` (the editor adds `.html`).
3. Paste the contents of **Index.html** into it.
4. Click **💾 Save**.

## Step 3 — Run setup once (1 min)

1. In the toolbar's function dropdown, pick **setup**, then click **▶ Run**.
2. Google asks for permission. Click **Review permissions**, choose your account, then **Advanced → Go to (unsafe)** → **Allow**. You see the "unsafe" warning because you wrote the script yourself and Google hasn't reviewed it. That's normal.
3. Go back to the Sheet. You now have the tabs **Boards, Tasks, Comments, Members, Activity**, a starter board called *General*, and you listed as a member.

## Step 4 — Add your 10 team members

Open the **Members** tab and add one row per person:

| email | name | role |
|---|---|---|
| sara@company.com | Sara Ali | member |

You can also add people later from the **Team** page inside the app.

## Step 5 — Publish as a web app (2 min)

1. In Apps Script, click **Deploy → New deployment**, then the gear icon → **Web app**.
2. Set:
   - **Execute as:** *Me*. The app reads and writes the Sheet for everyone, so you **don't** need to share the Sheet with the team.
   - **Who has access:** *Anyone within your organization* if you use Google Workspace, or *Anyone with a Google account* if you use Gmail.
3. Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).

## Step 6 — Build the Google Site (5 min)

1. Go to **sites.new**.
2. Name it, for example *Team Hub*.
3. Click **Insert → Embed → By URL**, paste the `/exec` URL, and choose **Whole page**. Drag the box to full width and as tall as possible.
4. Optional extra pages, all free and found under **Insert**:
   - **Calendar:** a shared Google Calendar for meetings and time off.
   - **Drive:** a shared folder for team files.
   - **Docs:** a meeting-notes doc or team wiki.
   - **Forms:** request intake. Answers land in a Sheet.
5. Click **Share** (person icon with +), add your team's emails, then **Publish**. Under *Who can view my published site*, choose *Specific people* or your organization.

Done. The team opens the Site and works on the boards.

---

## How people are identified

- **Google Workspace (same company domain):** the app detects each person's email automatically.
- **Gmail accounts:** Google hides other people's emails from the script, so the app asks **"Who are you?"** once per device and remembers the answer. Use **switch user** at the bottom-left to change it.

## Updating the app later

After you edit the code: **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. The URL stays the same, so the Site doesn't need changes.

## Customizing

- **Columns:** click ⚙ next to a board's name. The last column counts as "done".
- **Turn off emails:** in `Code.gs`, set `NOTIFY_BY_EMAIL = false`.
- **Backups:** in the Sheet, use **File → Version history**. You can also run **File → Make a copy** each month.
- **Reports:** the **Activity** tab logs who created, moved, updated, deleted and commented on what. You can build pivot tables or charts from it.

## Free-tier limits (plenty for 10 people)

| Limit | Gmail | Workspace |
|---|---|---|
| Assignment emails per day | 100 | 1,500 |
| Script runtime per day | 90 min | 6 h |
| Rows in the Sheet | 10 million cells | same |

Realistic capacity: several thousand tasks. If the app slows down, archive old "Done" tasks by moving those rows to a separate sheet.

## Preview without Google

Double-click **Index.html** to open it in a browser. It runs in **demo mode** with sample data, which is saved only in that browser.
