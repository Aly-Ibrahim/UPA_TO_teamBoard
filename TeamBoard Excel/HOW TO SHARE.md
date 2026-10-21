# TeamBoard.xlsx: setup and sharing

One Excel workbook for the whole team: projects, tasks, a Kanban board, dashboards, scorecards and a live "Room". It has **no macros**, so it works for free in **Excel for the web**, and everyone can edit it at the same time.

## What's inside

| Sheet | What it does |
|---|---|
| **Home** | Start page with clickable tiles for every view, plus live headline numbers |
| **Dashboard** | 7 KPI cards and 5 charts: completion ring, status donut, priority bars, workload per person, weekly pace. Also a "Needs attention" list, project health and a who's-working-on-what heatmap. Click the **slicers** to filter everything. |
| **Board** | Kanban lanes (To Do → In Progress → Review → Blocked → Done). Cards have a color strip for priority, the owner, due date, project code, tags and % done. Overdue cards turn pink. There's also a search box and slicers. |
| **Room** | Live activity feed, newest first: 🆕 new tasks, ✅ completed work, 🚀 new projects, 💬 messages. It also has a "Today" pulse bar, the newest project and the latest task. Post messages in the table on the right. |
| **My Work** | Pick your name to see your open tasks sorted by due date, what you finished recently, and your score and rank |
| **Calendar** | Month grid of due dates (⚠ overdue · ✓ done · • open) plus a "next 14 days" agenda. Pick the month and year at the top. |
| **Scorecards** | Team leaderboard with score, ★ rating, rank and 🥇🥈🥉 badges, plus a project scorecard with health: ✅ On Track · ⚠ At Risk · ⛔ Off Track · 🏁 Complete |
| **Tasks / Projects / Team** | The data. **Blue headers = you type**, grey headers = automatic |
| **Settings** | Rename the 5 statuses (board lanes) and the 4 priorities |

## Share it with your team (free)

1. Go to **onedrive.live.com** (a personal Microsoft account works) or your company's OneDrive or SharePoint.
2. **Upload** `TeamBoard.xlsx`.
3. Open the file and click **Share**. Choose **Can edit** and add your teammates' emails. They can open it in the browser without buying Office.
4. Optional: in Microsoft Teams, open your channel and click **+ (Add a tab) → Excel**, then pick the file. It then sits inside Teams as the team's board.

Everyone's edits appear for everyone else within a few seconds, and you see who's in the file at the top right. That's what makes the Room work as a live activity feed.

## Daily use

- **Add a task:** go to the **Tasks** sheet and type in the first empty row under the table. The table grows by itself and the ID is filled in automatically. Fill in **Added By**, and set **Added On** with **Ctrl + ;** (today's date) so the task shows up in the Room feed.
- **Move a card:** change the task's **Status**. The card jumps to the new lane on the Board. When a task is Done, set **Completed On** with **Ctrl + ;**.
- **New project:** add a row on **Projects** and give it a short **Code** (like WEB). It's announced in the Room.
- **Post a message:** add a row to the table on the **Room** sheet. For **When**, press **Ctrl + ;**, then **space**, then **Ctrl + Shift + ;** to enter the date and time.

## Good to know

- **Slicers and the My Work name are shared.** If one person filters or picks a name, everyone who has the file open sees that same view. Clear your slicers when you're done (the ⨯ on each slicer).
- **Dates you type yourself.** Excel can't stamp "Added On" automatically without macros. The Ctrl + ; shortcut takes one second.
- **Keep the Tasks table in order.** Add new tasks at the bottom rather than inserting or sorting in the middle, so task IDs stay the same.
- **Sample data:** the 10 people, 6 projects, 40 tasks and 8 messages are examples. Overwrite or delete those rows, but keep the tables and their header rows.
- **Needs** Excel for the web (free) or the Microsoft 365 Excel app. Older desktop versions (Excel 2019 and earlier) don't support the live formulas this workbook uses.
- **Limits:** each Board lane shows up to 25 cards (the lane header always shows the full count), and the Room feed shows the newest 40 events.
- `build_teamboard.py` is the script that generated the workbook. You don't need it to use the workbook.
