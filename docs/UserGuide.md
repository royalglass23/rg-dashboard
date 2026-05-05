# Royal Glass Work Order Dashboard
## User Guide

**Last updated:** May 2026
**For:** Roxy and office staff

---

## What Is This?

The Royal Glass Work Order Dashboard is a Google Sheets-based tool that automatically pulls your active jobs and materials from ServiceM8 and organises them in one place.

Instead of jumping between ServiceM8 tabs to track what's happening with each job, you can see everything at a glance — what jobs are active, what materials are needed, who is installing, what stage each item is at, and whether the site is ready.

The dashboard also uses AI to generate short, easy-to-read labels for each material item so your team can quickly identify what they are looking at without reading long ServiceM8 descriptions.

---

## What It Is Not

- It does not replace ServiceM8 — it reads from it
- It does not create jobs or invoices
- It does not send emails or notifications
- Changes made in this dashboard do not update ServiceM8

---

## The Two Main Sheets

### Work Orders
One row per active job. Shows the job number, client, address, description, date, and status.

You only fill in the **Status** column — everything else comes from ServiceM8 automatically.

### Job Materials
One row per item within each job. Shows what materials or products are involved and lets your team track the production status of each one.

Columns A to F come from ServiceM8 automatically.
Columns H to Q are for your team to fill in manually.

---

## Column Guide — Job Materials

| Column | What it is | Who fills it |
|---|---|---|
| Job No | Job number from SM8 | Auto |
| Client Name | Client name from SM8 | Auto |
| Site Address | Job address from SM8 | Auto |
| Date Created | When the job was created | Auto |
| Project Description | Job description from SM8 | Auto |
| Item - Short Label | AI-generated summary of the item | Auto (AI) |
| Installer | Who is doing the install | Staff |
| Product Type | Type of glass product | Staff |
| Design Status | Whether design is confirmed | Staff |
| Expected Completion Date | When it should be done | Staff |
| Confirmed Date | Actual confirmed date | Staff |
| Glass Status | Status of glass order | Staff |
| Hardware Status | Status of hardware order | Staff |
| Site Ready | Tick when site is ready | Staff |
| Overall Status | Overall job item status | Staff |
| Remarks | Any notes or comments | Staff |

---

## The Menu

Click **🔧 Royal Glass** in the top menu bar to see all options.

### Refresh Options

| Menu item | When to use it |
|---|---|
| 🔄 Refresh — Last 7 Days | Every morning. Updates jobs from the last 7 days. |
| 🔄 Refresh — Last 30 Days | After a long weekend or if something looks outdated. |
| 🗂️ Full Refresh (All Jobs) | If you suspect something is missing or after a reset. |

**Refresh does the following automatically:**
- Pulls updated job and material info from ServiceM8
- Removes jobs that are no longer active
- Removes items that have been invoiced or marked as deposit
- Generates short labels for any new items

### Backup Options

| Menu item | What it does |
|---|---|
| 💾 Backup Data | Saves a snapshot of both sheets right now |
| 📂 Load Backup | Restores a previous snapshot |
| 📋 List Backups | Shows what backups are available |

---

## How to Use It Day to Day

### Morning routine
1. Open the dashboard
2. Click **🔧 Royal Glass → Refresh — Last 7 Days**
3. Wait for the green tick toast at the bottom right
4. Review new or updated rows
5. Fill in any manual columns that need updating

### Updating a job item
1. Find the row in **Job Materials**
2. Click the cell you want to update (Installer, Glass Status, etc.)
3. Select from the dropdown or type a value
4. The change is saved automatically and recorded in the Audit Log

### Checking what a label means
Hover your mouse over any cell in the **Item - Short Label** column.
A small tooltip will appear showing the full original item name from ServiceM8.

### If a label looks wrong
1. Click the cell in column G (Item - Short Label) that has the wrong label
2. Press Delete to clear it
3. Run any refresh
4. A new label will be generated automatically

---

## Item - Short Label Format

Labels follow this format:

```
[Product Name] | L: [length] | H: [height] | [Finish]
```

**Examples:**

| What ServiceM8 says | What the label shows |
|---|---|
| Mini Post System 14.6m chrome internal stair | Mini Post System \| L: 14.6m \| Chrome |
| Shower Glass 8 hinged doors 1960mm chrome | Shower Glass ×8 \| H: 1960mm \| Chrome |
| Glass Splashback 11.6sqm polished | Glass Splashback \| 11.6sqm \| Polished |

If dimensions are unknown, the label will show **TBC**.

---

## Frequently Asked Questions

**Why is a job missing from the dashboard?**
The dashboard only shows jobs with the status "Work Order" in ServiceM8. If a job has moved to a different status (e.g. Quote, Completed), it will be removed from the dashboard automatically on the next refresh.

**Why does an item have no label?**
The item name from ServiceM8 may be blank, or the label is still generating. Run a refresh and check again. If it still has no label, the item name in ServiceM8 may be empty.

**Why is there a row with no material items?**
Some jobs in ServiceM8 have no materials attached yet. The dashboard still shows the job as a placeholder row so you know it exists.

**Why did a row disappear?**
The job or item is no longer active in ServiceM8, or it was an invoice/deposit item which is excluded automatically.

**Can I delete a row manually?**
No. Do not delete rows manually. The next refresh will either bring it back or remove it based on what is in ServiceM8.

**What does "Other" mean in the dropdowns?**
If you select "Other" in any dropdown, the cell clears and lets you type a custom value. This is for situations where none of the standard options fit.

**Why does the Status column on Work Orders not have an "Other" option?**
The Work Orders status is intentional — only the four standard values are allowed to keep reporting consistent.

**Can I undo a change I made to a manual column?**
Not directly through the dashboard. Contact Mike — the Audit Log records every change including the old and new value, so it can be reviewed and corrected.

---

## Troubleshooting

### The menu is not showing
Close and reopen the spreadsheet. The menu loads when the file opens.

### Refresh is not working
Check your internet connection. If it fails, a pop-up will tell you what went wrong. If it says "SM8_API_KEY not found", contact Mike.

### A label has not appeared after refreshing
Wait 30 seconds and refresh again. Label generation can sometimes be slow when there are many new items.

### The sheet looks empty after a refresh
This can happen if a Full Refresh ran with no active Work Orders matching the filter. Check ServiceM8 to confirm jobs have "Work Order" status.

### I accidentally edited something I should not have
Do not try to fix it yourself. Contact Mike — the Audit Log has a full record of what was changed and can be used to restore the correct value.

### Something looks wrong but I am not sure what
Take a screenshot and send it to Mike with a description of what you expected to see.

---

## Who to Contact

For anything related to this dashboard:
**Mike** — AI Programmer, Royal Glass

---

## Quick Reference Card

Print this section and keep it at your desk.

**Every morning:** 🔧 Royal Glass → Refresh Last 7 Days

**After a long break:** 🔧 Royal Glass → Full Refresh

**Before major changes:** 🔧 Royal Glass → Backup Data

**Label looks wrong:** Clear the cell in column G → run any refresh

**Something broken:** Screenshot it → contact Mike
