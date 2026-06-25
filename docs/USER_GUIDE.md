# User Guide — Bench Dashboard

## What Is This App?

Bench Dashboard is an internal EPAM tool used by delivery managers to track associates who are currently on the bench (between projects). It provides a single view of:

- **Course progress** — which required training courses each associate has completed
- **Client pipeline** — which clients an associate has been proposed to, and the outcome
- **Placement tracking** — when an associate moves off bench (gets selected), they are archived

The app is for **managers only** — associates do not log in or manage their own data. The manager enters and updates all information.

---

## Accessing the App

**URL:** https://saurabhsuman1-droid.github.io/bench-dashboard/

**Via Microsoft Teams:** The app is pinned as a tab in the bench channel. Click the "Bench Dashboard" tab.

**Login:** Click **Sign in with Microsoft** → use your EPAM email and password (same credentials as Outlook, Teams). You will only need to sign in once per browser session.

---

## Views

The header has three mode buttons:

| Button | What it shows |
|---|---|
| **Manager View** | Full editable view — add associates, update courses, log client interactions |
| **Team View** | Read-only summary — clean view suitable for sharing in meetings |
| **Analytics** | Aggregate stats — course clearance rates, selection rates, proposals by project code |

---

## Managing Associates

### Add an Associate
1. In **Manager View**, fill in **Name** and **Joining Date** in the top form
2. Click **Add Associate**
3. The associate card appears in the list

### Edit an Associate
Click the associate's row to expand it. You can:
- Update course completion percentages (0–100)
- Add/edit client interactions
- See days on bench (calculated automatically)

### Archive an Associate
When an associate gets placed (selected by a client):
1. Expand their row in Manager View
2. Scroll to the bottom of the expanded panel
3. Click **Archive** — the associate disappears from the active list
4. They can be viewed/restored from the **Admin** panel

---

## Course Tracking

Courses are managed globally from the **Admin** panel (gear icon in header).

### Add a New Course
Admin panel → **Manage Courses** → type the course name → **Add Course**
The course appears as a field for all active associates.

### Update Progress
In Manager View → expand an associate → slide or type the percentage (0–100) for each course.

**Colour indicators:**
- Green (≥ 80%) — Completed
- Yellow (≥ 40%) — Good progress
- Orange (> 0%) — Started
- Grey (0%) — Not started

No pass/fail threshold is applied — percentages reflect actual progress only.

---

## Client Pipeline

Each associate can have multiple client interactions recorded.

### Add a Client Entry
1. Expand an associate's row
2. Click **+ Add Client**
3. Fill in: Client Name, Project Code, Status, Proposed Date
4. Status options: `Proposed` → `Internal Round` → `Client Round` → `Selected` / `Rejected`

### Update Status
Edit any field inline. Dates for each stage update as status progresses.

---

## Analytics Panel

Click **Analytics** in the header to open the analytics panel.

| Metric | Description |
|---|---|
| Total Bench | Active (non-archived) associates |
| Proposed | Associates with at least one client proposal |
| In Pipeline | Associates in Internal/Client round |
| Placed | Associates with a Selected outcome |
| Selection Rate | Selected ÷ Total proposals (%) |
| Avg Days on Bench | Mean days since joining for active associates |
| Course Distribution | Stacked bar — completed / in progress / not started per course |
| Pipeline Funnel | Bar chart — proposals → rounds → selections |
| Proposals by Project | Chips showing submission count per project code |

---

## PDF Export

Click **Export PDF** (top right) to download a report.

The PDF contains:
- Summary table of all active associates with course completion status
- Two aggregate team-level charts (course distribution + pipeline funnel)
- Works correctly for any team size (charts are team-level, not per-person)

---

## Data Sync

The header shows real-time sync status:

| Indicator | Meaning |
|---|---|
| `· saving…` | Writing to OneDrive (debounced 1.2s after last change) |
| `· synced` | Data successfully saved |
| `· ⚠ sync failed` | Network error — refresh and try again |

All data is saved automatically. There is no manual Save button.
