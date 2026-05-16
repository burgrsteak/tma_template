# TMA Template — Task Management App

A **Google Apps Script** web app that runs entirely inside a Google Spreadsheet. It provides a full-featured task management system for operations teams, with role-based access control, real-time polling, analytics, shift reporting, and an admin settings panel — all served as a single-page application via `google.script.run`.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Runtime | Google Apps Script (V8) |
| Frontend | Vanilla HTML / CSS / JS (SPA served via `HtmlService`) |
| Database | Google Sheets (one sheet = one table) |
| Auth | Google Session + 12-hour cookie stored in `UserProperties` |
| Locking | `LockService` wraps all write operations |
| Email | `MailApp` for high-priority task alerts and shift handover mentions |

---

## Pages

| Page | Description |
|---|---|
| **Login** | Google-authenticated entry point. Checks for an active 12-hour session and routes to the dashboard automatically. |
| **Dashboard** | Primary workspace — task list, creation form, filters, bulk actions, time tracking, comments, and real-time polling. |
| **Analytics** | Metrics overview — task volume by status/type, team workload, SLA overdue breakdown, and average handle time. |
| **Shift Report** | Live roster of who has logged in within the last 24 hours, their active/inactive state, team, and shift schedule. |
| **Settings** | Admin panel for app configuration, user management, role/tier setup, task types, shift schedules, and role-switcher (Dev only). |

---

## Features

### Authentication & Sessions
- Google OAuth identity via `Session.getActiveUser()` — no passwords required.
- 12-hour session timer stored in `UserProperties`. Expired sessions are automatically invalidated on the next page load.
- Ghost sessions (duplicate open sessions for the same user) are auto-closed at login.
- Users not in the `Users` sheet are automatically registered as `Dev` (developer safety net) to prevent lockout after a DB re-init.
- Unauthorized users can submit an access request with a reason; admins receive an email notification immediately.

### Startup Preloading
- On login, `enterApplication()` fires a single parallel preload (`Promise.all`) that fetches:
  - App config and settings
  - Full user list (assignable users)
  - Tasks for the current view
  - Notifications
- The loading screen is shown until all preload calls resolve; the app is only revealed once data is ready.
- Timers and polling intervals start only after preload completes.
- Guards prevent duplicate `enterApplication()` calls and duplicate intervals.

### Role-Based Access Control (RBAC)
- Four-tier permission system:

| Tier | Default Roles | Permissions |
|---|---|---|
| 0 | Admin, Dev | Full access — settings, user management, all tasks |
| 1 | Manager | Admin view of tasks, bulk actions, user edits |
| 2 | Lead, QA | Team-level visibility |
| 3 | User | Own tasks only (created or assigned) |

- Tier configuration is fully dynamic and editable from the Settings page.
- All write APIs enforce tier checks server-side via `getUserTier_()`.

### Task Management
- **Create** tasks with title, description, type, subtype, status, priority (`High` / `Medium` / `Low`), assignee, deadline, and optional metadata.
- **Edit** any field on a task inline; changes are tracked in the activity comment thread.
- **Status updates** are logged as system comments with old → new value.
- **Soft delete** — deleted tasks are marked `isDeleted = true` and copied to a `DeletedTasks` archive sheet.
- **Bulk actions** — select multiple tasks and bulk-update status or assignee in one operation.
- **Task types** are fully configurable: each type has its own color, subtypes, allowed statuses, and custom metadata fields.

### Filtering & Views
- Admins can toggle between **My Tasks**, **My Team**, and **All Tasks** views.
- Team filter resolves the current user's team and shows only tasks assigned to or created by team members.
- Regular users always see only tasks they created or were assigned to.

### Comments & Activity Feed
- Threaded comments on every task.
- **System comments** are auto-posted for every meaningful change (creation, status change, assignment, bulk update, archive).
- Users can **edit** or **soft-delete** their own comments; Tier 0/1 admins can edit/delete any comment.
- A **Recent Activity** feed on the dashboard surfaces the last 50 comment events across all tasks.

### Time Tracking
- Per-task clock-in / clock-out via `API_toggleTime`.
- Duration in minutes is stored in `TimeLogs` when clocking out.
- Open logs (clocked in but not out) are visible to the user.

### Notifications
- In-app notification bell with unread count badge.
- Notifications are pushed server-side (`pushNotification_`) on task assignment and status changes.
- Users can mark individual notifications or all notifications as read.
- Polling interval checks for new notifications and real-time task updates on a configurable heartbeat.

### High-Priority Email Alerts
- Tasks created with `High` priority and an assignee automatically trigger an email to the assignee via `MailApp`.

### Analytics
- Task volume breakdown by **status** and **task type**.
- **Team workload** matrix (team × shift = task count).
- **Overdue detection** — a task is considered overdue if:
  - It has passed its explicit deadline, **or**
  - It has exceeded its SLA timer (configurable per priority level in Settings), **or**
  - Its title is date-prefixed (`YYYY_MM_DD`) and that date has passed.
- Overdue drill-down shows per-team details (task ID, title, priority, status, assignee).
- **Average Handle Time** — mean time from creation to completion across all completed tasks.

### Shift Report
- Live roster showing all session logins within the last 24 hours.
- Displays name, role, time in, time out, active status, team, and scheduled shift window.

### Shift Handover
- End-of-shift summary form with completed count, pending count, narrative summary, handoff notes, and blockers/escalations.
- Users can @-mention teammates; mentioned users receive an email with the handover content.
- All handover submissions are stored in `ShiftLogs`.

### Settings & Configuration (Tier 0/1 only)
- **App name and theme** — accent color and light/dark mode preference.
- **Task types** — add, edit, or remove types; configure subtypes, custom statuses, and metadata fields per type.
- **Roles and tiers** — assign roles to tier levels; controls what each user can see and do.
- **SLA timers** — set overdue thresholds in hours for High, Medium, and Low priority tasks.
- **Shift schedules** — define named shifts with start/end times.
- **Frequency options** — configure recurrence labels used on tasks.
- **User management** — add, edit, bulk-update, or delete user records; assign team, role, and shift.
- **Dev role switcher** — Tier 0 users can switch their own role to test different permission levels without modifying the sheet.

### System Maintenance
- `system_janitorCleanup()` is a scheduled Apps Script trigger that soft-archives tasks with a `Completed` or `Done` status older than 7 days.
- All write operations log to the `ActivityLog` sheet (timestamp, user, action, old data, new data).

---

## Database Tables (Sheets)

| Sheet | Purpose |
|---|---|
| `Tasks` | All task records |
| `Comments` | Task comments and system activity entries |
| `Users` | User roster with role, team, and shift |
| `SessionLogs` | Login/logout audit trail |
| `TimeLogs` | Per-task time tracking entries |
| `Notifications` | In-app notification queue |
| `ShiftLogs` | End-of-shift handover submissions |
| `ActivityLog` | Full audit log of all write operations |
| `AccessRequests` | Pending access requests from unauthorized users |
| `DeletedTasks` | Archive copy of soft-deleted tasks |

---

## Project Structure

```
tma_template/
├── Code.js               # doGet() entry point, HtmlService setup
├── TaskLogic.js          # All API functions (auth, tasks, comments, analytics, settings)
├── DatabaseConfig.js     # Sheet/table helpers, column map, withLock(), colIdx()
├── Utils.js              # generateUUID(), parsePayload(), shared utilities
├── Index.html            # App shell — loading screen, routing, enterApplication(), preload orchestration
├── ClientCore.html       # Shared client-side utilities included in every page
├── appsscript.json       # Manifest — scopes and webapp config
├── Pages/
│   ├── Login.html        # Login screen
│   ├── Dashboard.html    # Main task board
│   ├── Analytics.html    # Metrics and charts
│   ├── ShiftReport.html  # Active session roster
│   └── Settings.html     # Admin configuration panel
└── Components/           # Reusable UI component partials
```

---

## Deployment

1. Clone or copy this project into [Google Apps Script](https://script.google.com).
2. Create a linked Google Spreadsheet and ensure the sheet names match the table names in `DatabaseConfig.js`.
3. Run `system_initDatabase()` (if present) or manually create the required sheets with the correct headers.
4. Deploy as a **Web App** (`Execute as: User accessing the web app`, `Who has access: Anyone within [domain]`).
5. Share the web app URL with your team.

> The first user to log in who is missing from the `Users` sheet will be auto-registered as `Dev` so the owner is never permanently locked out.
