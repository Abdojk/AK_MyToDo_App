# AK MyToDo Hub

Every flagged Outlook email, Planner task, and calendar occurrence in
`Akhoury@info-sys.com`, laid out by date:
**Overdue · Today · Tomorrow · Next 7 days · Later · No due date**.

All three share each column, in one chronological stream, so a Thursday with four
meetings and two Planner tasks reads as exactly that.

The Hub is a browser-only single-page app. It signs in against the info-sys.com
Entra ID tenant, reads Microsoft Graph, and renders. There is no server, no
database, and no client secret.

Two actions write back. **Done** marks a task complete; **Move to** shifts its due
date to today, tomorrow, next week, or any date you type. Both patch the task in
Microsoft To Do. Nothing writes to the message resource.

## How it reads the mailbox

Flagging an email in Outlook creates a task in the built-in **Flagged email**
list in Microsoft To Do. Microsoft Graph exposes that list as the `todoTaskList`
whose `wellknownListName` is `flaggedEmails`. The Hub reads it:

```
GET   /me/todo/lists                          → find wellknownListName = flaggedEmails
GET   /me/todo/lists/{id}/tasks               → the flagged emails, with due dates
POST  /$batch  (GET linkedResources per task) → deep link back to each email
GET   /me/messages?$filter=flag/flagStatus…   → optional: sender name
PATCH /me/todo/lists/{id}/tasks/{taskId}      → Done and Move to
GET   /me/calendarView?startDateTime=…        → meetings, 30 days ahead
GET   /me/planner/tasks                       → Planner tasks assigned to you
GET   /planner/plans/{id}                     → plan names, batched
PATCH /planner/tasks/{taskId}                 → Done and Move to, with If-Match
```

The `/me/messages`, `/me/calendarView` and `/me/planner/tasks` calls are all
best-effort. If the tenant or the API rejects any of them, the Hub says so in a
banner and renders whatever it did get. None can cost the timeline its flagged
email.

`PATCH` is the only write verb the client issues, and it reaches only the two task
endpoints above. See `src/graph/write.ts` and `src/graph/planner.ts`.

## Writing back

Both actions apply optimistically: the card moves column, or leaves the timeline,
on the click. The Hub then issues the PATCH and replaces the local task with what
the server returns. A rejection restores the task exactly as it was and names the
failure in the error banner, so the timeline never shows a change the mailbox does
not hold.

Completing a task raises an undo bar. **Undo** patches the status back to
`notStarted`. The bar clears on the next refresh.

A due date is stored as midnight of the chosen calendar day **in your zone**,
converted to a UTC instant and sent as `{"dateTime": "…", "timeZone": "UTC"}`. It
therefore reads back as the same calendar day you picked, whatever the offset
between your zone and UTC.

## Step 0 — Register the app in Entra ID

Run once, in the info-sys.com tenant.

1. Open the Microsoft Entra admin centre → **App registrations** → **New
   registration**.
2. **Name**: `AK MyToDo Hub`.
3. **Supported account types**: *Accounts in this organizational directory only
   (Info-Sys only — Single tenant)*.
4. **Redirect URI**: platform **Single-page application (SPA)**, value
   `http://localhost:5173`.
5. Select **Register**.
6. Copy **Application (client) ID** and **Directory (tenant) ID** from the
   Overview page.
7. Go to **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions**, and add:

   | Permission | Why |
   |---|---|
   | `User.Read` | Sign-in and the account name in the header |
   | `Tasks.ReadWrite` | The Flagged email list, its tasks, and the two write actions |
   | `Mail.ReadBasic` | Optional: the sender name on each task card |
   | `Calendars.ReadBasic` | Optional: meetings on the timeline |

   `Tasks.ReadWrite` also covers Planner: reading your assigned tasks, reading
   plan names, and both write actions. Planner needs no permission of its own.

8. Select **Grant admin consent for Info-Sys** if the tenant blocks user consent.
   Skip it if the tenant allows user consent for these delegated scopes.
9. On **Authentication**, confirm the SPA platform holds the redirect URI and
   that the implicit grant checkboxes stay unticked. MSAL uses the authorization
   code flow with PKCE, not implicit flow.

Graph offers no permission narrower than `Tasks.ReadWrite` for updating a task, so
the write actions require it. To run the Hub read-only, set `graphScopes` back to
`Tasks.Read` in `src/auth/msalConfig.ts` and drop the `card-actions` block from
`src/components/TaskCard.tsx`.

Drop `Mail.ReadBasic` to run without sender names, or `Calendars.ReadBasic` to run
without meetings. Remove each from both the registration and `graphScopes`; the
Hub shows a banner in place of the missing data rather than failing.

Changing the scope list invalidates any token already cached in the browser. The
next sign-in raises a fresh consent prompt, which `useAuth.getToken` handles by
falling back to a popup.

## Run it

```bash
cp .env.example .env     # then paste in the client and tenant IDs from step 0
npm install
npm run dev              # http://localhost:5173
```

| Script | Does |
|---|---|
| `npm run dev` | Development server on port 5173 |
| `npm test` | Unit tests for the bucketing and normalisation logic |
| `npm run typecheck` | `tsc --noEmit` |
| `npm run build` | Type check, then a production bundle in `dist/` |

Deploying to a hosted URL needs that URL added as a second SPA redirect URI in
the registration, and `VITE_REDIRECT_URI` set to match.

## Layout

```
src/
├─ main.tsx              MSAL bootstrap and React root
├─ App.tsx               sign-in gate, Graph read, refresh, error surface
├─ auth/
│  ├─ msalConfig.ts      client ID, authority, scopes, cache policy
│  └─ useAuth.ts         signIn, signOut, getToken
├─ graph/
│  ├─ client.ts          bearer auth, paging, $batch, PATCH, 429 back-off
│  ├─ todo.ts            flagged list → tasks → linked resources
│  ├─ write.ts           completeTask, reopenTask, rescheduleTask
│  ├─ calendar.ts        calendarView over a 30-day window, optional
│  ├─ planner.ts         Planner read and writes, ETag handling, optional
│  └─ mail.ts            optional sender enrichment, joined on subject
├─ model/
│  ├─ types.ts           HubTask, HubEvent, HubItem, Bucket, raw Graph shapes
│  ├─ dates.ts           dateTimeTimeZone → instant, calendar-day maths
│  ├─ buckets.ts         bucketOf, groupItems
│  ├─ events.ts          event filtering, normalising, time formatting
│  ├─ planner.ts         priority bands, normalising, the task link
│  ├─ reschedule.ts      preset and typed dates → a Graph dueDateTime
│  ├─ mutate.ts          optimistic apply by item key, rollback, response merge
│  └─ normalise.ts       todoTask + linkedResource → HubTask
└─ components/           Timeline, BucketColumn, TaskCard, PlannerCard,
                         EventCard, TaskActions, SignInPanel
```

## Planner

`GET /me/planner/tasks` returns the tasks assigned to you across every plan. Plan
names come from `GET /planner/plans/{id}`, batched into one request.

Left out: tasks with **no due date**, and tasks at `percentComplete` 100. A
Planner backlog is largely undated, and the No due date column exists for flagged
email, so undated Planner work stays out rather than burying it.

Planner cards carry the same **Done** and **Move to** row as flagged email. Two
differences sit under them:

- `plannerTask.dueDateTime` is a bare `DateTimeOffset` in UTC, not the
  `dateTimeTimeZone` pair `todoTask` uses, so the write format differs. See
  `toPlannerDueDate` in `src/model/planner.ts`.
- Every Planner `PATCH` **requires** an `If-Match` ETag. When the task has changed
  in Planner since the Hub read it, Graph answers `412`; `src/graph/planner.ts`
  re-reads the task once for a fresh ETag and retries, and only a second failure
  reaches the rollback. The write also sends `Prefer: return=representation`,
  because Planner otherwise answers `204` with no body.

Priority follows Planner's own bands: 0-1 urgent, 2-4 important, 5-7 medium, 8-10
low. Only urgent and important show a pill, so the column stays quiet.

## The calendar

`GET /me/calendarView` returns occurrences, not series masters, so a weekly
meeting arrives as one card per week and the Hub never walks a recurrence
pattern. The window runs from midnight today in your zone to 30 days ahead
(`HORIZON_DAYS` in `src/graph/calendar.ts`).

Left out: cancelled occurrences, and invitations you declined. Kept: all-day
events, events already finished today, and events marked Free, which carry a
pill so a soft hold does not read as a commitment.

Event cards carry no **Done** or **Move to** row. The Hub holds no calendar write
permission and never creates, edits, accepts, or declines an event.

Only the default calendar is read. Other calendars and shared calendars are not.

### All-day events

Graph documents that an all-day event has its "start, and endtime … set to
midnight and be in the same time zone" — a floating midnight rather than an
instant. Converting it into your zone would move it a day backwards for any zone
behind UTC, so `bucketOf` takes a `floating` flag and compares the raw date parts
instead. `src/model/events.test.ts` covers exactly this: without the flag, an
event on the 13th shows as Today in Los Angeles rather than Tomorrow.

## Time zones

Graph returns a due date as a `dateTimeTimeZone`: a wall-clock string plus a zone
name. `src/model/dates.ts` converts that to an absolute instant, then compares
**calendar days** in the viewer's own zone, so a task due at 23:00 stays in Today
rather than sliding into Tomorrow. The zone comes from the browser
(`Intl.DateTimeFormat().resolvedOptions().timeZone`); reading the mailbox's own
zone instead would cost an extra `MailboxSettings.Read` permission.

Graph sometimes returns Windows zone names such as `Pacific Standard Time`, which
`Intl` rejects. Those fall back to UTC, which matches what Outlook stores on a
flag due date by default.

## Known unknowns

Three points could not be confirmed from the Microsoft Learn documentation and
are handled defensively rather than assumed. Confirm each against the live
mailbox in Graph Explorer, and simplify the code once you know the answer.

1. **`$expand=linkedResources` on the tasks endpoint.** Attempted first; a
   rejection falls back to a batched per-task read (`src/graph/todo.ts`).
2. **`$filter=flag/flagStatus eq 'flagged'` on `/me/messages`.** Attempted; a
   rejection disables sender enrichment and shows a banner (`src/graph/mail.ts`).
3. **Whether `linkedResource.externalId` holds a Graph message id.** Not relied
   on. Sender enrichment joins on the email subject instead, and skips any
   subject that matches more than one flagged message.
4. **Whether completing the To Do task marks the Outlook flag complete.** The task
   in the `flaggedEmails` list is the flag's counterpart, but no Microsoft Learn
   page consulted states what happens to `message.flag.flagStatus` when the task
   status changes. The card leaves the timeline either way, because the timeline
   filters on task status. Check the flag in Outlook straight after the first
   completion and record what it does.
5. **Whether `status: "completed"` is accepted on its own.** `completeTask` sends
   status alone, and on a `400` retries once with `completedDateTime` alongside
   (`src/graph/write.ts`).
6. **Whether `Calendars.ReadBasic` returns `webLink`.** The permission excludes
   "properties such as body, attachments, and extensions" — a non-exhaustive
   list. Not relied on: the Hub builds
   `https://outlook.office365.com/calendar/item/{id}` from the event id, which is
   the current form Microsoft documents, and uses `webLink` only when there is no
   id.
7. **The `responseStatus.response` enum beyond `declined`.** Only that literal
   filters, and anything else counts as not declined, so an unrecognised member
   can never hide an event.
8. **The Planner deep link.** Microsoft Learn documents no URL for opening a
   `plannerTask`. The Hub builds
   `https://tasks.office.com/{tenantId}/Home/Task/{taskId}` from
   `VITE_TENANT_ID` — the widely used pattern, not a documented one. Confirm it
   on the first click. Correcting it is a change to `plannerTaskUrl` in
   `src/model/planner.ts`, and an unset tenant yields no link rather than a
   broken one.
9. **Whether `/me/planner/tasks` pages.** No `$top` or `@odata.nextLink` is
   documented for it. The read goes through `client.getAll`, which follows
   `@odata.nextLink` when present and returns the single page when not.

## Not yet built

Accepting, declining, or creating events; other and shared calendars; free/busy
lookup for other people; a day or week grid view; creating Planner tasks,
assigning them, or moving buckets; editing any task's title, body, or importance;
checklist items; bulk actions across a column; drag and drop between columns;
delta-based background sync; notifications; and hosting beyond `localhost`. Each
is additive. None changes the shape above.
