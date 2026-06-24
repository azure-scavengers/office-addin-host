# S4U Booking — Deployment Runbook
End-to-end setup for the **POD Booking + Client DB** Excel task-pane add-inand its scheduled reconcile pipeline.
> Total wall-clock to a green production state: **~90 minutes**, of which only> ~15 minutes per POD requires Excel Desktop (Power Query authoring is the> only desktop-bound step).
---
## Table of contents
1. [Architecture](#architecture)2. [Prerequisites](#prerequisites)3. [Step 1 — Verify the master `Database` sheet](#step-1--verify-the-master-database-sheet)4. [Step 2 — Create the Power Query mirror in each POD](#step-2--create-the-power-query-mirror-in-each-pod)5. [Step 3 — Build the Power Automate reconcile flow](#step-3--build-the-power-automate-reconcile-flow)6. [Step 4 — End-to-end smoke test](#step-4--end-to-end-smoke-test)7. [Step 5 — Hand off to users](#step-5--hand-off-to-users)8. [Troubleshooting](#troubleshooting)9. [Operational rhythm](#operational-rhythm)10. [Rollback](#rollback)
---
## Architecture
```text                          ┌──────────────────────────────────────────────┐                          │  S4U CLIENT DB.xlsx   (silent, no user opens) │                          │  ─────────────────────────────────────────── │                          │  Database sheet  (the source of truth)        │                          │  Office Script: "Reconcile New Clients"       │                          └──────────────┬───────────────────────────────┘                                         │                          Power Query (refresh on open / Refresh All)                                         ▲                                         │   ┌──────────────────────────┐          │           ┌──────────────────────────┐   │  POD A.xlsx              │──────────┼───────────│  POD B.xlsx              │   │  - Database (hidden,     │          │           │  - Database (hidden,     │   │    Power Query mirror)   │          │           │    Power Query mirror)   │   │  - tbl_NewClientQueue    │          │           │  - tbl_NewClientQueue    │   │  - Ribbon:               │          │           │  - Ribbon:               │   │      POD Booking         │          │           │      POD Booking         │   │      Client DB           │          │           │      Client DB           │   └────────────┬─────────────┘          │           └────────────┬─────────────┘                │                        │                        │                └────────────────────────┼────────────────────────┘                                         │                            Power Automate (scheduled every 15 min)                                "S4U Reconcile New Clients"                                         │                              For each POD's pending queue row:                                 Run script → Master Database                                 Update row → Status = Processed```
**Key invariants**
| What | Where | Why ||---|---|---|| `Database` (plain sheet) | Master `S4U CLIENT DB.xlsx` | Source of truth; admins never touch it directly. || `Database` (hidden, PQ mirror) | Each POD workbook | Drives Fill Row and New Client dropdowns. || `tbl_NewClientQueue` (Excel Table) | Each POD workbook | Auto-created by the add-in; required by Power Automate. || Add-in HTML/JS | GitHub Pages (`azure-scavengers.github.io/office-addin-host`) | Stateless static hosting. || Manifest | M365 admin centre → Integrated apps | Deploys ribbon to POD user group. |
---
## Prerequisites
| Item | Required? | Notes ||---|---|---|| M365 admin role: Office Apps Admin (or higher) | yes | To upload the manifest. || Power Automate access (standard tier) | yes | Scheduled trigger only — **no Premium licence required**. || Excel Desktop (or Windows 365 / iPad Excel / borrowed PC) | once | Only for the one-time Power Query authoring in each POD. || SharePoint site that hosts the workbooks | yes | All four files (master + 3 PODs) must live there. || `S4U CLIENT DB.xlsx` on SharePoint | yes | The silent master. || `POD A.xlsx`, `POD B.xlsx`, `POD RoW.xlsx` on SharePoint | yes | The user-facing PODs. || GitHub Pages site hosting the add-in source | yes | Already live at `azure-scavengers.github.io/office-addin-host`. |
**Files already deployed** (mark as done before starting):
- [x] HTML / JS published to GitHub Pages- [x] `manifest.xml` uploaded to M365 admin centre- [x] `ReconcileNewClients.ts` saved in `S4U CLIENT DB.xlsx`
---
## Step 1 — Verify the master `Database` sheet
> Estimated time: **5 minutes** in the browser.
1. Open `S4U CLIENT DB.xlsx` in Excel for the Web (admin only, one-time).2. Confirm the **Database** sheet:   - Sheet name is **exactly** `Database` (case-sensitive).   - Row 1 contains headers including, at minimum, `User` and `Client`.3. Recognised header aliases (case-insensitive, any order):
   | Logical field | Accepted column headers |   |---|---|   | User | `User`, `Username`, `User Name` |   | Client | `Client`, `Client Name` |   | Client code (optional) | `Client Id`, `Client Code` |   | POD | `POD` |   | Sensitivity | `Sensitivity` |   | Reviewing | `Reviewing`, `INC`, `Flag` |
4. **Close** the workbook. No edits are required unless `User` / `Client` columns are missing.
> The script throws clear errors if these conditions aren't met. They will show up later in the Power Automate run log.
---
## Step 2 — Create the Power Query mirror in each POD
> Estimated time: **~15 minutes per POD × 3 = 45 minutes**.> **Requires Excel Desktop or Excel iPad** (Web cannot author Power Query connections). Once authored, refresh runs in the Web app fine.
Repeat for `POD A.xlsx`, `POD B.xlsx`, `POD RoW.xlsx`. Below uses POD A as the template.
### 2.1 Open the POD from SharePoint in Excel Desktop
1. In SharePoint → click `POD A.xlsx` to open in browser.2. Top-right → **Editing > Open in Desktop App**.
### 2.2 Author the Power Query connection
1. **Data** ribbon → **Get Data > From File > From SharePoint Folder**.2. Paste the **site root URL** (not the file URL), e.g.   `https://<tenant>.sharepoint.com/sites/<sitename>`.3. Sign in with your work account. Click **Connect**.4. Navigator opens listing every file in the site. Find `S4U CLIENT DB.xlsx`.5. Click **Transform Data**. Power Query Editor opens.6. Right-click the row for `S4U CLIENT DB.xlsx` → **Drill Down** into the `Content` column.   - The editor now shows the workbook's contents — sheets and tables.7. Find the row where **Name = `Database`** and **Kind = `Sheet`**.8. Click the **Table** link in the **Data** column. The preview switches to the actual rows.9. If headers show as `Column1, Column2, …`:   - **Home > Use First Row as Headers**.10. Clean up:    - Right-click any extra columns Power Query injected (e.g. `Source.Name`) → **Remove**.    - Select all columns → **Home > Detect Data Type**.11. Right rail → **Query Settings > Name** → rename to `qMasterClients`.
### 2.3 Load into a new sheet
1. **Home > Close & Load > Close & Load To…**.2. **Select how to view this data**: Table.3. **Where**: New worksheet → **OK**.4. A new sheet appears containing the master data.
### 2.4 Rename and hide
1. Right-click the new sheet tab → **Rename** → type exactly `Database` → Enter.2. Right-click again → **Hide**.
> The sheet name must be **exactly `Database`** (case-sensitive). The add-in looks for this literal string.
### 2.5 Enable refresh-on-open
1. **Data > Queries & Connections** (opens right-hand pane).2. Right-click `qMasterClients` → **Properties**.3. Tick both:   - **Refresh data when opening the file**   - **Enable background refresh**4. **OK**.
### 2.6 Save and verify
1. **Ctrl + S** to save back to SharePoint.2. **Close** the POD workbook.3. **Reopen** it from SharePoint.4. **POD Booking > Fill Row** → User dropdown should populate within a second.5. **Client DB > New Client** → Existing Client dropdown should populate.
If both populate, POD A is done. Repeat 2.1–2.6 for POD B and POD RoW.
---
## Step 3 — Build the Power Automate reconcile flow
> Estimated time: **20 minutes**, one-time, admin only.> Uses standard `Excel Online (Business)` connector. **No Premium licence required** — we use a Scheduled trigger.
### 3.1 Create the flow
1. Open `https://make.powerautomate.com`.2. **Create > Scheduled cloud flow**.3. Name: `S4U Reconcile New Clients`.4. Recurrence: every **15 minutes** (adjust as needed).5. Click **Create**.
### 3.2 For each POD, add a drain block
You will repeat this block three times — once per POD file (POD A / POD B / POD RoW). Below shows POD A; the only difference for B and RoW is the file selection.
#### Block A — POD A
1. **+ New step** → **Excel Online (Business) > List rows present in a table**:   - **Location**: SharePoint   - **Document Library**: where `POD A.xlsx` lives   - **File**: `POD A.xlsx`   - **Table**: `tbl_NewClientQueue`   - **Filter Query**: `Status eq 'Pending'`
> If the Table dropdown is empty, it means no one has clicked **Client DB > New Client > Submit** in `POD A.xlsx` yet. Submit a dummy entry there first to auto-create the table, then refresh this dropdown.
2. **+ New step** → **Apply to each** on the dynamic value `value` (output of step 1).
3. **Inside the Apply to each loop**, add **Excel Online (Business) > Run script**:   - **Location** / **Document Library** / **File**: `S4U CLIENT DB.xlsx`   - **Script**: `Reconcile New Clients`   - **payload** parameter (click **Add new item** to expand the object, or paste the JSON shape):
     ```jsonc     {       "clientName":  @{items('Apply_to_each')?['ClientName']},       "userName":    @{items('Apply_to_each')?['UserName']},       "pod":         @{items('Apply_to_each')?['POD']},       "sensitivity": @{items('Apply_to_each')?['Sensitivity']},       "reviewing":   @{items('Apply_to_each')?['Reviewing']}     }     ```
4. Still inside the loop, add **Excel Online (Business) > Update a row**:   - **Location** / **Library** / **File**: `POD A.xlsx`   - **Table**: `tbl_NewClientQueue`   - **Key Column**: `QueuedAt`   - **Key Value**: `@{items('Apply_to_each')?['QueuedAt']}`   - **Status**: `@{outputs('Run_script')?['body/result/status']}`   - **ProcessedAt**: `@{utcNow()}`   - **Error**: `@{outputs('Run_script')?['body/result/message']}`
#### Block B — POD B
1. Copy Block A's three actions (right-click each → **Copy to my clipboard**).2. Paste after Block A.3. Change the file from `POD A.xlsx` to `POD B.xlsx` in both **List rows** and **Update a row**.
#### Block C — POD RoW
Same as Block B, with `POD RoW.xlsx`.
### 3.3 Save & manually test
1. Top-right → **Save**.2. Top-right → **Test > Manually > Run**.3. Watch the run summary. All actions should be green.4. Pending rows in any POD should now show **Status = Processed** with a populated `ProcessedAt`.
---
## Step 4 — End-to-end smoke test
> Estimated time: **10 minutes**.
1. **As a POD user**, open `POD A.xlsx` in Excel for the Web.2. **POD Booking > Fill Row** → confirm dropdowns populate with existing master users/clients.3. **Client DB > New Client** → fill the form:   - Brand new client: `Acme Smoke Test / Jane Tester / POD A`.   - Click **Add New Client** → toast: *"Queued: client Acme Smoke Test added."*4. Right-click any sheet tab → **Unhide** → `_NewClientQueue` → **OK**.   - Confirm one row with `Status = Pending`. Note the `QueuedAt` value.   - Hide it again.5. **Power Automate** → **My flows** → **S4U Reconcile New Clients** → top-right **Run**.6. Wait ~30 s, refresh the **Run history** → run should be green.7. Re-open `_NewClientQueue` (unhide) → the row's `Status` is now `Processed`, `ProcessedAt` populated, `Error` empty. Hide it back.8. (Admin) Open `S4U CLIENT DB.xlsx` → `Database` sheet → new row at the bottom with `Jane Tester / Acme Smoke Test / POD A / High / INC`.9. Back in `POD A.xlsx` → **Data > Refresh All** (or close and reopen).10. **POD Booking > Fill Row** → User dropdown → pick `Jane Tester` → Client dropdown should now include `Acme Smoke Test`.
If steps 1–10 pass, the system is live end-to-end. Repeat the smoke test from POD B and POD RoW for full coverage.
---
## Step 5 — Hand off to users
Suggested 3-line note to POD users:
> The POD workbooks now have two new ribbon tabs:> - **POD Booking**: Fill Row, In Time, CSS / VA DL (same as before).> - **Client DB > New Client**: queue a new client / user mapping. The new entry appears in Fill Row dropdowns within ~15 minutes after you submit and click **Data > Refresh All**.>> Please report any issues to `<admin>`.
---
## Troubleshooting
<details><summary><b>New Client: "No clients found"</b></summary>
POD doesn't yet have its local `Database` mirror sheet (Power Query). Do [Step 2](#step-2--create-the-power-query-mirror-in-each-pod) for that POD.
</details>
<details><summary><b>Fill Row: dropdowns empty</b></summary>
Same root cause as above — local `Database` sheet missing or unrefreshed. Either complete Step 2, or in an existing POD click **Data > Refresh All**.
</details>
<details><summary><b>Power Automate: "List rows" Table dropdown is empty</b></summary>
No one has clicked **Client DB > New Client > Submit** in that POD yet, so `tbl_NewClientQueue` doesn't exist. Submit a dummy New Client entry in the POD, then click the **Refresh** icon next to the Table dropdown.
</details>
<details><summary><b>Flow's <code>Run script</code> returns <code>status: "Error", message: "Master workbook is missing worksheet 'Database'"</code></b></summary>
The Database sheet was renamed in the master. Rename it back to exactly `Database` (case-sensitive) and re-run the flow.
</details>
<details><summary><b>Flow's <code>Run script</code> returns <code>status: "Error", message: "Database sheet must have at least 'User' and 'Client' columns"</code></b></summary>
Row 1 of the master `Database` sheet is missing `User` and/or `Client` headers (or uses non-recognised names). Add or rename columns per the alias table in [Step 1](#step-1--verify-the-master-database-sheet).
</details>
<details><summary><b>Flow's <code>Run script</code> returns <code>status: "Skipped"</code></b></summary>
Not an error — the script found an exact `Client + User` duplicate already in the master Database. Safe to ignore. Verify in the master if needed.
</details>
<details><summary><b>Reconcile succeeded but POD users don't see the new client</b></summary>
The Power Query mirror in the POD hasn't refreshed yet. The user must click **Data > Refresh All**, or close and reopen the POD file. The mirror also refreshes automatically on next open.
</details>
<details><summary><b>New Client form opens but the form body looks empty / no fields</b></summary>
The hosted `newClient.html` on GitHub Pages is stale or wrong. Re-upload the file from `src/newClient.html`, then in the user's machine clear the Office Wef cache:
```powershellStop-Process -Name EXCEL -Force -ErrorAction SilentlyContinueRemove-Item "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" -Recurse -Force -ErrorAction SilentlyContinue```
</details>
<details><summary><b>Ribbon tabs don't appear on Excel Desktop</b></summary>
Centrally deployed add-ins can take **up to 24 hours** to propagate to Desktop on first deployment. Restart Excel. If still missing, check **Insert > Get Add-ins > Admin Managed** — the add-in should be listed there.
</details>
---
## Operational rhythm
| Event | Who | What ||---|---|---|| POD user opens their workbook | User | Power Query auto-refreshes the hidden `Database` mirror. || POD user adds a booking | User | Fill Row appends a row to the POD operational sheet. No master impact. || New client request | POD user | Opens `Client DB > New Client`, fills the form. Row goes to local `_NewClientQueue`. || Every 15 minutes | Power Automate | Drains every POD's queue into the master `Database` sheet. || POD user clicks Data > Refresh All | User | Pulls the latest master into their hidden mirror; new client appears in dropdowns. || Schema change (extra column in master) | Admin | Add a column to master `Database`. The add-in's header-alias detection ignores unknown columns; existing forms keep working. |
---
## Rollback
If something needs to be reverted:
1. **Add-in**: `https://admin.microsoft.com → Settings → Integrated apps`, find `POD Booking`, click **Remove app**. Ribbon disappears within minutes (Web) or on next Excel restart (Desktop). No data is touched.2. **Flow**: `https://make.powerautomate.com → My flows`, toggle the flow **Off**, or **Delete**. Existing queue rows simply stay at `Pending`.3. **Master script**: Open `S4U CLIENT DB.xlsx` → **Automate > All Scripts** → delete `Reconcile New Clients`.4. **POD Power Query**: Open each POD → **Data > Queries & Connections** → right-click `qMasterClients` → **Delete**. The hidden `Database` sheet can also be unhidden and deleted.
No POD operational data (bookings, In Time history, CSS / VA DL values) is impacted by any rollback step.
---
## File layout (repository)
```Excel_Migration/├── TaskPaneAddin/│   ├── manifest.xml                  ← uploaded to M365 admin centre│   ├── ReconcileNewClients.ts        ← pasted into S4U CLIENT DB.xlsx → Automate│   ├── README.md                     ← summary / architecture│   ├── SETUP-RUNBOOK.md              ← this file│   ├── DEPLOY.md                     ← legacy two-manifest deploy notes (kept for history)│   ├── assets/                       ← icons hosted on GitHub Pages│   │   ├── icon-16.png│   │   ├── icon-32.png│   │   └── icon-80.png│   └── src/                          ← hosted on GitHub Pages│       ├── common.js│       ├── form.html                 ← Fill Row pane│       ├── timepicker.html           ← In Time pane│       ├── daytimepicker.html        ← CSS / VA DL pane│       └── newClient.html            ← New Client pane (writes to queue)└── OfficeScripts/    └── ReconcileNewClients.ts        ← duplicate of the master script for reference```
---
## What's next (optional enhancements)
- **Email notification on Skipped/Error rows** — add a step to the flow that sends a Teams/Outlook message when `Run script > body/result/status` is not `Processed`.- **Auto-refresh PODs after reconcile** — the flow can also call a small Office Script in each POD to trigger a Power Query refresh, so users do not need to click `Data > Refresh All`.- **Audit trail** — add a `_AuditLog` sheet in the master that records every Processed row with `ProcessedBy` (the flow's connection account), already available as flow dynamic content.
These are non-blocking — the current pipeline works end-to-end without them.
---
_Maintained by the S4U Migration team. Last updated for code base version `manifest.xml v1.0.2.0`._
