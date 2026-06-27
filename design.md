# G12 Job Control — Mobile App Design

## Brand Identity
- **Primary colour:** G12 Navy `#1B2A4A`
- **Accent colour:** G12 Blue `#2563EB`
- **Success:** `#16A34A`
- **Warning:** `#D97706`
- **Error:** `#DC2626`
- **Background (light):** `#F8FAFC`
- **Surface (light):** `#FFFFFF`
- **Background (dark):** `#0F172A`
- **Surface (dark):** `#1E293B`
- **Typography:** Montserrat (headings), Inter / System (body)

---

## Screen List

### Auth
1. **Login Screen** — Microsoft 365 sign-in button, G12 logo, tagline

### Main Tab Bar (5 tabs)
2. **Dashboard** (Home) — Summary cards: active jobs, pending approvals, compliance alerts
3. **Jobs** — Searchable list of active jobcodes from SharePoint
4. **Day Sheets** — Capture and review payroll-ready day sheets
5. **Compliance** — Subcontractor compliance register
6. **Profile / Settings** — User info, sync status, sign out

### Modal / Stack Screens
7. **Job Detail** — Full job record, site info, documents
8. **New Day Sheet** — Form: jobcode, worker, date, hours, allowances, notes
9. **Day Sheet Detail** — View/approve/reject a submitted day sheet
10. **Subcontractor Detail** — Full compliance record, expiry dates, status
11. **New Subcontractor** — Form to add/update subcontractor compliance record
12. **Sync Status** — Offline queue viewer, last sync timestamp

---

## Primary Content and Functionality

### Dashboard
- Summary cards: Active Jobs count, Pending Day Sheet Approvals, Compliance Alerts (expired/expiring)
- Quick action buttons: New Day Sheet, View Jobs
- Last sync timestamp and connectivity indicator

### Jobs Screen
- FlatList of jobs from SharePoint "Jobs" list
- Each card: Jobcode, Job Name, Status badge, Site Location
- Search/filter bar
- Pull-to-refresh
- Tap → Job Detail

### Job Detail Screen
- Header: Jobcode + Job Name
- Fields: Client, Site Address, Start Date, Status, Assigned Supervisor
- Linked Day Sheets count
- Linked Documents (SharePoint document library)
- Action: Add Day Sheet for this job

### Day Sheets Screen
- Tabs: "Pending Approval" | "Approved" | "Exported"
- FlatList of day sheets
- Each card: Worker Name, Jobcode, Date, Hours, Status badge
- Tap → Day Sheet Detail
- FAB: New Day Sheet

### New Day Sheet Form
- Jobcode picker (from Jobs list)
- Worker name (pre-filled from auth user)
- Date picker
- Start time / Finish time
- Break duration
- Calculated: Ordinary hours, Overtime
- Allowances (multi-select)
- Notes text area
- Submit button → saves to SharePoint + local queue

### Day Sheet Detail
- Read-only view of all fields
- Approval section (for supervisors): Approve / Reject buttons
- Payroll export status + Xero reference

### Compliance Screen
- FlatList of subcontractors
- Status badges: Active (green), Expiring Soon (amber), Blocked (red)
- Search by company name or trade
- Filter by status
- Tap → Subcontractor Detail
- FAB: Add Subcontractor

### Subcontractor Detail
- Company info: Name, ABN, Trade, Contact
- Compliance fields: Insurance expiry, Licence expiry
- Induction status, SWMS status
- Prequalification status
- Mobilisation approval (toggle for supervisors)
- Active jobcodes
- Blocked/Not Approved indicator

### Profile / Settings
- User avatar + name + email (from Microsoft account)
- Tenant: G12 Group
- Last sync time
- Offline queue count
- Sign Out button
- App version

---

## Key User Flows

### Flow 1: Field Worker Submits Day Sheet
1. Opens app → Dashboard
2. Taps "New Day Sheet" FAB or quick action
3. Selects jobcode from picker
4. Fills in date, start/finish, breaks, allowances
5. Taps Submit → saved to local queue + synced to SharePoint
6. Returns to Day Sheets screen, sees new entry in "Pending Approval"

### Flow 2: Supervisor Approves Day Sheet
1. Opens app → Day Sheets tab → "Pending Approval"
2. Taps day sheet card
3. Reviews all fields
4. Taps "Approve" → status updates in SharePoint
5. Day sheet moves to "Approved" tab, ready for payroll export

### Flow 3: Check Subcontractor Compliance Before Mobilisation
1. Opens app → Compliance tab
2. Searches for subcontractor by name
3. Taps card → Subcontractor Detail
4. Reviews all expiry dates and statuses
5. If all clear → toggles "Mobilisation Approved"
6. If blocked → sees red "Blocked" badge, cannot approve

### Flow 4: Offline Capture
1. Field worker has no internet
2. Creates day sheet → saved to AsyncStorage offline queue
3. App shows "Offline — 1 item queued"
4. When back online → auto-syncs to SharePoint
5. Shows "Synced" confirmation

---

## Layout Principles
- Portrait orientation, one-handed use
- Tab bar at bottom (iOS HIG)
- Cards with rounded corners (12px), subtle shadow
- Status badges: pill shape, colour-coded
- FAB (Floating Action Button) for primary create actions
- Pull-to-refresh on all list screens
- Empty states with icon + message + action button
- Loading skeletons (not spinners) for list screens
