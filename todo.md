# G12 Job Control — TODO

## Branding & Theme
- [x] Generate G12 app icon and logo (Navy + hard hat / construction motif)
- [x] Apply G12 Navy (#1B2A4A) as primary colour in theme.config.js
- [x] Load Montserrat font for headings
- [x] Update app name to "G12 Job Control"

## Auth & API Layer
- [x] Install and configure expo-auth-session for Microsoft OAuth (MSAL)
- [x] Create lib/auth/microsoft.ts — MSAL auth flow with Azure AD
- [x] Create lib/api/sharepoint.ts — typed helpers for Jobs, DaySheets, Subcontractors lists
- [x] Store access token securely with expo-secure-store
- [x] Login screen with Microsoft 365 sign-in button
- [x] AuthContext provider with session restore

## Navigation
- [x] Configure 5-tab bottom navigation: Dashboard, Jobs, Day Sheets, Compliance, Profile
- [x] Add icon mappings for all tabs
- [x] Stack navigators for detail screens

## Dashboard Screen
- [x] Summary cards: Active Jobs, Pending Approvals, Compliance Alerts
- [x] Quick action buttons: New Day Sheet, View Jobs, Compliance
- [x] Connectivity indicator + last sync timestamp
- [x] Offline queue banner

## Jobs Module
- [x] Jobs list screen (FlatList, pull-to-refresh, search)
- [x] Job detail screen (all fields, linked day sheets action)
- [x] Fetch from SharePoint "Jobs" list via Graph API

## Day Sheets Module
- [x] Day sheets list screen (tabs: Pending / Approved / Exported)
- [x] New day sheet form (jobcode picker, hours, allowances)
- [x] Day sheet detail + supervisor approval/reject actions
- [x] Submit to SharePoint "DaySheets" list
- [x] Calculated fields: ordinary hours, overtime

## Subcontractor Compliance Module
- [x] Compliance list screen (status badges, search, filter)
- [x] Subcontractor detail screen (all compliance fields)
- [x] New subcontractor form
- [x] Mobilisation approval toggle (supervisor only)
- [x] Expiry alert logic (< 30 days = warning, expired = blocked)

## Offline Support
- [x] AsyncStorage offline queue for day sheet submissions
- [x] Auto-sync on reconnect (manual sync from Profile)
- [x] Offline indicator banner on Dashboard

## Profile / Settings
- [x] User info from Microsoft account (name, email, initials avatar)
- [x] Sign out with confirmation
- [x] Sync status + offline queue count + manual sync
- [x] App version display

## Azure AD Configuration
- [x] EXPO_PUBLIC_AZURE_CLIENT_ID set (G12-PnP-SharePoint-Access)
- [x] EXPO_PUBLIC_AZURE_TENANT_ID set (G12 Consulting tenant)
- [x] Mobile redirect URI registered in Entra ID
