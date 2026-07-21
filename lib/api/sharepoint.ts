/**
 * SharePoint Graph API service for Ausslope Job Control.
 *
 * ALL SITES AND LISTS VERIFIED 2026-07-21 AGAINST LIVE AUSSLOPE TENANT
 *
 * PRIMARY SITE — Job Register, Workers, Projects (asr-operations2):
 *   https://ausslope.sharepoint.com/sites/asr-operations2
 *   Site ID: ausslope.sharepoint.com,aedb5034-4468-4dab-9b6e-94cf53db15be,b3103679-4573-4af6-9a4b-5d7d8bfebe12
 *
 * SECONDARY SITE — Day Sheets, Subcontractor Register, Cost Codes (ProjectControls):
 *   https://ausslope.sharepoint.com/sites/ProjectControls
 *   Site ID: ausslope.sharepoint.com,06c7be2e-83f7-4c3a-93fe-54f6542eaa01,b3103679-4573-4af6-9a4b-5d7d8bfebe12
 *
 * VERIFIED FIELD INTERNAL NAMES:
 *
 * Job Register (ProjectControls) — Title=job code, JobName, Client, SiteAddress,
 *   Jurisdiction, Status, ProjectType, StartDate, EndDate, Value, Supervisor (User), Notes
 *
 * Day Sheets (ProjectControls) — Title=Reference, JobCode (Lookup→Job Register),
 *   Date, SubcontractorCrew (Lookup→Subcontractor Register), LabourHours,
 *   PlantHours (Note), WorkDescription (Note), CostCode, ApprovedBy (User),
 *   Status, OrdinaryHours, OvertimeHours, PayrollExportStatus, XeroReference
 *
 * Subcontractor Register (ProjectControls) — Title=Company, TradeScope, ABN,
 *   PrimaryContact, PublicLiabilityExpiry, ProfessionalIndemnityExpiry,
 *   WorkersCompExpiry, PrequalStatus, Active (Boolean), InductionStatus,
 *   SWMSStatus, LinkedJobs (LookupMulti), PartyType
 *
 * Workers (asr-operations2) — Title=worker name, Person (User), EmployeeID,
 *   Position, WorkerType, StartDate, Status, Employer (Lookup)
 *
 * M365 Group for role gating:
 *   "Ausslope Job Control" → AusslopeJobControl@ausslope.com.au
 */

import { getStoredToken, refreshAccessToken } from "@/lib/auth/microsoft";
import * as FileSystem from "expo-file-system/legacy";
import { Platform } from "react-native";

// ─── Demo Data (fallback when SharePoint is not yet configured) ───────────────

const DEMO_JOBS: Job[] = [
  { id: "demo-1", jobNumber: "ASR-001", jobCode: "ASR-001", jobName: "Kellyville Ridge Retaining Wall", client: "Stockland Development", siteAddress: "45 Kellyville Ridge Rd, Kellyville NSW 2155", status: "Active", startDate: "2026-03-10", completionDate: "2026-08-30", contractValue: "$485,000", projectManager: "Tim Kinsella", superintendent: "Darragh Rabbitte", jobType: "Civil", description: "Construction of 120m reinforced concrete retaining wall with drainage system.", priority: "High" },
  { id: "demo-2", jobNumber: "ASR-002", jobCode: "ASR-002", jobName: "Baulkham Hills Earthworks", client: "Mirvac Group", siteAddress: "12 Windsor Rd, Baulkham Hills NSW 2153", status: "Active", startDate: "2026-04-01", completionDate: "2026-09-15", contractValue: "$320,000", projectManager: "Tim Kinsella", superintendent: "Sameer Joshi", jobType: "Earthworks", description: "Site preparation and bulk earthworks for residential development.", priority: "Medium" },
  { id: "demo-3", jobNumber: "ASR-003", jobCode: "ASR-003", jobName: "Bella Vista Drainage Upgrade", client: "Hills Shire Council", siteAddress: "Bella Vista Drive, Bella Vista NSW 2153", status: "Active", startDate: "2026-05-15", completionDate: "2026-10-01", contractValue: "$210,000", projectManager: "Tim Kinsella", superintendent: "Darragh Rabbitte", jobType: "Civil", description: "Stormwater drainage upgrade including new pits, pipes and kerb and gutter.", priority: "Medium" },
  { id: "demo-4", jobNumber: "ASR-004", jobCode: "ASR-004", jobName: "Pennant Hills Road Widening", client: "Transport for NSW", siteAddress: "Pennant Hills Rd, Thornleigh NSW 2120", status: "On Hold", startDate: "2026-06-01", completionDate: "2026-12-20", contractValue: "$750,000", projectManager: "Tim Kinsella", superintendent: "Sameer Joshi", jobType: "Civil", description: "Road widening and pavement rehabilitation works.", priority: "High" },
  { id: "demo-5", jobNumber: "ASR-005", jobCode: "ASR-005", jobName: "Castle Hill Substation Civil Works", client: "Endeavour Energy", siteAddress: "Castle Hill Rd, Castle Hill NSW 2154", status: "Completed", startDate: "2025-10-01", completionDate: "2026-03-31", contractValue: "$180,000", projectManager: "Tim Kinsella", superintendent: "Darragh Rabbitte", jobType: "Civil", description: "Civil works for new 11kV substation including slab, pits and conduits.", priority: "Low" },
];

const DEMO_DAYSHEETS: DaySheet[] = [
  { id: "demo-ds-1", jobCode: "ASR-001", jobName: "Kellyville Ridge Retaining Wall", workerName: "Jake Morrison", workerEmail: "jake.morrison@ausslope.com.au", date: "2026-06-27", startTime: "07:00", finishTime: "15:30", breakMinutes: 30, ordinaryHours: 8, overtimeHours: 0, allowances: "Travel, Tool", notes: "Completed formwork for section A.", approvalStatus: "Pending", approvedBy: "", approvedDate: "", payrollExportStatus: "Not Exported", xeroReference: "", site: "Kellyville Ridge", trade: "Concretor" },
  { id: "demo-ds-2", jobCode: "ASR-001", jobName: "Kellyville Ridge Retaining Wall", workerName: "Ryan Chen", workerEmail: "ryan.chen@ausslope.com.au", date: "2026-06-27", startTime: "07:00", finishTime: "17:00", breakMinutes: 30, ordinaryHours: 8, overtimeHours: 1.5, allowances: "Travel", notes: "Steel fixing and pour preparation.", approvalStatus: "Pending", approvedBy: "", approvedDate: "", payrollExportStatus: "Not Exported", xeroReference: "", site: "Kellyville Ridge", trade: "Steel Fixer" },
  { id: "demo-ds-3", jobCode: "ASR-002", jobName: "Baulkham Hills Earthworks", workerName: "Liam Walsh", workerEmail: "liam.walsh@ausslope.com.au", date: "2026-06-26", startTime: "06:30", finishTime: "15:00", breakMinutes: 30, ordinaryHours: 8, overtimeHours: 0, allowances: "Travel, Meal", notes: "Cut and fill operations, bulk earthworks.", approvalStatus: "Approved", approvedBy: "Tim Kinsella", approvedDate: "2026-06-27", payrollExportStatus: "Not Exported", xeroReference: "", site: "Baulkham Hills", trade: "Plant Operator" },
  { id: "demo-ds-4", jobCode: "ASR-003", jobName: "Bella Vista Drainage Upgrade", workerName: "Sam Nguyen", workerEmail: "sam.nguyen@ausslope.com.au", date: "2026-06-26", startTime: "07:00", finishTime: "15:30", breakMinutes: 30, ordinaryHours: 8, overtimeHours: 0, allowances: "Travel", notes: "Installed 3x stormwater pits.", approvalStatus: "Approved", approvedBy: "Tim Kinsella", approvedDate: "2026-06-26", payrollExportStatus: "Exported", xeroReference: "XERO-20260626", site: "Bella Vista", trade: "Pipe Layer" },
  { id: "demo-ds-5", jobCode: "ASR-002", jobName: "Baulkham Hills Earthworks", workerName: "Jake Morrison", workerEmail: "jake.morrison@ausslope.com.au", date: "2026-06-25", startTime: "07:00", finishTime: "16:00", breakMinutes: 30, ordinaryHours: 8, overtimeHours: 0.5, allowances: "Travel, Tool", notes: "Compaction testing and subgrade prep.", approvalStatus: "Rejected", approvedBy: "Tim Kinsella", approvedDate: "2026-06-26", payrollExportStatus: "Not Exported", xeroReference: "", site: "Baulkham Hills", trade: "Concretor" },
];

const DEMO_SUBCONTRACTORS: Subcontractor[] = [
  { id: "demo-sub-1", companyName: "Apex Concrete Solutions", abn: "51 234 567 890", trade: "Concretor", contactName: "Mark Thompson", contactPhone: "0412 345 678", contactEmail: "mark@apexconcrete.com.au", insuranceExpiry: "2027-03-31", licenceExpiry: "2027-06-30", licenceNumber: "CL-48291", prequalificationStatus: "Approved", complianceStatus: "Active", inductionStatus: "Complete", swmsStatus: "Approved", mobilisationApproved: true, activeJobCodes: ["ASR-001", "ASR-002"], notes: "Preferred concretor. Excellent safety record.", insuranceDocUrl: "https://ausslope.sharepoint.com/sites/ProjectControls/Shared%20Documents/Subcontractors/Apex_Insurance_2027.pdf", licenceDocUrl: "https://ausslope.sharepoint.com/sites/ProjectControls/Shared%20Documents/Subcontractors/Apex_Licence_2027.pdf", swmsDocUrl: "https://ausslope.sharepoint.com/sites/ProjectControls/Shared%20Documents/Subcontractors/Apex_SWMS_2026.pdf" },
  { id: "demo-sub-2", companyName: "Precision Earthmoving Pty Ltd", abn: "72 345 678 901", trade: "Plant Operator", contactName: "Steve Holloway", contactPhone: "0423 456 789", contactEmail: "steve@precisionearthmove.com.au", insuranceExpiry: "2026-07-15", licenceExpiry: "2026-08-01", licenceNumber: "PL-29384", prequalificationStatus: "Approved", complianceStatus: "Expiring Soon", inductionStatus: "Complete", swmsStatus: "Approved", mobilisationApproved: true, activeJobCodes: ["ASR-002"], notes: "Insurance expires soon — renewal in progress.", insuranceDocUrl: "https://ausslope.sharepoint.com/sites/ProjectControls/Shared%20Documents/Subcontractors/Precision_Insurance_2026.pdf", licenceDocUrl: "", swmsDocUrl: "https://ausslope.sharepoint.com/sites/ProjectControls/Shared%20Documents/Subcontractors/Precision_SWMS_2026.pdf" },
  { id: "demo-sub-3", companyName: "Blue Sky Drainage", abn: "88 456 789 012", trade: "Pipe Layer", contactName: "Karen Lee", contactPhone: "0434 567 890", contactEmail: "karen@blueskydrainage.com.au", insuranceExpiry: "2025-12-31", licenceExpiry: "2026-02-28", licenceNumber: "DL-11293", prequalificationStatus: "Pending", complianceStatus: "Blocked", inductionStatus: "Pending", swmsStatus: "Pending", mobilisationApproved: false, activeJobCodes: [], notes: "Insurance and licence both expired. Do not mobilise.", insuranceDocUrl: "", licenceDocUrl: "", swmsDocUrl: "" },
  { id: "demo-sub-4", companyName: "SafeForm Pty Ltd", abn: "63 567 890 123", trade: "Steel Fixer", contactName: "Tony Ricci", contactPhone: "0445 678 901", contactEmail: "tony@safeform.com.au", insuranceExpiry: "2027-01-31", licenceExpiry: "2027-04-30", licenceNumber: "SF-38291", prequalificationStatus: "Approved", complianceStatus: "Active", inductionStatus: "Complete", swmsStatus: "Approved", mobilisationApproved: true, activeJobCodes: ["ASR-001"], notes: "", insuranceDocUrl: "https://ausslope.sharepoint.com/sites/ProjectControls/Shared%20Documents/Subcontractors/SafeForm_Insurance_2027.pdf", licenceDocUrl: "https://ausslope.sharepoint.com/sites/ProjectControls/Shared%20Documents/Subcontractors/SafeForm_Licence_2027.pdf", swmsDocUrl: "https://ausslope.sharepoint.com/sites/ProjectControls/Shared%20Documents/Subcontractors/SafeForm_SWMS_2026.pdf" },
  { id: "demo-sub-5", companyName: "Groundworks NSW", abn: "44 678 901 234", trade: "Concretor", contactName: "Paul Denton", contactPhone: "0456 789 012", contactEmail: "paul@groundworksnsw.com.au", insuranceExpiry: "2026-11-30", licenceExpiry: "2026-12-31", licenceNumber: "GW-49201", prequalificationStatus: "Not Started", complianceStatus: "Active", inductionStatus: "Not Started", swmsStatus: "Not Submitted", mobilisationApproved: false, activeJobCodes: [], notes: "New subcontractor — onboarding in progress.", insuranceDocUrl: "", licenceDocUrl: "", swmsDocUrl: "" },
];

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SITE_ID =
  process.env.EXPO_PUBLIC_SHAREPOINT_SITE_ID ??
  // Primary site: asr-operations2 (Jobs, Workers)
  "ausslope.sharepoint.com,aedb5034-4468-4dab-9b6e-94cf53db15be,b3103679-4573-4af6-9a4b-5d7d8bfebe12";

// Secondary site: ProjectControls (Day Sheets, Subcontractors, Site Photos)
const SECONDARY_SITE_ID =
  process.env.EXPO_PUBLIC_SHAREPOINT_SECONDARY_SITE_ID ??
  "ausslope.sharepoint.com,06c7be2e-83f7-4c3a-93fe-54f6542eaa01,b3103679-4573-4af6-9a4b-5d7d8bfebe12";

// M365 group email for Ausslope Job Control (supervisors/managers)
const G12_JOB_CONTROL_GROUP_EMAIL =
  process.env.EXPO_PUBLIC_SUPERVISORS_GROUP_EMAIL ??
  "AusslopeJobControl@ausslope.com.au";

// ─── Types ────────────────────────────────────────────────────────────────────

export interface Job {
  id: string;
  jobNumber: string;
  jobCode: string;
  jobName: string;
  client: string;
  siteAddress: string;
  status: string;
  startDate: string;
  completionDate: string;
  contractValue: string;
  projectManager: string;
  superintendent: string;
  jobType: string;
  description: string;
  priority: string;
}

export interface DaySheet {
  id: string;
  jobCode: string;
  jobName: string;
  workerName: string;
  workerEmail: string;
  date: string;
  startTime: string;
  finishTime: string;
  breakMinutes: number;
  ordinaryHours: number;
  overtimeHours: number;
  allowances: string;
  notes: string;
  approvalStatus: "Pending" | "Approved" | "Rejected";
  approvedBy: string;
  approvedDate: string;
  payrollExportStatus: "Not Exported" | "Exported" | "Error";
  xeroReference: string;
  site: string;
  trade: string;
  photoUrls?: string[];
}

export interface Subcontractor {
  id: string;
  companyName: string;
  abn: string;
  trade: string;
  contactName: string;
  contactPhone: string;
  contactEmail: string;
  insuranceExpiry: string;
  licenceExpiry: string;
  licenceNumber: string;
  prequalificationStatus: string;
  complianceStatus: "Active" | "Expiring Soon" | "Blocked";
  // Legacy fields kept for UI compatibility
  inductionStatus: "Complete" | "Pending" | "Not Started";
  swmsStatus: "Approved" | "Pending" | "Not Submitted";
  mobilisationApproved: boolean;
  activeJobCodes: string[];
  notes: string;
  insuranceDocUrl?: string;
  licenceDocUrl?: string;
  swmsDocUrl?: string;
}

// ─── Core API helper ──────────────────────────────────────────────────────────

async function getToken(): Promise<string> {
  let token = await getStoredToken();
  if (!token) token = await refreshAccessToken();
  if (!token) throw new Error("Not authenticated");
  return token;
}

async function graphFetch(path: string, options?: RequestInit): Promise<Response> {
  const token = await getToken();
  const makeReq = (t: string) =>
    fetch(`${GRAPH_BASE}${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${t}`,
        "Content-Type": "application/json",
        Accept: "application/json",
        ...(options?.headers ?? {}),
      },
    });

  let res = await makeReq(token);
  if (res.status === 401) {
    const newToken = await refreshAccessToken();
    if (!newToken) throw new Error("Session expired. Please sign in again.");
    res = await makeReq(newToken);
  }
  return res;
}

async function getListItems(
  listName: string,
  filter?: string,
  orderBy?: string,
  siteId?: string
): Promise<any[]> {
  const site = siteId ?? SITE_ID;
  let url = `/sites/${site}/lists/${encodeURIComponent(listName)}/items?expand=fields&$top=200`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  if (orderBy) url += `&$orderby=${encodeURIComponent(orderBy)}`;
  const res = await graphFetch(url);
  if (!res.ok) throw new Error(`Failed to fetch ${listName}: ${res.statusText}`);
  const data = await res.json();
  return data.value ?? [];
}

async function createListItem(listName: string, fields: Record<string, any>, siteId?: string): Promise<any> {
  const site = siteId ?? SITE_ID;
  const res = await graphFetch(
    `/sites/${site}/lists/${encodeURIComponent(listName)}/items`,
    { method: "POST", body: JSON.stringify({ fields }) }
  );
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Failed to create ${listName} item: ${err}`);
  }
  return res.json();
}

async function updateListItem(
  listName: string,
  itemId: string,
  fields: Record<string, any>,
  siteId?: string
): Promise<void> {
  const site = siteId ?? SITE_ID;
  const res = await graphFetch(
    `/sites/${site}/lists/${encodeURIComponent(listName)}/items/${itemId}/fields`,
    { method: "PATCH", body: JSON.stringify(fields) }
  );
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Failed to update ${listName} item: ${err}`);
  }
}

// ─── Role Gating ──────────────────────────────────────────────────────────────

/**
 * Check if the current user is a member of the Ausslope Job Control M365 group.
 * Members of this group are treated as supervisors with approval permissions.
 */
export async function checkIsSupervisor(): Promise<boolean> {
  try {
    const res = await graphFetch(
      `/me/memberOf?$select=mail,displayName&$top=100`
    );
    if (!res.ok) return false;
    const data = await res.json();
    const groups: any[] = data.value ?? [];
    return groups.some(
      (g) =>
        g.mail?.toLowerCase() === G12_JOB_CONTROL_GROUP_EMAIL.toLowerCase() ||
        g.displayName === "Ausslope Job Control"
    );
  } catch {
    return false;
  }
}

// ─── Jobs ─────────────────────────────────────────────────────────────────────

function mapJob(raw: any): Job {
  const f = raw.fields ?? {};
  return {
    id: raw.id,
    // Job Register (ProjectControls): Title = job code, JobName = display name
    jobNumber: f.Title ?? "",
    jobCode: f.Title ?? "",
    jobName: f.JobName ?? f.Title ?? "",
    client: f.Client ?? "",
    siteAddress: f.SiteAddress ?? "",
    status: f.Status ?? "Active",
    startDate: f.StartDate ?? "",
    completionDate: f.EndDate ?? "",
    contractValue: f.Value != null ? `$${Number(f.Value).toLocaleString()}` : "",
    projectManager: f.SupervisorLookupValue ?? f.Supervisor ?? "",
    superintendent: f.SupervisorLookupValue ?? f.Supervisor ?? "",
    jobType: f.ProjectType ?? "",
    description: f.Notes ?? "",
    priority: f.Priority ?? "Normal",
  };
}

export async function getJobs(): Promise<Job[]> {
  try {
    // Job Register lives on the SECONDARY site (ProjectControls)
    const items = await getListItems("Job Register", undefined, "fields/Title asc", SECONDARY_SITE_ID);
    if (items.length === 0) return DEMO_JOBS;
    return items.map(mapJob);
  } catch {
    return DEMO_JOBS;
  }
}

export async function getJob(id: string): Promise<Job | null> {
  // Demo data fallback
  const demo = DEMO_JOBS.find((j) => j.id === id);
  if (demo) return demo;
  try {
    const res = await graphFetch(
      `/sites/${SECONDARY_SITE_ID}/lists/${encodeURIComponent("Job Register")}/items/${id}?expand=fields`
    );
    if (!res.ok) return null;
    return mapJob(await res.json());
  } catch {
    return null;
  }
}

export async function createJob(job: Omit<Job, "id">): Promise<Job> {
  const fields: Record<string, any> = {
    // Title = job code (primary key in Job Register)
    Title: job.jobNumber,
    JobName: job.jobName,
    Status: job.status,
    StartDate: job.startDate,
    EndDate: job.completionDate,
    Value: parseFloat(job.contractValue.replace(/[^0-9.]/g, "")) || undefined,
    ProjectType: job.jobType,
  };
  const raw = await createListItem("Job Register", fields, SECONDARY_SITE_ID);
  return mapJob(raw);
}

// ─── Day Sheets ───────────────────────────────────────────────────────────────

function mapDaySheet(raw: any): DaySheet {
  const f = raw.fields ?? {};
  return {
    id: raw.id,
    // Real field names verified 2026-07-21 against Day Sheets list on ProjectControls
    jobCode: f.JobCodeLookupValue ?? f.JobCode ?? "",
    jobName: f.JobCodeLookupValue ?? f.JobCode ?? "",
    workerName: f.SubcontractorCrewLookupValue ?? f.SubcontractorCrew ?? "",
    workerEmail: "",  // Not in list — derived from SubcontractorCrew lookup
    date: f.Date ?? "",
    startTime: "",    // Not in list — captured via OrdinaryHours/OvertimeHours split
    finishTime: "",
    breakMinutes: 0,
    ordinaryHours: Number(f.LabourHours ?? f.OrdinaryHours ?? 0),
    overtimeHours: Number(f.OvertimeHours ?? 0),
    allowances: "",
    notes: f.WorkDescription ?? "",
    approvalStatus: (f.Status as DaySheet["approvalStatus"]) ?? "Pending",
    approvedBy: f.ApprovedByLookupValue ?? f.ApprovedBy ?? "",
    approvedDate: f.Modified ?? "",
    payrollExportStatus: (f.PayrollExportStatus as DaySheet["payrollExportStatus"]) ?? "Not Exported",
    xeroReference: f.XeroReference ?? "",
    site: f.JobCodeLookupValue ?? "",
    trade: f.CostCode ?? "",
  };
}

export async function getDaySheets(filter?: string): Promise<DaySheet[]> {
  try {
    const items = await getListItems("Day Sheets", filter, "fields/Date desc", SECONDARY_SITE_ID);
    if (items.length === 0) return DEMO_DAYSHEETS;
    return items.map(mapDaySheet);
  } catch {
    return DEMO_DAYSHEETS;
  }
}

export async function createDaySheet(
  sheet: Omit<DaySheet, "id" | "photoUrls">
): Promise<DaySheet> {
  const fields: Record<string, any> = {
    // Real field names verified 2026-07-21 against Day Sheets list on ProjectControls
    // Title (Reference) = auto-generated or set to a meaningful reference
    Title: `${sheet.jobCode}-${sheet.date}`,
    // JobCode is a Lookup to Job Register — pass the lookup ID if available, else text
    JobCode: sheet.jobCode,
    Date: sheet.date,
    // SubcontractorCrew is a Lookup to Subcontractor Register
    SubcontractorCrew: sheet.workerName,
    LabourHours: sheet.ordinaryHours,
    OrdinaryHours: sheet.ordinaryHours,
    OvertimeHours: sheet.overtimeHours,
    WorkDescription: sheet.notes,
    CostCode: sheet.trade,
    Status: "Draft",
    PayrollExportStatus: "Not Exported",
  };
  const raw = await createListItem("Day Sheets", fields, SECONDARY_SITE_ID);
  return mapDaySheet(raw);
}

export async function approveDaySheet(id: string, approvedBy: string): Promise<void> {
  await updateListItem("Day Sheets", id, {
    Status: "Approved",
  }, SECONDARY_SITE_ID);
}

export async function rejectDaySheet(id: string, approvedBy: string, reason?: string): Promise<void> {
  await updateListItem("Day Sheets", id, {
    Status: "Rejected",
    WorkDescription: reason ? `REJECTED: ${reason}` : undefined,
  }, SECONDARY_SITE_ID);
}

/**
 * Mark one or more day sheets as exported in SharePoint.
 * Sets Payroll_Export_Status to "Exported" and records a Xero reference.
 *
 * @param ids - Array of SharePoint item IDs to mark exported
 * @param xeroReference - Optional reference string (e.g. batch date or Xero batch ID)
 */
export async function markDaySheetsExported(
  ids: string[],
  xeroReference?: string
): Promise<{ success: number; failed: number }> {
  const ref = xeroReference ?? `XERO-${new Date().toISOString().split("T")[0]}`;
  let success = 0;
  let failed = 0;
  await Promise.allSettled(
    ids.map(async (id) => {
      try {
        // Skip demo data — IDs starting with "demo-" are not in SharePoint
        if (id.startsWith("demo-")) {
          success++;
          return;
        }
        await updateListItem("Day Sheets", id, {
          PayrollExportStatus: "Exported",
          XeroReference: ref,
        }, SECONDARY_SITE_ID);
        success++;
      } catch {
        failed++;
      }
    })
  );
  return { success, failed };
}

// ─── Photo Upload ─────────────────────────────────────────────────────────────

/**
 * Upload a photo to the "Site Photos" document library in SharePoint.
 * Associates the photo with a job code and day sheet via the file name.
 *
 * @param localUri - Local file URI from expo-image-picker
 * @param jobCode - Job code to associate the photo with
 * @param workerName - Worker name for file naming
 * @param date - Date string (YYYY-MM-DD)
 * @returns SharePoint web URL of the uploaded file
 */
export async function uploadSitePhoto(
  localUri: string,
  jobCode: string,
  workerName: string,
  date: string
): Promise<string> {
  const token = await getToken();

  // Read the file as base64
  let base64Data: string;
  if (Platform.OS === "web") {
    // Web: fetch the blob URL
    const response = await fetch(localUri);
    const blob = await response.blob();
    base64Data = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const result = reader.result as string;
        resolve(result.split(",")[1]);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } else {
    base64Data = await FileSystem.readAsStringAsync(localUri, {
      encoding: FileSystem.EncodingType.Base64,
    });
  }

  // Build a unique file name
  const timestamp = Date.now();
  const safeName = workerName.replace(/[^a-zA-Z0-9]/g, "_");
  const safeJobCode = jobCode.replace(/[^a-zA-Z0-9]/g, "_");
  const fileName = `${safeJobCode}_${safeName}_${date}_${timestamp}.jpg`;

  // Upload to Site Photos library
  const uploadUrl = `/sites/${SECONDARY_SITE_ID}/drives`;
  // First get the drive ID for Site Photos
  const drivesRes = await graphFetch(uploadUrl);
  if (!drivesRes.ok) throw new Error("Failed to get drives");
  const drives = await drivesRes.json();
  const sitePhotosDrive = drives.value?.find(
    (d: any) => d.name === "Site Photos" || d.webUrl?.includes("Site%20Photos")
  );

  let uploadPath: string;
  if (sitePhotosDrive) {
    uploadPath = `/drives/${sitePhotosDrive.id}/root:/${fileName}:/content`;
  } else {
    // Fallback: upload to default document library
    uploadPath = `/sites/${SECONDARY_SITE_ID}/drive/root:/${fileName}:/content`;
  }

  const binaryData = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
  const uploadRes = await fetch(`${GRAPH_BASE}${uploadPath}`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "image/jpeg",
    },
    body: binaryData,
  });

  if (!uploadRes.ok) {
    const err = await uploadRes.text();
    throw new Error(`Failed to upload photo: ${err}`);
  }

  const uploaded = await uploadRes.json();
  return uploaded.webUrl ?? uploaded["@microsoft.graph.downloadUrl"] ?? "";
}

// ─── Subcontractors ───────────────────────────────────────────────────────────

function calcComplianceStatus(
  insuranceExpiry: string,
  licenceExpiry: string
): Subcontractor["complianceStatus"] {
  const today = new Date();
  const ins = new Date(insuranceExpiry);
  const lic = new Date(licenceExpiry);
  const daysIns = (ins.getTime() - today.getTime()) / 86400000;
  const daysLic = (lic.getTime() - today.getTime()) / 86400000;
  if (
    (insuranceExpiry && daysIns < 0) ||
    (licenceExpiry && daysLic < 0)
  ) return "Blocked";
  if (
    (insuranceExpiry && daysIns < 30) ||
    (licenceExpiry && daysLic < 30)
  ) return "Expiring Soon";
  return "Active";
}

function mapSubcontractor(raw: any): Subcontractor {
  const f = raw.fields ?? {};
  // Derive compliance from insurance expiry dates (real fields verified 2026-07-21)
  const plExpiry = f.PublicLiabilityExpiry ?? "";
  const wcExpiry = f.WorkersCompExpiry ?? "";
  const piExpiry = f.ProfessionalIndemnityExpiry ?? "";
  const worstExpiry = [plExpiry, wcExpiry, piExpiry].filter(Boolean).sort()[0] ?? "";
  return {
    id: raw.id,
    // Real field names verified 2026-07-21 against Subcontractor Register on ProjectControls
    companyName: f.Title ?? "",
    abn: f.ABN ?? "",
    trade: f.TradeScope ?? "",
    contactName: f.PrimaryContact ?? "",
    contactPhone: "",  // Not in list — use PrimaryContact text field
    contactEmail: "",  // Not in list — use PrimaryContact text field
    insuranceExpiry: plExpiry,
    licenceExpiry: wcExpiry,
    licenceNumber: "",  // Not in list
    prequalificationStatus: f.PrequalStatus ?? "Not Started",
    complianceStatus: calcComplianceStatus(plExpiry, worstExpiry),
    inductionStatus: (f.InductionStatus as Subcontractor["inductionStatus"]) ?? "Not Started",
    swmsStatus: (f.SWMSStatus as Subcontractor["swmsStatus"]) ?? "Not Submitted",
    mobilisationApproved: f.Active === true || f.Active === "Yes",
    activeJobCodes: Array.isArray(f.LinkedJobs)
      ? f.LinkedJobs.map((l: any) => l.LookupValue ?? "").filter(Boolean)
      : [],
    notes: "",
  };
}

export async function getSubcontractors(): Promise<Subcontractor[]> {
  try {
    const items = await getListItems(
      "Subcontractor Register",
      undefined,
      "fields/Title asc",
      SECONDARY_SITE_ID
    );
    if (items.length === 0) return DEMO_SUBCONTRACTORS;
    return items.map(mapSubcontractor);
  } catch {
    return DEMO_SUBCONTRACTORS;
  }
}

export async function getSubcontractor(id: string): Promise<Subcontractor | null> {
  // Demo data fallback
  const demo = DEMO_SUBCONTRACTORS.find((s) => s.id === id);
  if (demo) return demo;
  try {
    const res = await graphFetch(
      `/sites/${SECONDARY_SITE_ID}/lists/${encodeURIComponent("Subcontractor Register")}/items/${id}?expand=fields`
    );
    if (!res.ok) return null;
    return mapSubcontractor(await res.json());
  } catch {
    return null;
  }
}

export async function createSubcontractor(
  sub: Omit<Subcontractor, "id" | "complianceStatus">
): Promise<Subcontractor> {
  const fields: Record<string, any> = {
    // Real field names verified 2026-07-21 against Subcontractor Register on ProjectControls
    Title: sub.companyName,
    ABN: sub.abn,
    TradeScope: sub.trade,
    PrimaryContact: sub.contactName,
    PublicLiabilityExpiry: sub.insuranceExpiry || undefined,
    WorkersCompExpiry: sub.licenceExpiry || undefined,
    PrequalStatus: sub.prequalificationStatus,
    InductionStatus: sub.inductionStatus,
    SWMSStatus: sub.swmsStatus,
    Active: sub.mobilisationApproved,
  };
  const raw = await createListItem("Subcontractor Register", fields, SECONDARY_SITE_ID);
  return mapSubcontractor(raw);
}

export async function updateSubcontractorMobilisation(
  id: string,
  approved: boolean
): Promise<void> {
  await updateListItem("Subcontractor Register", id, {
    Active: approved,
  }, SECONDARY_SITE_ID);
}

export async function updateSubcontractorPrequal(
  id: string,
  status: string
): Promise<void> {
  await updateListItem("Subcontractor Register", id, {
    PrequalStatus: status,
  }, SECONDARY_SITE_ID);
}

// ─── Lookup helpers for form pickers ─────────────────────────────────────────

/**
 * Returns a lightweight list of {id, label} pairs from the Subcontractor Register
 * for use in the Day Sheet "Subcontractor / Crew" picker.
 * Falls back to demo data if the list is empty or unavailable.
 */
export async function getSubcontractorList(): Promise<{ id: string; label: string }[]> {
  try {
    const items = await getListItems(
      "Subcontractor Register",
      "fields/Active eq true",
      "fields/Title asc",
      SECONDARY_SITE_ID
    );
    if (items.length === 0) {
      return DEMO_SUBCONTRACTORS.map((s) => ({ id: s.id, label: s.companyName }));
    }
    return items.map((item: any) => ({
      id: String(item.id),
      label: item.fields?.Title ?? "",
    }));
  } catch {
    return DEMO_SUBCONTRACTORS.map((s) => ({ id: s.id, label: s.companyName }));
  }
}

/**
 * Returns a lightweight list of {id, code, name} from the Job Register
 * for use in the Day Sheet "Job Code" picker.
 * Falls back to demo data if the list is empty or unavailable.
 */
export async function getJobList(): Promise<{ id: string; code: string; name: string }[]> {
  try {
    const items = await getListItems(
      "Job Register",
      "fields/Status ne 'Completed'",
      "fields/Title asc",
      SECONDARY_SITE_ID
    );
    if (items.length === 0) {
      return DEMO_JOBS.map((j) => ({ id: j.id, code: j.jobCode, name: j.jobName }));
    }
    return items.map((item: any) => ({
      id: String(item.id),
      code: item.fields?.Title ?? "",
      name: item.fields?.JobName ?? item.fields?.Title ?? "",
    }));
  } catch {
    return DEMO_JOBS.map((j) => ({ id: j.id, code: j.jobCode, name: j.jobName }));
  }
}

/**
 * Returns cost code options from the Cost Codes list on ProjectControls.
 * Falls back to a hardcoded list if the SharePoint list is empty.
 */
export async function getCostCodes(): Promise<string[]> {
  const FALLBACK_COST_CODES = [
    "01 - Preliminaries",
    "02 - Earthworks",
    "03 - Concrete",
    "04 - Steel",
    "05 - Drainage",
    "06 - Pavements",
    "07 - Landscaping",
    "08 - Traffic Control",
    "09 - Plant & Equipment",
    "10 - Labour",
    "11 - Materials",
    "12 - Subcontractors",
    "99 - Variations",
  ];
  try {
    const items = await getListItems("Cost Codes", undefined, "fields/Title asc", SECONDARY_SITE_ID);
    if (items.length === 0) return FALLBACK_COST_CODES;
    return items.map((item: any) => item.fields?.Title ?? "").filter(Boolean);
  } catch {
    return FALLBACK_COST_CODES;
  }
}
