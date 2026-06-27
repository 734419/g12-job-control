/**
 * SharePoint Graph API service for G12 Job Control.
 *
 * Connects to: https://g12group.sharepoint.com/sites/G12JobControl
 * Uses Microsoft Graph API v1.0 to read/write SharePoint lists.
 *
 * List names expected in the SharePoint site:
 *   - Jobs
 *   - DaySheets
 *   - Subcontractors
 */

import { getStoredToken, refreshAccessToken } from "@/lib/auth/microsoft";

const SITE_URL = "https://g12group.sharepoint.com/sites/G12JobControl";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// Cache the site ID to avoid repeated lookups
let _siteId: string | null = null;

// ─── Types ────────────────────────────────────────────────────────────────────

export interface Job {
  id: string;
  jobCode: string;
  jobName: string;
  client: string;
  siteAddress: string;
  status: "Active" | "Completed" | "On Hold" | "Cancelled";
  startDate: string;
  supervisor: string;
  description?: string;
}

export interface DaySheet {
  id: string;
  jobCode: string;
  workerName: string;
  workerEmail: string;
  date: string;
  startTime: string;
  finishTime: string;
  breakMinutes: number;
  ordinaryHours: number;
  overtimeHours: number;
  allowances: string[];
  notes?: string;
  approvalStatus: "Pending" | "Approved" | "Rejected";
  approvedBy?: string;
  approvedDate?: string;
  payrollExportStatus: "Not Exported" | "Exported";
  xeroReference?: string;
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
  inductionStatus: "Complete" | "Pending" | "Not Started";
  swmsStatus: "Approved" | "Pending" | "Not Submitted";
  prequalificationStatus: "Approved" | "Pending" | "Not Started";
  mobilisationApproved: boolean;
  activeJobCodes: string[];
  complianceStatus: "Active" | "Expiring Soon" | "Blocked";
  notes?: string;
}

// ─── Core API helper ──────────────────────────────────────────────────────────

async function getToken(): Promise<string> {
  let token = await getStoredToken();
  if (!token) {
    token = await refreshAccessToken();
  }
  if (!token) throw new Error("Not authenticated");
  return token;
}

async function graphFetch(path: string, options?: RequestInit): Promise<Response> {
  const token = await getToken();
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
      ...(options?.headers ?? {}),
    },
  });

  if (res.status === 401) {
    // Try refresh once
    const newToken = await refreshAccessToken();
    if (!newToken) throw new Error("Session expired. Please sign in again.");
    return fetch(`${GRAPH_BASE}${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${newToken}`,
        "Content-Type": "application/json",
        Accept: "application/json",
        ...(options?.headers ?? {}),
      },
    });
  }

  return res;
}

async function getSiteId(): Promise<string> {
  if (_siteId) return _siteId;
  const hostname = "g12group.sharepoint.com";
  const sitePath = "/sites/G12JobControl";
  const res = await graphFetch(`/sites/${hostname}:${sitePath}`);
  if (!res.ok) throw new Error(`Failed to get site ID: ${res.statusText}`);
  const data = await res.json();
  _siteId = data.id;
  return _siteId!;
}

async function getListItems(listName: string, select?: string, filter?: string): Promise<any[]> {
  const siteId = await getSiteId();
  let url = `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items?expand=fields`;
  if (select) url += `&$select=${select}`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  const res = await graphFetch(url);
  if (!res.ok) throw new Error(`Failed to fetch ${listName}: ${res.statusText}`);
  const data = await res.json();
  return data.value ?? [];
}

async function createListItem(listName: string, fields: Record<string, any>): Promise<any> {
  const siteId = await getSiteId();
  const res = await graphFetch(
    `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`,
    {
      method: "POST",
      body: JSON.stringify({ fields }),
    }
  );
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Failed to create ${listName} item: ${err}`);
  }
  return res.json();
}

async function updateListItem(listName: string, itemId: string, fields: Record<string, any>): Promise<void> {
  const siteId = await getSiteId();
  const res = await graphFetch(
    `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items/${itemId}/fields`,
    {
      method: "PATCH",
      body: JSON.stringify(fields),
    }
  );
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Failed to update ${listName} item: ${err}`);
  }
}

// ─── Jobs ─────────────────────────────────────────────────────────────────────

function mapJob(raw: any): Job {
  const f = raw.fields ?? {};
  return {
    id: raw.id,
    jobCode: f.JobCode ?? f.Title ?? "",
    jobName: f.JobName ?? f.Title ?? "",
    client: f.Client ?? "",
    siteAddress: f.SiteAddress ?? "",
    status: f.Status ?? "Active",
    startDate: f.StartDate ?? "",
    supervisor: f.Supervisor ?? "",
    description: f.Description ?? "",
  };
}

export async function getJobs(): Promise<Job[]> {
  const items = await getListItems("Jobs");
  return items.map(mapJob);
}

export async function getJob(id: string): Promise<Job | null> {
  const siteId = await getSiteId();
  const res = await graphFetch(
    `/sites/${siteId}/lists/Jobs/items/${id}?expand=fields`
  );
  if (!res.ok) return null;
  const raw = await res.json();
  return mapJob(raw);
}

// ─── Day Sheets ───────────────────────────────────────────────────────────────

function mapDaySheet(raw: any): DaySheet {
  const f = raw.fields ?? {};
  const allowancesRaw = f.Allowances ?? "";
  const allowances = allowancesRaw
    ? allowancesRaw.split(";").map((s: string) => s.trim()).filter(Boolean)
    : [];
  return {
    id: raw.id,
    jobCode: f.JobCode ?? "",
    workerName: f.WorkerName ?? "",
    workerEmail: f.WorkerEmail ?? "",
    date: f.Date ?? "",
    startTime: f.StartTime ?? "",
    finishTime: f.FinishTime ?? "",
    breakMinutes: Number(f.BreakMinutes ?? 0),
    ordinaryHours: Number(f.OrdinaryHours ?? 0),
    overtimeHours: Number(f.OvertimeHours ?? 0),
    allowances,
    notes: f.Notes ?? "",
    approvalStatus: f.ApprovalStatus ?? "Pending",
    approvedBy: f.ApprovedBy ?? "",
    approvedDate: f.ApprovedDate ?? "",
    payrollExportStatus: f.PayrollExportStatus ?? "Not Exported",
    xeroReference: f.XeroReference ?? "",
  };
}

export async function getDaySheets(filter?: string): Promise<DaySheet[]> {
  const items = await getListItems("DaySheets", undefined, filter);
  return items.map(mapDaySheet);
}

export async function createDaySheet(sheet: Omit<DaySheet, "id">): Promise<DaySheet> {
  const fields: Record<string, any> = {
    Title: `${sheet.jobCode} - ${sheet.workerName} - ${sheet.date}`,
    JobCode: sheet.jobCode,
    WorkerName: sheet.workerName,
    WorkerEmail: sheet.workerEmail,
    Date: sheet.date,
    StartTime: sheet.startTime,
    FinishTime: sheet.finishTime,
    BreakMinutes: sheet.breakMinutes,
    OrdinaryHours: sheet.ordinaryHours,
    OvertimeHours: sheet.overtimeHours,
    Allowances: sheet.allowances.join("; "),
    Notes: sheet.notes ?? "",
    ApprovalStatus: "Pending",
    PayrollExportStatus: "Not Exported",
  };
  const raw = await createListItem("DaySheets", fields);
  return mapDaySheet(raw);
}

export async function approveDaySheet(id: string, approvedBy: string): Promise<void> {
  await updateListItem("DaySheets", id, {
    ApprovalStatus: "Approved",
    ApprovedBy: approvedBy,
    ApprovedDate: new Date().toISOString(),
  });
}

export async function rejectDaySheet(id: string, approvedBy: string): Promise<void> {
  await updateListItem("DaySheets", id, {
    ApprovalStatus: "Rejected",
    ApprovedBy: approvedBy,
    ApprovedDate: new Date().toISOString(),
  });
}

// ─── Subcontractors ───────────────────────────────────────────────────────────

function mapSubcontractor(raw: any): Subcontractor {
  const f = raw.fields ?? {};
  const jobCodesRaw = f.ActiveJobCodes ?? "";
  const activeJobCodes = jobCodesRaw
    ? jobCodesRaw.split(";").map((s: string) => s.trim()).filter(Boolean)
    : [];

  const today = new Date();
  const insuranceExpiry = new Date(f.InsuranceExpiry ?? "");
  const licenceExpiry = new Date(f.LicenceExpiry ?? "");
  const daysToInsurance = (insuranceExpiry.getTime() - today.getTime()) / (1000 * 60 * 60 * 24);
  const daysToLicence = (licenceExpiry.getTime() - today.getTime()) / (1000 * 60 * 60 * 24);

  let complianceStatus: Subcontractor["complianceStatus"] = "Active";
  if (daysToInsurance < 0 || daysToLicence < 0) {
    complianceStatus = "Blocked";
  } else if (daysToInsurance < 30 || daysToLicence < 30) {
    complianceStatus = "Expiring Soon";
  }

  return {
    id: raw.id,
    companyName: f.CompanyName ?? f.Title ?? "",
    abn: f.ABN ?? "",
    trade: f.Trade ?? "",
    contactName: f.ContactName ?? "",
    contactPhone: f.ContactPhone ?? "",
    contactEmail: f.ContactEmail ?? "",
    insuranceExpiry: f.InsuranceExpiry ?? "",
    licenceExpiry: f.LicenceExpiry ?? "",
    inductionStatus: f.InductionStatus ?? "Not Started",
    swmsStatus: f.SWMSStatus ?? "Not Submitted",
    prequalificationStatus: f.PrequalificationStatus ?? "Not Started",
    mobilisationApproved: f.MobilisationApproved === true || f.MobilisationApproved === "Yes",
    activeJobCodes,
    complianceStatus,
    notes: f.Notes ?? "",
  };
}

export async function getSubcontractors(): Promise<Subcontractor[]> {
  const items = await getListItems("Subcontractors");
  return items.map(mapSubcontractor);
}

export async function updateSubcontractorMobilisation(
  id: string,
  approved: boolean
): Promise<void> {
  await updateListItem("Subcontractors", id, {
    MobilisationApproved: approved ? "Yes" : "No",
  });
}

export async function createSubcontractor(
  sub: Omit<Subcontractor, "id" | "complianceStatus">
): Promise<Subcontractor> {
  const fields: Record<string, any> = {
    Title: sub.companyName,
    CompanyName: sub.companyName,
    ABN: sub.abn,
    Trade: sub.trade,
    ContactName: sub.contactName,
    ContactPhone: sub.contactPhone,
    ContactEmail: sub.contactEmail,
    InsuranceExpiry: sub.insuranceExpiry,
    LicenceExpiry: sub.licenceExpiry,
    InductionStatus: sub.inductionStatus,
    SWMSStatus: sub.swmsStatus,
    PrequalificationStatus: sub.prequalificationStatus,
    MobilisationApproved: sub.mobilisationApproved ? "Yes" : "No",
    ActiveJobCodes: sub.activeJobCodes.join("; "),
    Notes: sub.notes ?? "",
  };
  const raw = await createListItem("Subcontractors", fields);
  return mapSubcontractor(raw);
}
