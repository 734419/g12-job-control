/**
 * SharePoint Graph API service for G12 Job Control.
 *
 * Site: https://g12group.sharepoint.com/sites/G12JobControl
 * Site ID: g12group.sharepoint.com,350ccf6f-27ad-4edc-b7a2-e295b71d3ab2,ea9063e3-c201-47fb-90c0-2f44b60a52b4
 *
 * Real SharePoint list names (verified via Graph API 2026-06-27):
 *   - "Job Register"          → Jobs
 *   - "Day Sheets"            → Day Sheets (created by app on first run if missing)
 *   - "Subcontractor Register" → Subcontractors
 *   - "Site Photos"           → Document library for photo uploads
 *
 * Column internal names (verified via Graph API):
 *   Job Register: Title, Job_x0020_Number*, Client, Site_x0020_Address, Job_x0020_Status,
 *                 Start_x0020_Date, Completion_x0020_Date, Contract_x0020_Value,
 *                 Project_x0020_Manager, Job_x0020_Type, Job_x0020_Description,
 *                 Superintendent, Priority, Jobcode
 *   Subcontractor Register: Title*, Company_x0020_Name*, Trade, Contact_x0020_Email,
 *                 Contact_x0020_Name, Contact_x0020_Phone, ABN, Insurance_x0020_Expiry,
 *                 Licence_x0020_Expiry, Licence_x0020_Number, Prequalification_x0020_Status
 *   Day Sheets: Title, Jobcode*, Worker_x0020_Name*, Worker_x0020_Email, Work_x0020_Date*,
 *               Start_x0020_Time, Finish_x0020_Time, Break_x0020_Minutes, Ordinary_x0020_Hours,
 *               Overtime_x0020_Hours, Allowances, Notes, Approval_x0020_Status,
 *               Approved_x0020_By, Approved_x0020_Date, Payroll_x0020_Export_x0020_Status,
 *               Xero_x0020_Reference, Site, Trade, Job_x0020_Name
 *
 * M365 Group for role gating:
 *   "G12 Job Control" → G12JobControl@g12consulting.com.au
 */

import { getStoredToken, refreshAccessToken } from "@/lib/auth/microsoft";
import * as FileSystem from "expo-file-system/legacy";
import { Platform } from "react-native";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SITE_ID =
  "g12group.sharepoint.com,350ccf6f-27ad-4edc-b7a2-e295b71d3ab2,ea9063e3-c201-47fb-90c0-2f44b60a52b4";

// M365 group ID for G12 Job Control (supervisors/managers)
// Fetched via Graph API: G12JobControl@g12consulting.com.au
const G12_JOB_CONTROL_GROUP_EMAIL = "G12JobControl@g12consulting.com.au";

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
  orderBy?: string
): Promise<any[]> {
  let url = `/sites/${SITE_ID}/lists/${encodeURIComponent(listName)}/items?expand=fields&$top=200`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  if (orderBy) url += `&$orderby=${encodeURIComponent(orderBy)}`;
  const res = await graphFetch(url);
  if (!res.ok) throw new Error(`Failed to fetch ${listName}: ${res.statusText}`);
  const data = await res.json();
  return data.value ?? [];
}

async function createListItem(listName: string, fields: Record<string, any>): Promise<any> {
  const res = await graphFetch(
    `/sites/${SITE_ID}/lists/${encodeURIComponent(listName)}/items`,
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
  fields: Record<string, any>
): Promise<void> {
  const res = await graphFetch(
    `/sites/${SITE_ID}/lists/${encodeURIComponent(listName)}/items/${itemId}/fields`,
    { method: "PATCH", body: JSON.stringify(fields) }
  );
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Failed to update ${listName} item: ${err}`);
  }
}

// ─── Role Gating ──────────────────────────────────────────────────────────────

/**
 * Check if the current user is a member of the G12 Job Control M365 group.
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
        g.displayName === "G12 Job Control"
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
    jobNumber: f.Job_x0020_Number ?? "",
    jobCode: f.Jobcode ?? f.Job_x0020_Number ?? "",
    jobName: f.Title ?? "",
    client: f.Client ?? "",
    siteAddress: f.Site_x0020_Address ?? "",
    status: f.Job_x0020_Status ?? "Active",
    startDate: f.Start_x0020_Date ?? "",
    completionDate: f.Completion_x0020_Date ?? "",
    contractValue: f.Contract_x0020_Value ?? "",
    projectManager: f.Project_x0020_Manager ?? "",
    superintendent: f.Superintendent ?? "",
    jobType: f.Job_x0020_Type ?? "",
    description: f.Job_x0020_Description ?? "",
    priority: f.Priority ?? "Normal",
  };
}

export async function getJobs(): Promise<Job[]> {
  const items = await getListItems("Job Register", undefined, "fields/Job_x0020_Number asc");
  return items.map(mapJob);
}

export async function getJob(id: string): Promise<Job | null> {
  const res = await graphFetch(
    `/sites/${SITE_ID}/lists/${encodeURIComponent("Job Register")}/items/${id}?expand=fields`
  );
  if (!res.ok) return null;
  return mapJob(await res.json());
}

// ─── Day Sheets ───────────────────────────────────────────────────────────────

function mapDaySheet(raw: any): DaySheet {
  const f = raw.fields ?? {};
  return {
    id: raw.id,
    jobCode: f.Jobcode ?? "",
    jobName: f.Job_x0020_Name ?? "",
    workerName: f.Worker_x0020_Name ?? "",
    workerEmail: f.Worker_x0020_Email ?? "",
    date: f.Work_x0020_Date ?? "",
    startTime: f.Start_x0020_Time ?? "",
    finishTime: f.Finish_x0020_Time ?? "",
    breakMinutes: Number(f.Break_x0020_Minutes ?? 0),
    ordinaryHours: Number(f.Ordinary_x0020_Hours ?? 0),
    overtimeHours: Number(f.Overtime_x0020_Hours ?? 0),
    allowances: f.Allowances ?? "",
    notes: f.Notes ?? "",
    approvalStatus: (f.Approval_x0020_Status as DaySheet["approvalStatus"]) ?? "Pending",
    approvedBy: f.Approved_x0020_By ?? "",
    approvedDate: f.Approved_x0020_Date ?? "",
    payrollExportStatus:
      (f.Payroll_x0020_Export_x0020_Status as DaySheet["payrollExportStatus"]) ?? "Not Exported",
    xeroReference: f.Xero_x0020_Reference ?? "",
    site: f.Site ?? "",
    trade: f.Trade ?? "",
  };
}

export async function getDaySheets(filter?: string): Promise<DaySheet[]> {
  const items = await getListItems("Day Sheets", filter, "fields/Work_x0020_Date desc");
  return items.map(mapDaySheet);
}

export async function createDaySheet(
  sheet: Omit<DaySheet, "id" | "photoUrls">
): Promise<DaySheet> {
  const fields: Record<string, any> = {
    Title: `${sheet.jobCode} - ${sheet.workerName} - ${sheet.date}`,
    Jobcode: sheet.jobCode,
    Job_x0020_Name: sheet.jobName,
    Worker_x0020_Name: sheet.workerName,
    Worker_x0020_Email: sheet.workerEmail,
    Work_x0020_Date: sheet.date,
    Start_x0020_Time: sheet.startTime,
    Finish_x0020_Time: sheet.finishTime,
    Break_x0020_Minutes: sheet.breakMinutes,
    Ordinary_x0020_Hours: sheet.ordinaryHours,
    Overtime_x0020_Hours: sheet.overtimeHours,
    Allowances: sheet.allowances,
    Notes: sheet.notes,
    Approval_x0020_Status: "Pending",
    Payroll_x0020_Export_x0020_Status: "Not Exported",
    Site: sheet.site,
    Trade: sheet.trade,
  };
  const raw = await createListItem("Day Sheets", fields);
  return mapDaySheet(raw);
}

export async function approveDaySheet(id: string, approvedBy: string): Promise<void> {
  await updateListItem("Day Sheets", id, {
    Approval_x0020_Status: "Approved",
    Approved_x0020_By: approvedBy,
    Approved_x0020_Date: new Date().toISOString().split("T")[0],
  });
}

export async function rejectDaySheet(id: string, approvedBy: string, reason?: string): Promise<void> {
  await updateListItem("Day Sheets", id, {
    Approval_x0020_Status: "Rejected",
    Approved_x0020_By: approvedBy,
    Approved_x0020_Date: new Date().toISOString().split("T")[0],
    Notes: reason ? `REJECTED: ${reason}` : undefined,
  });
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
  const uploadUrl = `/sites/${SITE_ID}/drives`;
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
    uploadPath = `/sites/${SITE_ID}/drive/root:/${fileName}:/content`;
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
  return {
    id: raw.id,
    companyName: f.Company_x0020_Name ?? f.Title ?? "",
    abn: f.ABN ?? "",
    trade: f.Trade ?? "",
    contactName: f.Contact_x0020_Name ?? "",
    contactPhone: f.Contact_x0020_Phone ?? "",
    contactEmail: f.Contact_x0020_Email ?? "",
    insuranceExpiry: f.Insurance_x0020_Expiry ?? "",
    licenceExpiry: f.Licence_x0020_Expiry ?? "",
    licenceNumber: f.Licence_x0020_Number ?? "",
    prequalificationStatus: f.Prequalification_x0020_Status ?? "Not Started",
    complianceStatus: calcComplianceStatus(
      f.Insurance_x0020_Expiry ?? "",
      f.Licence_x0020_Expiry ?? ""
    ),
    // Legacy/UI fields - not in SharePoint, managed locally or via future columns
    inductionStatus: (f.Induction_x0020_Status as Subcontractor["inductionStatus"]) ?? "Not Started",
    swmsStatus: (f.SWMS_x0020_Status as Subcontractor["swmsStatus"]) ?? "Not Submitted",
    mobilisationApproved: f.Mobilisation_x0020_Approved === true || f.Mobilisation_x0020_Approved === "Yes",
    activeJobCodes: (f.Active_x0020_Job_x0020_Codes ?? "").split(";").map((s: string) => s.trim()).filter(Boolean),
    notes: f.Notes ?? "",
  };
}

export async function getSubcontractors(): Promise<Subcontractor[]> {
  const items = await getListItems(
    "Subcontractor Register",
    undefined,
    "fields/Company_x0020_Name asc"
  );
  return items.map(mapSubcontractor);
}

export async function getSubcontractor(id: string): Promise<Subcontractor | null> {
  const res = await graphFetch(
    `/sites/${SITE_ID}/lists/${encodeURIComponent("Subcontractor Register")}/items/${id}?expand=fields`
  );
  if (!res.ok) return null;
  return mapSubcontractor(await res.json());
}

export async function createSubcontractor(
  sub: Omit<Subcontractor, "id" | "complianceStatus">
): Promise<Subcontractor> {
  const fields: Record<string, any> = {
    Title: sub.companyName,
    Company_x0020_Name: sub.companyName,
    ABN: sub.abn,
    Trade: sub.trade,
    Contact_x0020_Name: sub.contactName,
    Contact_x0020_Phone: sub.contactPhone,
    Contact_x0020_Email: sub.contactEmail,
    Insurance_x0020_Expiry: sub.insuranceExpiry,
    Licence_x0020_Expiry: sub.licenceExpiry,
    Licence_x0020_Number: sub.licenceNumber,
    Prequalification_x0020_Status: sub.prequalificationStatus,
  };
  const raw = await createListItem("Subcontractor Register", fields);
  return mapSubcontractor(raw);
}

export async function updateSubcontractorMobilisation(
  id: string,
  approved: boolean
): Promise<void> {
  await updateListItem("Subcontractor Register", id, {
    Mobilisation_x0020_Approved: approved ? "Yes" : "No",
  });
}

export async function updateSubcontractorPrequal(
  id: string,
  status: string
): Promise<void> {
  await updateListItem("Subcontractor Register", id, {
    Prequalification_x0020_Status: status,
  });
}
