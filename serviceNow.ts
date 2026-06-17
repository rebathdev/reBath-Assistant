// serviceNow.ts
// Creates real incidents via the ServiceNow Table API.
//
// Required environment variables:
//   SERVICENOW_INSTANCE   e.g. "yourcompany" (just the subdomain) OR a full URL
//                         like "https://yourcompany.service-now.com"
//   SERVICENOW_USERNAME   integration user with itil/incident_manager rights
//   SERVICENOW_PASSWORD   password for that user
//
// Optional:
//   SERVICENOW_DEFAULT_CALLER_ID   sys_id of a default caller (fallback when the
//                                  Teams user can't be matched to a ServiceNow user)
//   SERVICENOW_ASSIGNMENT_GROUP    sys_id of the group tickets should land in
//
// Auth note: this uses HTTP Basic auth for simplicity. For production, prefer an
// OAuth client-credentials flow against ServiceNow and swap getAuthHeader() to mint
// a bearer token. Basic auth is acceptable behind HTTPS with a locked-down
// integration account, but OAuth is the better long-term posture.

import { redactForStorage } from "./security";

export interface ServiceNowIncidentInput {
  shortDescription: string;
  description: string;
  // ServiceNow uses 1 (high) .. 3 (low) for both impact and urgency.
  impact: "1" | "2" | "3";
  urgency: "1" | "2" | "3";
  category?: string;
  subcategory?: string;
  callerEmail?: string; // used to look up the ServiceNow caller record
  callerName?: string;
  // Free-form extra context appended into the work notes / description.
  contactType?: string; // e.g. "chat"
}

export interface ServiceNowResult {
  ok: boolean;
  incidentNumber?: string; // e.g. INC0012345
  sysId?: string;
  error?: string;
}

const resolveBaseUrl = (): string | undefined => {
  const raw = process.env.SERVICENOW_INSTANCE?.trim();
  if (!raw) return undefined;
  if (raw.startsWith("http://") || raw.startsWith("https://")) {
    return raw.replace(/\/+$/, "");
  }
  return `https://${raw}.service-now.com`;
};

export const isServiceNowConfigured = (): boolean => {
  return Boolean(
    resolveBaseUrl() &&
      process.env.SERVICENOW_USERNAME?.trim() &&
      process.env.SERVICENOW_PASSWORD?.trim()
  );
};

const getAuthHeader = (): string => {
  const user = process.env.SERVICENOW_USERNAME || "";
  const pass = process.env.SERVICENOW_PASSWORD || "";
  const encoded = Buffer.from(`${user}:${pass}`).toString("base64");
  return `Basic ${encoded}`;
};

const TIMEOUT_MS = 15000;

const fetchWithTimeout = async (
  url: string,
  init: RequestInit
): Promise<Response> => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    return await fetch(url, { ...init, signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
};

// Best-effort caller lookup by email. Returns a sys_id or undefined.
const findCallerSysId = async (
  baseUrl: string,
  email?: string
): Promise<string | undefined> => {
  if (!email) return undefined;
  try {
    const url = `${baseUrl}/api/now/table/sys_user?sysparm_query=email=${encodeURIComponent(
      email
    )}&sysparm_limit=1&sysparm_fields=sys_id`;

    const res = await fetchWithTimeout(url, {
      method: "GET",
      headers: {
        Authorization: getAuthHeader(),
        Accept: "application/json",
      },
    });

    if (!res.ok) return undefined;
    const body = (await res.json()) as { result?: Array<{ sys_id?: string }> };
    return body.result?.[0]?.sys_id;
  } catch {
    return undefined;
  }
};

export const createIncident = async (
  input: ServiceNowIncidentInput
): Promise<ServiceNowResult> => {
  const baseUrl = resolveBaseUrl();

  if (!baseUrl || !isServiceNowConfigured()) {
    return {
      ok: false,
      error:
        "ServiceNow is not configured. Set SERVICENOW_INSTANCE, SERVICENOW_USERNAME, and SERVICENOW_PASSWORD.",
    };
  }

  // Defense in depth: scrub anything credential-shaped before it lands in the
  // incident record, even though the caller should already have scrubbed.
  const shortDescription = redactForStorage(input.shortDescription).text.slice(0, 160);
  const description = redactForStorage(input.description).text;

  // Try to attach the real caller; fall back to a configured default.
  const callerSysId =
    (await findCallerSysId(baseUrl, input.callerEmail)) ||
    process.env.SERVICENOW_DEFAULT_CALLER_ID?.trim() ||
    undefined;

  const payload: Record<string, string> = {
    short_description: shortDescription,
    description,
    impact: input.impact,
    urgency: input.urgency,
    contact_type: input.contactType || "chat",
  };

  if (input.category) payload.category = input.category;
  if (input.subcategory) payload.subcategory = input.subcategory;
  if (callerSysId) payload.caller_id = callerSysId;
  if (process.env.SERVICENOW_ASSIGNMENT_GROUP?.trim()) {
    payload.assignment_group = process.env.SERVICENOW_ASSIGNMENT_GROUP.trim();
  }

  try {
    const res = await fetchWithTimeout(
      `${baseUrl}/api/now/table/incident`,
      {
        method: "POST",
        headers: {
          Authorization: getAuthHeader(),
          "Content-Type": "application/json",
          Accept: "application/json",
        },
        body: JSON.stringify(payload),
      }
    );

    if (!res.ok) {
      const detail = await res.text().catch(() => "");
      return {
        ok: false,
        error: `ServiceNow returned ${res.status} ${res.statusText}. ${detail.slice(
          0,
          300
        )}`,
      };
    }

    const body = (await res.json()) as {
      result?: { number?: string; sys_id?: string };
    };

    return {
      ok: true,
      incidentNumber: body.result?.number,
      sysId: body.result?.sys_id,
    };
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Unknown ServiceNow error";
    return { ok: false, error: message };
  }
};
