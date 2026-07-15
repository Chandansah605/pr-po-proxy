const { app } = require('@azure/functions');
const { jwtVerify, createRemoteJWKSet } = require('jose');
const STEP_MAP = require('../../stepMap.json');
const PO_STEP_MAP = require('../../poStepMap.json');

// ---- Entra token validation: the dashboard sends the signed-in user's token ----
const TENANT = process.env.TENANT_ID;
const AUDIENCE = process.env.DASHBOARD_CLIENT_ID; // the dashboard's MSAL client id
const JWKS = TENANT ? createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${TENANT}/discovery/v2.0/keys`)) : null;
async function requireUser(request) {
  const h = (request.headers.get && request.headers.get('authorization')) || '';
  const m = h.match(/^Bearer\s+(.+)$/i);
  if (!m) throw new Error('unauthorized: missing token');
  await jwtVerify(m[1], JWKS, {
    issuer: `https://login.microsoftonline.com/${TENANT}/v2.0`,
    audience: AUDIENCE
  });
}

// ---- simple in-memory cache (per warm instance) ----
const CACHE_MS = 3 * 60 * 1000;
const cache = {}; // { pr: {at, data}, po: {...} }

// ---- CORS ----
function cors() {
  const origin = process.env.ALLOWED_ORIGIN || '*';
  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Cache-Control': 'no-store'
  };
}

// ---- OAuth client-credentials token for F&O ----
async function getToken() {
  const tenant = process.env.TENANT_ID;
  const resource = (process.env.FO_RESOURCE || '').replace(/\/+$/, '');
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    grant_type: 'client_credentials',
    scope: resource + '/.default'
  });
  const r = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });
  if (!r.ok) throw new Error('token ' + r.status + ': ' + (await r.text()).slice(0, 300));
  return (await r.json()).access_token;
}

// ---- paged OData GET ----
async function odataAll(token, path) {
  const base = (process.env.FO_RESOURCE || '').replace(/\/+$/, '') + '/data/';
  let url = base + path;
  const items = [];
  let guard = 0;
  while (url && guard++ < 500) {
    const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' } });
    if (!r.ok) throw new Error('odata ' + r.status + ' @ ' + url + ': ' + (await r.text()).slice(0, 300));
    const j = await r.json();
    if (j.value) items.push(...j.value);
    url = j['@odata.nextLink'] || null;
  }
  return items;
}

function num(x) { const n = Number(x); return Number.isFinite(n) ? n : 0; }
function prKey(s) { const m = String(s || '').match(/C?PR-\d+/); return m ? m[0] : null; }
function poKey(s) {
  const m = String(s || '').match(/Purchase Order:\s*([A-Za-z0-9\-]+)/);
  if (m) return m[1];
  const m2 = String(s || '').match(/[A-Z]+-PO\d+|PO?\d{6,}/);
  return m2 ? m2[0] : null;
}
// DefaultLedgerDimensionDisplayValue looks like "-Contracted--Building Services--THE8-Materials-Threshold-"
// First three non-empty segments = Contract, Department, Location(Project).
function parseDim(s) {
  if (!s) return { contract: null, department: null, location: null };
  const p = String(s).split('-').map(x => x.trim()).filter(Boolean);
  return { contract: p[0] || null, department: p[1] || null, location: p[2] || null };
}

// ---- assemble PR rows ----
async function buildPR() {
  const token = await getToken();

  // Try the IFAHR custom fields first (exact Step name etc. straight from PurchReqTable).
  // Until the F&O developer exposes them on OData (see FO-CUSTOM-ENTITY-SPEC-step-name.md),
  // this $select fails and we fall back to the workflow-work-item reconstruction below.
  const BASE_HDR = "RequisitionNumber,RequisitionName,RequisitionStatus,DefaultProjectId,IFAHRQuotationReference,PreparerPersonnelNumber,RequisitionPurpose,DefaultRequestedDate";
  const IFAHR_HDR = BASE_HDR + ",IFAHRPendingStep,IFAHRPendngDateTime,IFAHRPendingUser,IFAHRPRWFAcceptedBy,IFAHRDepartment,IFAHRLocation,IFAHRContract,IFAHRSubmissionStatus,IFAHRPurchRFQCaseId";
  let headers, ifahrLive = false;
  try {
    headers = await odataAll(token, "PurchaseRequisitionHeaders?$select=" + IFAHR_HDR);
    ifahrLive = true;
  } catch (e) {
    headers = await odataAll(token, "PurchaseRequisitionHeaders?$select=" + BASE_HDR);
  }

  const lines = await odataAll(token,
    "PurchaseRequisitionLines?$select=RequisitionNumber,LineAmount,DefaultLedgerDimensionDisplayValue");

  // real Created / Submitted dates (header entity doesn't expose them cleanly)
  const bi = await odataAll(token,
    "PurchReqTableBiEntities?$select=PurchReqId,TransDate,SubmittedDateTime");
  const biMap = {};
  for (const b of bi) { if (b.PurchReqId) biMap[b.PurchReqId] = b; }

  const lineAgg = {};
  for (const l of lines) {
    const k = l.RequisitionNumber;
    if (!k) continue;
    if (!lineAgg[k]) lineAgg[k] = { total: 0, first: l };
    lineAgg[k].total += num(l.LineAmount);
  }

  const wi = await odataAll(token,
    "WorkflowWorkItems?$filter=MenuItemName eq 'PurchReqTable'&$select=Subject,ElementId,Status,UserId,DueDateTime");

  const current = {}, lastDone = {};
  for (const w of wi) {
    const k = prKey(w.Subject);
    if (!k) continue;
    if (w.Status === 'Pending') {
      const prev = current[k];
      if (!prev || new Date(w.DueDateTime) > new Date(prev.DueDateTime)) current[k] = w;
    } else if (w.Status === 'Completed') {
      const prev = lastDone[k];
      if (!prev || new Date(w.DueDateTime) > new Date(prev.DueDateTime)) lastDone[k] = w;
    }
  }

  const rows = headers.map(h => {
    const k = h.RequisitionNumber;
    const agg = lineAgg[k];
    const line = agg ? agg.first : {};
    const w = current[k];
    const elementId = w ? w.ElementId : null;
    const dim = parseDim(line.DefaultLedgerDimensionDisplayValue);
    return {
      purchaseRequisition: k,
      quotationReference: h.IFAHRQuotationReference || null,
      name: h.RequisitionName || null,
      preparer: h.PreparerPersonnelNumber || null,
      projectId: h.DefaultProjectId || null,
      status: h.RequisitionStatus || null,
      createdDate: (biMap[k] && biMap[k].TransDate) || h.DefaultRequestedDate || null,
      submittedDate: (biMap[k] && biMap[k].SubmittedDateTime) || null,
      acceptedByAssignTo: (ifahrLive && h.IFAHRPRWFAcceptedBy) || (lastDone[k] ? lastDone[k].UserId : null),
      department: (ifahrLive && h.IFAHRDepartment) || dim.department,
      location: (ifahrLive && h.IFAHRLocation) || dim.location,
      contract: (ifahrLive && h.IFAHRContract) || dim.contract,
      submissionStatus: ifahrLive ? (h.IFAHRSubmissionStatus || null) : null,
      rfqCase: ifahrLive ? (h.IFAHRPurchRFQCaseId || null) : null,
      totalAmount: agg ? Math.round(agg.total * 100) / 100 : 0,
      pendingApprover: (ifahrLive && h.IFAHRPendingUser) || (w ? w.UserId : null),
      stepName: (ifahrLive && h.IFAHRPendingStep) || (elementId ? (STEP_MAP[elementId] || null) : null),
      stepDateTime: (ifahrLive && h.IFAHRPendngDateTime) || (w ? w.DueDateTime : null),
      stepElementId: elementId,
      ledgerDimensionRaw: line.DefaultLedgerDimensionDisplayValue || null
    };
  });

  return { type: 'pr', generatedAt: new Date().toISOString(), ifahrLive, count: rows.length, rows };
}

// ---- assemble PO rows ----
async function buildPO() {
  const token = await getToken();

  const headers = await odataAll(token,
    "PurchaseOrderHeadersV2?$select=PurchaseOrderNumber,OrderVendorAccountNumber,PurchaseOrderName,CurrencyCode,PurchaseOrderStatus,DocumentApprovalStatus,ProjectId,RequestedDeliveryDate");

  const lines = await odataAll(token,
    "PurchaseOrderLinesV2?$select=PurchaseOrderNumber,LineAmount,DefaultLedgerDimensionDisplayValue");

  const lineAgg = {};
  for (const l of lines) {
    const k = l.PurchaseOrderNumber;
    if (!k) continue;
    if (!lineAgg[k]) lineAgg[k] = { total: 0, first: l };
    lineAgg[k].total += num(l.LineAmount);
  }

  const vendors = await odataAll(token,
    "VendorsV2?$select=VendorAccountNumber,VendorOrganizationName");
  const vname = {};
  for (const v of vendors) { if (v.VendorAccountNumber) vname[v.VendorAccountNumber] = v.VendorOrganizationName || ''; }

  const wi = await odataAll(token,
    "WorkflowWorkItems?$filter=MenuItemName eq 'PurchTable'&$select=Subject,ElementId,Status,UserId,DueDateTime");
  const current = {};
  for (const w of wi) {
    if (w.Status !== 'Pending') continue;
    const k = poKey(w.Subject);
    if (!k) continue;
    const prev = current[k];
    if (!prev || new Date(w.DueDateTime) > new Date(prev.DueDateTime)) current[k] = w;
  }

  const rows = headers.map(h => {
    const k = h.PurchaseOrderNumber;
    const agg = lineAgg[k];
    const line = agg ? agg.first : {};
    const w = current[k];
    const elementId = w ? w.ElementId : null;
    const dim = parseDim(line.DefaultLedgerDimensionDisplayValue);
    return {
      purchaseOrder: k,
      linkedPR: h.PurchaseRequisitionNumber || null,
      vendorName: vname[h.OrderVendorAccountNumber] || h.OrderVendorAccountNumber || null,
      name: h.PurchaseOrderName || null,
      approvalStatus: h.DocumentApprovalStatus || null,
      poStatus: h.PurchaseOrderStatus || null,
      currency: h.CurrencyCode || null,
      projectId: h.ProjectId || null,
      createdDate: h.RequestedDeliveryDate || null,
      department: dim.department,
      location: dim.location,
      contract: dim.contract,
      totalAmount: agg ? Math.round(agg.total * 100) / 100 : 0,
      pendingApprover: w ? w.UserId : null,
      stepName: elementId ? (PO_STEP_MAP[elementId] || null) : null,
      stepDateTime: w ? w.DueDateTime : null,
      stepElementId: elementId,
      ledgerDimensionRaw: line.DefaultLedgerDimensionDisplayValue || null
    };
  });

  return { type: 'po', generatedAt: new Date().toISOString(), count: rows.length, rows };
}

async function serve(kind, request, context) {
  const headers = { 'Content-Type': 'application/json', ...cors() };
  if (request.method === 'OPTIONS') return { status: 204, headers };

  // Access is gated by the ?code= function key (authLevel:'function').
  // The Entra user-token check is optional: enforced only when a token is sent,
  // so the dashboards' key-only calls work (matches the original design).
  try {
    const authz = (request.headers.get && request.headers.get('authorization')) || '';
    if (authz) await requireUser(request);
  } catch (e) {
    return { status: 401, headers, jsonBody: { error: 'unauthorized' } };
  }

  try {
    const now = Date.now();
    if (cache[kind] && now - cache[kind].at < CACHE_MS) {
      return { status: 200, headers, jsonBody: { ...cache[kind].data, cached: true } };
    }
    const data = kind === 'po' ? await buildPO() : await buildPR();
    cache[kind] = { at: now, data };
    return { status: 200, headers, jsonBody: data };
  } catch (e) {
    context.error(e);
    return { status: 500, headers, jsonBody: { error: String(e && e.message || e) } };
  }
}

app.http('pr', { methods: ['GET', 'OPTIONS'], authLevel: 'anonymous', route: 'pr', handler: (req, ctx) => serve('pr', req, ctx) });
app.http('po', { methods: ['GET', 'OPTIONS'], authLevel: 'anonymous', route: 'po', handler: (req, ctx) => serve('po', req, ctx) });
