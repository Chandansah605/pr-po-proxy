const { app } = require('@azure/functions');
const STEP_MAP = require('../../stepMap.json');

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

  // 1) headers (paged in full)
  const headers = await odataAll(token,
    "PurchaseRequisitionHeaders?$select=RequisitionNumber,RequisitionName,RequisitionStatus,DefaultProjectId,IFAHRQuotationReference,PreparerPersonnelNumber,RequisitionPurpose,DefaultRequestedDate");

  // 2) lines (for total amount + financial dimensions)
  const lines = await odataAll(token,
    "PurchaseRequisitionLines?$select=RequisitionNumber,LineAmount,DefaultLedgerDimensionDisplayValue");

  const lineAgg = {}; // RequisitionNumber -> {total, first}
  for (const l of lines) {
    const k = l.RequisitionNumber;
    if (!k) continue;
    if (!lineAgg[k]) lineAgg[k] = { total: 0, first: l };
    lineAgg[k].total += num(l.LineAmount);
  }

  // 3) workflow work items for purchase requisitions (pending = current step)
  const wi = await odataAll(token,
    "WorkflowWorkItems?$filter=MenuItemName eq 'PurchReqTable'&$select=Subject,ElementId,Status,UserId,DueDateTime");

  const current = {}; // PR -> latest pending work item
  for (const w of wi) {
    if (w.Status !== 'Pending') continue;
    const k = prKey(w.Subject);
    if (!k) continue;
    const prev = current[k];
    if (!prev || new Date(w.DueDateTime) > new Date(prev.DueDateTime)) current[k] = w;
  }

  // 4) merge
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
      preparer: h.PreparerPersonnelNumber || null,        // TODO: resolve personnel number -> worker name
      projectId: h.DefaultProjectId || null,
      status: h.RequisitionStatus || null,
      createdDate: h.DefaultRequestedDate || null,         // TODO: confirm true Created/Submitted dates
      submittedDate: null,                                 // TODO
      acceptedByAssignTo: null,                            // TODO
      department: dim.department,
      location: dim.location,
      contract: dim.contract,
      totalAmount: agg ? Math.round(agg.total * 100) / 100 : 0,
      pendingApprover: w ? w.UserId : null,
      stepName: elementId ? (STEP_MAP[elementId] || null) : null,
      stepDateTime: w ? w.DueDateTime : null,
      stepElementId: elementId,                            // used to complete stepMap.json
      ledgerDimensionRaw: line.DefaultLedgerDimensionDisplayValue || null
    };
  });

  return { type: 'pr', generatedAt: new Date().toISOString(), count: rows.length, rows };
}

async function serve(kind, request, context) {
  const headers = { 'Content-Type': 'application/json', ...cors() };
  if (request.method === 'OPTIONS') return { status: 204, headers };

  try {
    const now = Date.now();
    if (cache[kind] && now - cache[kind].at < CACHE_MS) {
      return { status: 200, headers, jsonBody: { ...cache[kind].data, cached: true } };
    }
    if (kind === 'po') {
      return { status: 501, headers, jsonBody: { error: 'PO not implemented yet — PR validated first.' } };
    }
    const data = await buildPR();
    cache[kind] = { at: now, data };
    return { status: 200, headers, jsonBody: data };
  } catch (e) {
    context.error(e);
    return { status: 500, headers, jsonBody: { error: String(e && e.message || e) } };
  }
}

app.http('pr', { methods: ['GET', 'OPTIONS'], authLevel: 'function', route: 'pr', handler: (req, ctx) => serve('pr', req, ctx) });
app.http('po', { methods: ['GET', 'OPTIONS'], authLevel: 'function', route: 'po', handler: (req, ctx) => serve('po', req, ctx) });
