/*
 * VISITOR TELEMETRY EMAIL — fully self-contained daily 9 AM report.
 *
 * What it does (no Power Automate involved):
 *   1. TIMER  (05:00 UTC = 9:00 AM Dubai, every day)
 *   2. Refreshes the visitor data straight from the OneDrive VMS sources
 *      (src/shared/vmsRefresh.js) and commits a fresh visitor.xlsx to GitHub
 *      so the dashboard shows the same data as the email.
 *   3. Queries Dataverse DIRECTLY with client credentials — the two queries and every
 *      row condition are copied 1:1 from the dashboard's fetchTelemetryBookings()
 *      (Visitor-Competitor-Dashboard/index.html), so S & C counts match the dashboard.
 *   4. Builds the "Visitor Telemetry — Last 3 Days" table as Outlook-safe HTML.
 *   5. Sends it via Microsoft Graph FROM CK's mailbox to the team.
 *
 * HTTP endpoints for testing:
 *   GET /api/telemetry-email?code=...&refresh=1&dryrun=1  -> test the VMS refresh only (no commit, no email)
 *   GET /api/telemetry-email?code=...&refresh=1           -> refresh + commit visitor.xlsx, no email
 *   GET /api/telemetry-email?code=...&format=html         -> preview the email (no send)
 *   GET /api/telemetry-email?code=...&send=1              -> full run: refresh + commit + SEND
 *
 * Required Function App settings (portal -> Environment variables):
 *   TENANT_ID, CLIENT_ID, CLIENT_SECRET   (already set for the PR/PO proxy)
 *   DV_RESOURCE = https://operations-ifahr-live.crm15.dynamics.com
 *   MAIL_FROM   = Chandan.kumar@striveservicesgroup.com
 *   MAIL_TO     = semicolon-separated recipient list
 *   GH_TOKEN    = fine-grained PAT (contents RW on Visitor-Competitor-Dashboard) for the visitor.xlsx commit
 * One-time admin steps (see Visitor_Telemetry_Email_setup.md):
 *   - Application user for CLIENT_ID in the Power Platform environment (read role)
 *   - Graph application permissions Mail.Send + Files.Read.All, with admin consent
 *
 * Counting rules (identical to the dashboard):
 *   - S&C = Candoo-division bookings, Reactive work-order types only, cancelled
 *     excluded, "Team Leader/Member" resources excluded, laundry & fit-out
 *     departments excluded from bookings.
 *   - S&C Laundry = msdyn_workorders with department "Laundry*", not cancelled.
 *   - S&C Fit-out = OUR check-ins in visitor.xlsx (company Candoo/S&C/Strive, purpose fit-out).
 *   - De-dupe: date|unit|work-order-title|department = one visit.
 *   - Other = visitor.xlsx "Unit Visit" rows, excluding our companies and Dima.
 *   - "Not updated" = no visitor-log rows for that project/date.
 *   - Dates = the 3 most recent dates present in the visitor data (Dubai days).
 */
const { app } = require('@azure/functions');
const XLSX = require('xlsx');
const { refreshVms, commitVisitorXlsx } = require('../shared/vmsRefresh.js');

const VMS_URL = process.env.VMS_URL || 'https://strive-services-group.github.io/Visitor-Competitor-Dashboard/visitor.xlsx';
const DASHBOARD_URL = process.env.DASHBOARD_URL || 'https://strive-services-group.github.io/Visitor-Competitor-Dashboard/';
const DV_DIVISION_CANDOO = process.env.DV_DIVISION_CANDOO || '93bd43eb-0a44-ef11-a316-6045bd6a8335'; // Home Services - Candoo

const TELE_SVCS = ['LAUNDRY', 'HOUSEKEEP', 'MAINT', 'FITOUT'];
const TELE_LBL = { LAUNDRY: 'Laundry', HOUSEKEEP: 'Housekeeping', MAINT: 'Maintenance/ Handyman', FITOUT: 'Fitout' };
const TELE_PROJECTS = ['BALQIS RESIDENCE', 'THE8', 'SOUTH RESIDENCE', 'NORTH RESIDENCE', 'MAG 318', 'SHORELINE 7AND8'];
const TELE_PROJ_LBL = { 'BALQIS RESIDENCE': 'Balqis Residence', 'THE8': 'THE8', 'SOUTH RESIDENCE': 'South Residence', 'NORTH RESIDENCE': 'North Residence', 'MAG 318': 'MAG 318', 'SHORELINE 7AND8': 'Shoreline 7AND8' };
// Total units per project — Balqis fixed at 280 (user-confirmed universe); others from
// the dashboard's Locations master (distinct Building/Unit per project).
const PROJ_UNITS = { 'BALQIS RESIDENCE': 280, 'THE8': 195, 'SOUTH RESIDENCE': 281, 'NORTH RESIDENCE': 281, 'MAG 318': 439, 'SHORELINE 7AND8': 250 };

const F = '@OData.Community.Display.V1.FormattedValue';

/* ================= helpers copied from the dashboard ================= */
function canonService(s) {
  const u = (s || '').toUpperCase();
  if (u.includes('POOL') || u.includes('SWIM')) return 'POOL';
  if (u.includes('MAINT') || u.includes('HANDYMAN') || u.includes('AMC')) return 'MAINT';
  if (u.includes('CLEAN') || u.includes('HOUSE')) return 'HOUSEKEEP';
  if (u.includes('LAUNDR')) return 'LAUNDRY';
  if (u.includes('FIT') && u.includes('OUT')) return 'FITOUT';
  if (u.includes('INSPECT') || u.includes('SURVEY')) return 'INSPECT';
  return 'OTHER';
}
function dvIsOurCompany(name) {
  const n = (name || '').trim().toLowerCase();
  return n.indexOf('candoo') === 0 || n.indexOf('s & c') === 0 || n.indexOf('s&c') === 0 || n.indexOf('strive') === 0;
}
function teleIsOurLaundry(c) { return /dima/i.test(c || ''); }
function fmt(v) { return v || ''; }
function dubaiDate(iso) { return new Date(new Date(iso).getTime() + 4 * 3600 * 1000).toISOString().slice(0, 10); }
function teleLast7Dates() { // same 7-day Dubai window the dashboard fetches
  const out = [], now = new Date(Date.now() + 4 * 3600 * 1000);
  for (let i = 0; i < 7; i++) { const d = new Date(now); d.setUTCDate(d.getUTCDate() - i); out.push(d.toISOString().slice(0, 10)); }
  return out;
}
function teleFmtDate(d) {
  try { return new Date(d + 'T00:00:00Z').toLocaleDateString('en-GB', { weekday: 'short', day: '2-digit', month: 'short', timeZone: 'UTC' }).replace(',', ''); }
  catch (e) { return d; }
}
function teleHeat(p) { // Excel-style red->yellow->green for Share %
  if (p == null || isNaN(p)) return '';
  p = Math.max(0, Math.min(100, p));
  let a, b, t; const s = [[251, 226, 225], [254, 245, 208], [220, 240, 214]];
  if (p <= 50) { t = p / 50; a = s[0]; b = s[1]; } else { t = (p - 50) / 50; a = s[1]; b = s[2]; }
  return 'rgb(' + Math.round(a[0] + (b[0] - a[0]) * t) + ',' + Math.round(a[1] + (b[1] - a[1]) * t) + ',' + Math.round(a[2] + (b[2] - a[2]) * t) + ')';
}
function escHtml(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/* ================= OAuth (client credentials) ================= */
async function getToken(scopeBase) { // scopeBase e.g. https://operations-ifahr-live.crm15.dynamics.com or https://graph.microsoft.com
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    grant_type: 'client_credentials',
    scope: scopeBase.replace(/\/+$/, '') + '/.default'
  });
  const r = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });
  const j = await r.json();
  if (!r.ok || !j.access_token) throw new Error('token ' + r.status + ' ' + (j.error_description || j.error || ''));
  return j.access_token;
}

/* ================= Dataverse fetch — 1:1 with fetchTelemetryBookings() ================= */
async function dvGet(url, token) {
  const r = await fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'odata.include-annotations="*"' // <- gives the FormattedValue names the logic needs
    }
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error('odata ' + r.status + ' ' + JSON.stringify(j.error || j).slice(0, 300));
  return j;
}

async function fetchTelemetryRows(context) {
  const resource = (process.env.DV_RESOURCE || '').replace(/\/+$/, '');
  if (!resource) throw new Error('DV_RESOURCE app setting is not set');
  const token = await getToken(resource);
  const API = resource + '/api/data/v9.2';

  // same 7-day window + Dubai-day cut as the dashboard
  const dates = teleLast7Dates();
  const minD = dates[dates.length - 1];
  const cut = new Date(minD + 'T00:00:00Z'); cut.setUTCHours(cut.getUTCHours() - 4);

  const rows = [];

  // ---- Candoo bookings (expanded work order), same $select/$expand/$filter ----
  let url = API + '/bookableresourcebookings?$select=ssg_plannedstartdate,_resource_value,_bookingstatus_value'
    + '&$expand=msdyn_workorder($select=_ssg_project_value,_ssg_department_value,_ssg_building_value,_ssg_unit_value,ssg_title,_msdyn_workordertype_value)'
    + '&$filter=' + encodeURIComponent('msdyn_workorder/_ssg_division_value eq ' + DV_DIVISION_CANDOO + ' and ssg_plannedstartdate ge ' + cut.toISOString());
  let guard = 0;
  while (url && guard < 80) {
    const j = await dvGet(url, token);
    (j.value || []).forEach(r => {
      const w = r.msdyn_workorder || {};
      const stt = fmt(r['_bookingstatus_value' + F]).toLowerCase(); if (stt.indexOf('cancel') >= 0) return;
      const res = fmt(r['_resource_value' + F]); const resL = res.toLowerCase();
      if (resL.indexOf('team leader') >= 0 || resL.indexOf('team member') >= 0) return;
      const wtype = fmt(w['_msdyn_workordertype_value' + F]); if (!/reactive/i.test(wtype)) return; // only Reactive-*
      const dept0 = fmt(w['_ssg_department_value' + F]); if (/laundry/i.test(dept0)) return;        // laundry from work orders
      if (canonService(dept0) === 'FITOUT') return;                                                 // fit-out from visitor data
      if (!r.ssg_plannedstartdate) return;
      const d = dubaiDate(r.ssg_plannedstartdate);
      if (dates.indexOf(d) < 0) return;
      rows.push({
        date: d,
        project: fmt(w['_ssg_project_value' + F]).trim().toUpperCase(),
        service: canonService(dept0),
        unit: (fmt(w['_ssg_building_value' + F]).trim() + '_' + fmt(w['_ssg_unit_value' + F]).trim()),
        title: (w.ssg_title || ''),
        dept: dept0
      });
    });
    url = j['@odata.nextLink'] || null; guard++;
  }

  // ---- Our LAUNDRY straight from msdyn_workorders (dept "Laundry Services") ----
  let wurl = API + '/msdyn_workorders?$select=msdyn_name,ssg_title,createdon,msdyn_systemstatus,_ssg_building_value,_ssg_unit_value,_ssg_project_value,_ssg_department_value,_msdyn_workordertype_value'
    + '&$filter=' + encodeURIComponent('_ssg_division_value eq ' + DV_DIVISION_CANDOO + ' and createdon ge ' + cut.toISOString());
  let wguard = 0;
  while (wurl && wguard < 80) {
    const wj = await dvGet(wurl, token);
    (wj.value || []).forEach(wo => {
      const wdept = fmt(wo['_ssg_department_value' + F]); if (!/laundry/i.test(wdept)) return;
      const wst = fmt(wo['msdyn_systemstatus' + F]); if (/cancel/i.test(wst)) return;
      if (!wo.createdon) return;
      const wd = dubaiDate(wo.createdon);
      if (dates.indexOf(wd) < 0) return;
      rows.push({
        date: wd,
        project: fmt(wo['_ssg_project_value' + F]).trim().toUpperCase(),
        service: 'LAUNDRY',
        unit: (fmt(wo['_ssg_building_value' + F]).trim() + '_' + fmt(wo['_ssg_unit_value' + F]).trim()),
        title: (wo.ssg_title || ''),
        dept: wdept
      });
    });
    wurl = wj['@odata.nextLink'] || null; wguard++;
  }

  // ---- De-dupe: same unit + date + WO title + purpose = ONE visit ----
  const seen = {}, out = [];
  rows.forEach(rr => {
    const k = rr.date + '|' + rr.unit + '|' + String(rr.title || '').trim().toUpperCase() + '|' + String(rr.dept || rr.service || '').trim().toUpperCase();
    if (seen[k]) return; seen[k] = 1; out.push(rr);
  });
  if (context) context.log('telemetry rows after de-dupe:', out.length);
  return out;
}

/* ================= visitor.xlsx -> records (same parser as dashboard) ================= */
function parseVms(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  let hi = -1;
  for (let i = 0; i < Math.min(aoa.length, 10); i++) {
    if (aoa[i].some(c => String(c).trim().toLowerCase() === 'company name')) { hi = i; break; }
  }
  if (hi === -1) throw new Error('visitor.xlsx: header row not found');
  const headers = aoa[hi].map(h => String(h).trim());
  const idx = {
    date: headers.findIndex(h => /check.*in.*date/i.test(h)),
    type: headers.findIndex(h => /check.*in.*type/i.test(h)),
    purpose: headers.findIndex(h => /purpose/i.test(h)),
    company: headers.findIndex(h => /company/i.test(h)),
    unit: headers.findIndex(h => /unit/i.test(h)),
    project: headers.findIndex(h => /project/i.test(h))
  };
  const records = [];
  for (let i = hi + 1; i < aoa.length; i++) {
    const row = aoa[i];
    if (!row || !row[idx.company]) continue;
    if (idx.type >= 0 && String(row[idx.type] || '').trim().toLowerCase() !== 'unit visit') continue;
    if (/dima/i.test(String(row[idx.company] || ''))) continue; // Dima excluded everywhere
    const d = row[idx.date];
    let dateStr = '';
    if (typeof d === 'number') {
      const pd = XLSX.SSF.parse_date_code(d);
      if (pd) dateStr = `${pd.y}-${String(pd.m).padStart(2, '0')}-${String(pd.d).padStart(2, '0')}`;
    } else if (d instanceof Date) {
      dateStr = d.toISOString().slice(0, 10);
    } else if (typeof d === 'string' && d) {
      const pd = new Date(d);
      if (!isNaN(pd)) dateStr = pd.toISOString().slice(0, 10);
    }
    records.push({
      date: dateStr,
      purpose: String(row[idx.purpose] || '').trim(),
      company: String(row[idx.company] || '').trim(),
      unit: String(row[idx.unit] || '').trim(),
      project: idx.project >= 0 ? String(row[idx.project] || '').trim().toUpperCase() : ''
    });
  }
  return records;
}

/* ================= build the table model + email HTML ================= */
function buildEmail(vmsRecords, scRows) {
  // 3 most recent dates present in the visitor data, oldest -> newest
  const dset = {};
  vmsRecords.forEach(r => { if (r && r.date) dset[r.date] = 1; });
  let dates = Object.keys(dset).sort().reverse().slice(0, 3);
  if (!dates.length) {
    const t = new Date(Date.now() + 4 * 3600 * 1000);
    for (let i = 0; i < 3; i++) { const d = new Date(t); d.setUTCDate(d.getUTCDate() - i); dates.push(d.toISOString().slice(0, 10)); }
  }
  dates = dates.slice().reverse();
  const dateSet = {}; dates.forEach(d => dateSet[d] = 1);

  // competitor + our-fit-out + "data exists" maps from visitor data
  const comp = {}, vmsSeen = {}, ourfit = {};
  vmsRecords.forEach(r => {
    if (!r || !dateSet[r.date]) return;
    const p = (r.project || '').toUpperCase(); if (TELE_PROJECTS.indexOf(p) < 0) return;
    (vmsSeen[r.date] = vmsSeen[r.date] || {})[p] = true;
    if (dvIsOurCompany(r.company)) {
      if (canonService(r.purpose) === 'FITOUT') { (ourfit[r.date] = ourfit[r.date] || {}); ourfit[r.date][p] = (ourfit[r.date][p] || 0) + 1; }
      return;
    }
    if (teleIsOurLaundry(r.company)) return;
    const svc = canonService(r.purpose);
    (comp[r.date] = comp[r.date] || {}); (comp[r.date][p] = comp[r.date][p] || {});
    comp[r.date][p][svc] = (comp[r.date][p][svc] || 0) + 1;
  });

  // S&C map from Dataverse rows (already filtered + de-duped over the 7-day window)
  const urmMap = {};
  (scRows || []).forEach(r => {
    if (!dateSet[r.date]) return;
    (urmMap[r.date] = urmMap[r.date] || {});
    (urmMap[r.date][r.project] = urmMap[r.date][r.project] || {});
    urmMap[r.date][r.project][r.service] = (urmMap[r.date][r.project][r.service] || 0) + 1;
  });
  function urmFor(dt, p) {
    const base = (urmMap[dt] || {})[p] || {};
    const out = {}; for (const k in base) out[k] = base[k];
    out.FITOUT = (ourfit[dt] || {})[p] || 0; // our fit-out comes from visitor data
    return out;
  }

  /* ----- email-safe HTML — colour-grouped, FIXED-WIDTH layout (does not shrink) ----- */
  const FONT = 'Aptos,Segoe UI,Arial,sans-serif';
  const NAVYD = '#14315E';                                   // dark navy (labels + first 3 header cols)
  const GRP = {                                              // [group header colour, sub-header colour]
    ALL: ['#1D4ED8', '#1E40AF'],
    LAUNDRY: ['#6D28D9', '#4C1D95'],
    HOUSEKEEP: ['#16A34A', '#166534'],
    MAINT: ['#EA580C', '#9A3412'],
    FITOUT: ['#0891B2', '#0E7490']
  };
  const TICON = { ALL: '&#128101;', LAUNDRY: '&#129530;', HOUSEKEEP: '&#129529;', MAINT: '&#128736;&#65039;', FITOUT: '&#127959;&#65039;' }; // 👥 🧺 🧹 🛠️ 🏗️
  // fixed column widths (px) — table-layout:fixed keeps them constant on any screen
  const W_PROJ = 150, W_UNITS = 80, W_DATE = 105, W_SC = 52, W_OT = 122, W_PCT = 72;
  const TOTAL_W = W_PROJ + W_UNITS + W_DATE + 5 * (W_SC + W_OT + W_PCT);

  // no internal header borders — translucent borders render as white hairlines in Outlook at some zoom levels
  const thBase = 'color:#ffffff;font-family:' + FONT + ';font-weight:700;font-size:12.5px;padding:9px 8px;text-align:center;white-space:nowrap;border:0;';
  const thSub = 'color:#ffffff;font-family:' + FONT + ';font-weight:700;font-size:10.5px;padding:6px 4px;text-align:center;white-space:nowrap;border:0;';

  const groups = [['ALL', 'All Visitors']].concat(TELE_SVCS.map(s => [s, TELE_LBL[s]]));
  let head = '<tr>'
    + '<th rowspan="2" width="' + W_PROJ + '" style="' + thBase + 'background:' + NAVYD + ';">&#127970; Project</th>'
    + '<th rowspan="2" width="' + W_UNITS + '" style="' + thBase + 'background:' + NAVYD + ';">&#128230; Total Units</th>'
    + '<th rowspan="2" width="' + W_DATE + '" style="' + thBase + 'background:' + NAVYD + ';">&#128197; Date</th>';
  groups.forEach(g => {
    head += '<th colspan="3" width="' + (W_SC + W_OT + W_PCT) + '" style="' + thBase + 'background:' + GRP[g[0]][0] + ';">' + (TICON[g[0]] || '') + ' ' + escHtml(g[1]) + '</th>';
  });
  head += '</tr><tr>';
  groups.forEach(g => {
    const sub = GRP[g[0]][1];
    head += '<th width="' + W_SC + '" style="' + thSub + 'background:' + sub + ';">S&nbsp;&amp;&nbsp;C</th>'
      + '<th width="' + W_OT + '" style="' + thSub + 'background:' + sub + ';">Other</th>'
      + '<th width="' + W_PCT + '" style="' + thSub + 'background:' + sub + ';">Share&nbsp;%</th>';
  });
  head += '</tr>';

  // Share % pill (like the mock-up): red < 20 <= amber < 50 <= green
  function pill(pn) {
    if (pn == null) return '<span style="color:#9aa7b5;">-</span>';
    let bg = '#fde8e8', fg = '#c81e1e';
    if (pn >= 50) { bg = '#def7ec'; fg = '#03543f'; }
    else if (pn >= 20) { bg = '#fdf6b2'; fg = '#723b13'; }
    return '<span style="display:inline-block;padding:4px 12px;border-radius:5px;background:' + bg + ';color:' + fg + ';font-weight:700;font-size:12px;">' + pn + '%</span>';
  }

  let body = '';
  TELE_PROJECTS.forEach((p, pi) => {
    const apts = PROJ_UNITS[p] || 0;
    const blockBg = pi % 2 ? '#f7f9fc' : '#ffffff'; // alternate project blocks
    const tdD = 'padding:8px 6px;border-bottom:1px solid #e5e9f0;border-right:1px solid #edf1f6;font-family:' + FONT + ';font-size:12px;text-align:center;white-space:nowrap;vertical-align:middle;background:' + blockBg + ';';
    dates.forEach((dt, di) => {
      const u = urmFor(dt, p), o = (comp[dt] || {})[p] || {};
      let uAll = 0, oAll = 0, kk; for (kk in u) uAll += u[kk]; for (kk in o) oAll += o[kk];
      const noVms = !(vmsSeen[dt] && vmsSeen[dt][p]);
      const grpTop = di === 0 ? 'border-top:2px solid #d6dee8;' : '';
      body += '<tr>';
      if (di === 0) {
        body += '<td rowspan="' + dates.length + '" style="' + tdD + grpTop + 'font-weight:800;color:' + NAVYD + ';font-size:13.5px;text-align:left;padding-left:10px;">&#127970; ' + escHtml(TELE_PROJ_LBL[p] || p) + '</td>';
        body += '<td rowspan="' + dates.length + '" style="' + tdD + grpTop + 'font-weight:800;color:#1D4ED8;font-size:15px;">' + apts.toLocaleString() + '</td>';
      }
      body += '<td style="' + tdD + grpTop + 'font-weight:700;color:#2563EB;">' + escHtml(teleFmtDate(dt)) + ' &rsaquo;</td>';
      function cell(uv, ov) {
        const ud = uv ? '<span style="color:#1f2937;font-weight:700;">' + uv + '</span>' : '<span style="color:#9aa7b5;">0</span>';
        if (noVms) {
          return '<td style="' + tdD + grpTop + '">' + ud + '</td>' +
            '<td style="' + tdD + grpTop + 'font-size:10px;"><span style="color:#c2410c;">&#9888;&#65039;</span> <i style="color:#92703f;font-weight:600;">Not updated</i></td>' +
            '<td style="' + tdD + grpTop + 'color:#9aa7b5;">-</td>';
        }
        const has = (uv + ov) > 0;
        const pn = has ? Math.round(uv / (uv + ov) * 100) : null;
        const od = ov ? '<span style="color:#1f2937;font-weight:700;">' + ov + '</span>' : '<span style="color:#9aa7b5;">0</span>';
        return '<td style="' + tdD + grpTop + '">' + ud + '</td>'
          + '<td style="' + tdD + grpTop + '">' + od + '</td>'
          + '<td style="' + tdD + grpTop + '">' + pill(pn) + '</td>';
      }
      body += cell(uAll, oAll);
      TELE_SVCS.forEach(s => { body += cell(u[s] || 0, o[s] || 0); });
      body += '</tr>';
    });
  });

  const latest = dates[dates.length - 1];
  const nowDubai = new Date(Date.now() + 4 * 3600 * 1000);
  const stamp = nowDubai.toISOString().slice(0, 10) + ' ' + nowDubai.toISOString().slice(11, 16);
  const subject = 'Visitor Telemetry (Flagship Projects) - Last 3 Days (' + teleFmtDate(dates[0]) + ' – ' + teleFmtDate(latest) + ')';

  // NOTE: body text uses HTML entities only (&#183; = the dot separator) — literal UTF-8
  // characters like — and · can render as "â€"" in Outlook depending on charset detection.
  const html =
    '<div style="font-family:' + FONT + ';color:#22303c;">' +
    '<div style="font-family:' + FONT + ';font-weight:700;font-size:16px;color:#145A95;border-left:4px solid #618FB4;padding-left:10px;margin:0 0 4px;">VISITOR TELEMETRY (FLAGSHIP PROJECTS) - LAST 3 DAYS</div>' +
    '<div style="font-family:' + FONT + ';font-size:13px;color:#607083;margin:0 0 10px;">S &amp; C = our visits &#183; Other = competitor &#183; snapshot taken ' + stamp + ' (Dubai) &#183; <a href="' + DASHBOARD_URL + '#visitor" style="color:#145A95;font-weight:700;text-decoration:none;">Open the Live Dashboard</a> for more details and drill-through</div>' +
    '<div style="border:1px solid #dbe3ec;border-radius:12px;overflow:hidden;display:inline-block;">' +
    '<table cellpadding="0" cellspacing="0" border="0" width="' + TOTAL_W + '" style="border-collapse:collapse;table-layout:fixed;width:' + TOTAL_W + 'px;background:#ffffff;">' + head + body + '</table>' +
    '</div>' +
    '<div style="font-family:' + FONT + ';font-size:10px;color:#8b98a5;margin-top:8px;">Automated daily 9:00 AM report &#183; Strive Services Group &#183; data: Candoo bookings &amp; work orders (Dynamics 365) + building visitor logs</div>' +
    '</div>';

  return { subject, html, dates };
}

/* ================= send from CK's mailbox via Microsoft Graph ================= */
async function sendMail(subject, html, context) {
  const from = process.env.MAIL_FROM;
  const toList = (process.env.MAIL_TO || '').split(/[;,]/).map(s => s.trim()).filter(Boolean);
  if (!from) throw new Error('MAIL_FROM app setting is not set');
  if (!toList.length) throw new Error('MAIL_TO app setting is not set');
  const token = await getToken('https://graph.microsoft.com');
  const r = await fetch('https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(from) + '/sendMail', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: 'HTML', content: html },
        toRecipients: toList.map(a => ({ emailAddress: { address: a } }))
      },
      saveToSentItems: true
    })
  });
  if (r.status !== 202) {
    const j = await r.json().catch(() => ({}));
    throw new Error('sendMail ' + r.status + ' ' + JSON.stringify(j.error || j).slice(0, 300));
  }
  if (context) context.log('email sent to', toList.length, 'recipients');
}

/* ================= the full job ================= */
// 1) refresh visitor data straight from the OneDrive sources (headless, no laptop),
//    and commit the fresh visitor.xlsx to GitHub so the dashboard matches the email;
// 2) fall back to the last published visitor.xlsx if the refresh fails;
// 3) query Dataverse, build the table, send.
async function runReport(context, doSend) {
  let vms = null, refreshInfo = null;
  try {
    const rf = await refreshVms(getToken, VMS_URL, context);
    vms = rf.records;
    refreshInfo = { total: rf.total, counts: rf.counts, missing: rf.missing, committed: false };
    try { refreshInfo.committed = await commitVisitorXlsx(rf.xlsxBuffer, context); }
    catch (e) { if (context) context.warn('GitHub commit failed (email still uses fresh data): ' + e.message); }
  } catch (e) {
    if (context) context.warn('VMS refresh failed, falling back to last published visitor.xlsx: ' + e.message);
  }
  if (!vms) {
    const vr = await fetch(VMS_URL + '?t=' + Date.now());
    if (!vr.ok) throw new Error('visitor.xlsx fetch failed: ' + vr.status);
    vms = parseVms(Buffer.from(await vr.arrayBuffer()));
  }
  const scRows = await fetchTelemetryRows(context);
  const out = buildEmail(vms, scRows);
  out.refresh = refreshInfo || { fallback: 'last published visitor.xlsx' };
  if (doSend) await sendMail(out.subject, out.html, context);
  return out;
}

/* ================= triggers ================= */
// Daily at 05:00 UTC = 9:00 AM Dubai
app.timer('telemetry-email-daily', {
  schedule: '0 0 5 * * *',
  handler: async (timer, context) => {
    try { await runReport(context, true); }
    catch (e) { context.error('telemetry email FAILED:', e.message); throw e; }
  }
});

// Manual preview / manual send (for testing):
//   ?format=html          -> shows the email in the browser, does NOT send
//   ?send=1               -> builds and SENDS now (also refreshes + commits visitor.xlsx)
//   ?refresh=1&dryrun=1   -> ONLY test the OneDrive VMS refresh: returns per-project
//                            row counts, commits nothing, sends nothing
//   ?refresh=1            -> refresh + commit visitor.xlsx to GitHub, no email
app.http('telemetry-email', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'function',
  route: 'telemetry-email',
  handler: async (request, context) => {
    try {
      const url = new URL(request.url);
      if (url.searchParams.get('refresh') === '1') {
        const rf = await refreshVms(getToken, VMS_URL, context);
        let committed = false;
        if (url.searchParams.get('dryrun') !== '1') committed = await commitVisitorXlsx(rf.xlsxBuffer, context);
        return { status: 200, jsonBody: { total: rf.total, counts: rf.counts, missing: rf.missing, committed } };
      }
      const doSend = url.searchParams.get('send') === '1';
      const out = await runReport(context, doSend);
      if (url.searchParams.get('format') === 'html' && !doSend) {
        return { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8' }, body: out.html };
      }
      return { status: 200, jsonBody: { sent: doSend, subject: out.subject, dates: out.dates, refresh: out.refresh } };
    } catch (e) {
      context.error('telemetry-email failed:', e);
      return { status: 500, jsonBody: { error: e.message } };
    }
  }
});

module.exports = { buildEmail, parseVms, canonService };
