// ============================================================
// TREATS WEB APP v12 — read sales (MongoDB API) + inventory log + production plan + write all of them
// Deploy as Web App: Execute as "Me", Access "Anyone"
// AFTER editing this file, you MUST re-deploy (Deploy → Manage deployments → Edit → New version)
//
// CHANGES vs v11:
//   - getSalesData() now sources sales from the durable MongoDB-backed API
//     (GET <SALES_API_BASE>/api/treats-sales-saturday, Bearer token stored in
//     Script Properties key TREATS_SALES_API_TOKEN) instead of the rolling daily
//     GID tabs. Fixes inventory drift: the old per-day tabs aged out after ~35 days,
//     so once a stock anchor (last stocktake, or the 9 May CLOSING_DATE baseline)
//     passed that window the in-between sales were lost and levels overstated.
//     Mongo keeps all history; we pull every row since SALES_SINCE_DATE so the
//     dashboard can deduct sales back to each SKU's anchor.
//   - TabIndex / per-day tabs AND the TodaySales intraday tab are no longer read —
//     sales come solely from Mongo. The sync runs a few times daily, so today's sales
//     appear within hours (intraday lag accepted; getSalesData touches no sheet now).
//   - Returned row shape is unchanged ([date, sku, uom, qty, seq, source]) so the
//     dashboard needs no change.
//
// CHANGES vs v9:
//   - ProductionPlan schema extended with new column: I fulfilledByUuid
//     This records which Production log entry (by UUID) fulfilled this plan slot.
//     Lets the calendar correctly attribute night-shift productions logged the
//     next morning (e.g. Mon-night plan + Tue-morning log → silent match).
//   - setPlanEntry now accepts an optional fulfilledByUuid field (empty string
//     clears it).
//   - No breaking changes — old rows without column I just read as empty/null.
//
// CHANGES vs v8:
//   - Added ProductionPlan tab support — persistent storage for planned production
//     slots (so plan-vs-actual comparison survives reloads).
//   - getCombinedJSON() now also returns `productionPlan` (current + future weeks only)
//   - doPost() now routes `action`-based requests in addition to the legacy
//     inventory-log append. Actions: 'logInventory' (default), 'setPlanEntry',
//     'deletePlanEntry'. Existing dashboard calls continue to work unchanged.
//   - Tab schema (ProductionPlan): A entryId | B weekStartDate (DD/MM/YYYY) |
//     C dayDate (DD/MM/YYYY) | D shift ('morning'|'night'|'thaw') | E sku |
//     F rawQty (number) | G note | H createdAt (ISO) | I fulfilledByUuid (v10)
//
// CHANGES vs v7:
//   - Fixed dedup bug between archived daily tabs and TodaySales that caused
//     sales to double-count when date cells were stored in inconsistent formats.
//
// CHANGES vs v6:
//   - getSalesData() now also reads from the 'TodaySales' tab (gid 1446806035)
//   - TodaySales is populated at 12pm SGT and cleared at 12am SGT — provides intraday sales
//   - Dedup safeguard: if a date appears in BOTH an archived daily tab AND TodaySales,
//     the archived tab wins (TodaySales rows for that date are skipped). This protects
//     against the 1am-archive-runs-while-TodaySales-still-has-data race.
//
// CHANGES vs v5:
//   - Added 'Stocktake' event type — records absolute physical count (signed; allows negatives)
//   - Stock Discrepancy is now the companion variance event (kept separate for audit trail)
//   - Dashboard writes BOTH rows on stocktake: one Stocktake (absolute), one Stock Discrepancy (diff)
//
// CHANGES vs v4:
//   - doGet() returns both sales data AND InventoryLog rows in one response
//     under keys `salesRows` and `inventoryLog`
//   - Helper getInventoryLogData() reads all rows from the InventoryLog tab
// ============================================================

const SPREADSHEETID = '1wvqLGnHMIZXlRbsCNdQzlQfDcbzXFtLcDyttQADXu88';
const INVENTORY_LOG_TAB = 'InventoryLog';  // must match the tab name in the sheet
const PRODUCTION_PLAN_TAB = 'ProductionPlan';   // v9: persistent planned production slots

// ── v12: sales source = durable MongoDB API (read-only) ──────────────────────
const SALES_API_BASE = 'https://sg.omakase.pet';
const SALES_API_PATH = '/api/treats-sales-saturday';
const SALES_API_PAGE = 1000;          // API hard max records per page
// Fetch floor (inclusive), YYYY-MM-DD. MUST be <= the oldest active anchor across
// ALL SKUs — i.e. the dashboard's 9 May 2026 CLOSING_DATE baseline, or any SKU's
// most recent stocktake, whichever is earliest. The dashboard deducts sales since
// each SKU's anchor to derive current stock; if this floor is LATER than an
// un-stocktaken SKU's anchor, that SKU's level silently overstates again. Only
// advance this after a FULL all-SKU stocktake re-anchors everything.
const SALES_SINCE_DATE = '2026-02-02';
// Bearer token lives in Script Properties (Project Settings → Script properties),
// NEVER in the client HTML. Key: TREATS_SALES_API_TOKEN.
function getSalesApiToken_() {
  const t = PropertiesService.getScriptProperties().getProperty('TREATS_SALES_API_TOKEN');
  if (!t) throw new Error('TREATS_SALES_API_TOKEN not set in Script Properties');
  return t;
}

// Shared secret — must match SYNC_SECRET in the dashboard HTML.
// Change this string to something only you and the dashboard know.
// This blocks random POSTs to your endpoint from people who somehow find the URL.
const SYNC_SECRET = 'treats-ops-2026-shared-secret';

// ============================================================
// GET — return sales data + inventory log in one call
// ============================================================
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify(getCombinedJSON(e && e.parameter)))
    .setMimeType(ContentService.MimeType.JSON);
}

function getCombinedJSON(params) {
  params = params || {};
  try {
    const sales = getSalesData(params.date_from, params.date_to);
    const inv = getInventoryLogData();
    const plan = getProductionPlanData();

    return {
      success: true,
      lastUpdated: new Date().toISOString(),
      // sales side
      rowCount: sales.rowCount,
      tabsProcessed: sales.tabsProcessed,
      salesRows: sales.rows,
      rows: sales.rows,             // backwards-compat alias for dashboards on v4
      dedupReport: sales.dedupReport,   // v8: diagnostic for date dedup
      // inventory log side
      inventoryLogCount: inv.length,
      inventoryLog: inv,
      // v9: production plan side
      productionPlanCount: plan.length,
      productionPlan: plan,
      test:"test"
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// --- Sales reader (v12.1: MongoDB API only) ---
// All sales come from the durable MongoDB-backed API. The old per-day GID tabs AND
// the TodaySales intraday tab are no longer read — Mongo is the single source, so no
// cross-source dedup is needed. We pull every row since SALES_SINCE_DATE (paginated)
// so the dashboard can deduct sales back to each SKU's anchor (fixes level drift),
// while the burn-rate MA still uses only the last 4 weeks downstream.
//
// Dates flow through parseAnyDateCell → normalized DD/MM/YYYY (Asia/Singapore), so the
// dashboard reads them identically whether the API stores ISO or DD/MM/YYYY.
// Convert dashboard dd-mm-yyyy (or dd/mm/yyyy) to YYYY-MM-DD for the sales API.
function dashDMYToISO_(s) {
  if (!s) return null;
  const p = String(s).trim().split(/[-\/]/);
  if (p.length !== 3) return null;
  const dd = p[0].padStart(2, '0');
  const mm = p[1].padStart(2, '0');
  const yyyy = p[2];
  if (!/^\d{4}$/.test(yyyy)) return null;
  return yyyy + '-' + mm + '-' + dd;
}

function getSalesData(dateFrom, dateTo) {
  // Sales come entirely from the durable MongoDB API now. The TodaySales intraday tab
  // has been dropped — the sync runs a few times daily, so today's sales land within
  // hours (intraday lag accepted). This function no longer touches any sheet; the only
  // remaining sheet reads are in getInventoryLogData / getProductionPlanData.
  const rows = [];
  let rowSeq = 0;
  const token = getSalesApiToken_();
  const dateGte = dashDMYToISO_(dateFrom) || SALES_SINCE_DATE;
  const dateLte = dashDMYToISO_(dateTo);

  // Paginated pull of every row since SALES_SINCE_DATE so the dashboard can deduct
  // sales back to each SKU's anchor — this is what fixes the post-35-day level drift.
  // The burn-rate MA still slices only the last 4 weeks downstream, so the extra
  // history is free for the rate calc.
  let skip = 0, total = null, pages = 0, fetched = 0;
  do {
    let url = SALES_API_BASE + SALES_API_PATH +
      '?date_gte=' + encodeURIComponent(dateGte) +
      '&sort=date&order=asc&limit=' + SALES_API_PAGE + '&skip=' + skip;
    if (dateLte) url += '&date_lte=' + encodeURIComponent(dateLte);
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code !== 200) throw new Error('Sales API HTTP ' + code + ': ' + resp.getContentText().slice(0, 200));
    const body = JSON.parse(resp.getContentText());
    if (!body.success) throw new Error('Sales API error: ' + (body.message || 'unknown'));
    if (total === null) total = body.total || 0;

    const batch = body.data || [];
    if (batch.length === 0) break;   // guard against a misreported total

    for (let i = 0; i < batch.length; i++) {
      const doc = batch[i];
      const sku = String(doc.sku || '').trim();
      const uom = parseFloat(doc.uom);   // grams (or pcs) per pack
      const qty = parseFloat(doc.qty);   // number of packs
      if (!sku || isNaN(uom) || isNaN(qty) || uom === 0 || qty === 0) continue;

      const dateStr = parseAnyDateCell(doc.date);   // canonical DD/MM/YYYY (accepts ISO or DD/MM/YYYY)
      if (!dateStr) continue;

      rowSeq++;
      // Same shape the dashboard already consumes: [date, gramsPerPack(=uom), packs(=qty), ...]
      rows.push([dateStr, sku, uom, qty, rowSeq, 'mongo']);
    }

    fetched += batch.length;
    pages++;
    skip += SALES_API_PAGE;
  } while (fetched < total && pages < 100);   // safety cap: 100 pages (~100k rows)

  const tabsProcessed = [{ tab: 'mongo:treats-sales-saturday', rows: rows.length, total: total, pages: pages }];
  // dedupReport kept for response-shape compatibility; no cross-source dedup needed
  // now that Mongo is the single source.
  const dedupReport = { source: 'mongo', total: total, pages: pages };

  return { rowCount: rows.length, tabsProcessed, rows, dedupReport };
}

// Helper: convert a cell value into a canonical DD/MM/YYYY string in Asia/Singapore tz.
// Returns '' if the value can't be parsed.
//
// Accepts: Date objects, "DD/MM/YYYY", "D/M/YYYY", "DD-MM-YYYY", "YYYY-MM-DD".
// Always emits zero-padded "DD/MM/YYYY" so set-membership works reliably.
function parseAnyDateCell(val) {
  if (val == null || val === '') return '';
  let d = null;
  if (val instanceof Date) {
    d = val;
  } else {
    const s = String(val).trim();
    if (!s) return '';

    // Try DD/MM/YYYY or D/M/YYYY (with / or - separator)
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) {
      d = new Date(+m[3], +m[2] - 1, +m[1]);
    } else {
      // Try YYYY-MM-DD (ISO-ish)
      m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
      if (m) {
        d = new Date(+m[1], +m[2] - 1, +m[3]);
      } else {
        // Last-ditch: let JS parse it
        const parsed = new Date(s);
        if (!isNaN(parsed.getTime())) d = parsed;
      }
    }
  }
  if (!d || isNaN(d.getTime())) return '';
  // Use Utilities.formatDate to lock the timezone — protects against the script
  // running in a timezone different from where dates were originally entered.
  return Utilities.formatDate(d, 'Asia/Singapore', 'dd/MM/yyyy');
}

// Backwards-compat alias: existing code or future callers expecting formatCellDate.
function formatCellDate(val) {
  return parseAnyDateCell(val);
}

// --- InventoryLog reader ---
// Reads every row from the InventoryLog tab and returns as an array of objects.
// Columns: A Date, B EventType, C TreatName, D SKU, E Quantity, F UUID, G LinkedOrderId (optional, v11+)
function getInventoryLogData() {
  const ss = SpreadsheetApp.openById(SPREADSHEETID);
  const sheet = ss.getSheetByName(INVENTORY_LOG_TAB);
  if (!sheet) return [];   // no tab = no logs yet; not an error

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];   // header only

  // Read columns A:G from row 2 to lastRow. If sheet was created before v11
  // and only has 6 columns, getLastColumn() returns 6 and column G is empty.
  const lastCol = Math.max(6, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const out = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const uuid = row[5];
    if (!uuid) continue;   // skip blank rows

    // Date column — could be a Date object or a string
    const dateStr = formatCellDate(row[0]);

    out.push({
      date: dateStr,
      eventType: String(row[1]).trim(),
      treatName: String(row[2]).trim(),
      sku: String(row[3]).trim(),
      quantity: Number(row[4]),
      uuid: String(uuid).trim(),
      linkedOrderId: row[6] ? String(row[6]).trim() : ''   // v11: order linkage (Receipt rows only); '' for legacy/standalone
    });
  }
  return out;
}

// --- ProductionPlan reader (v10) ---
// Returns all rows from the ProductionPlan tab. The tab is auto-created on
// first write if it doesn't exist, so on first run this returns []. Rows are
// returned in the order they appear in the sheet (typically insertion order).
//
// Filtering by date range is done client-side — the volume is tiny (one row
// per planned slot, typically <50 active slots at a time across all weeks).
//
// Columns: A entryId | B weekStartDate | C dayDate | D shift | E sku |
//          F rawQty | G note | H createdAt | I fulfilledByUuid (v10)
function getProductionPlanData() {
  const ss = SpreadsheetApp.openById(SPREADSHEETID);
  const sheet = ss.getSheetByName(PRODUCTION_PLAN_TAB);
  if (!sheet) return [];   // tab doesn't exist yet → empty plan

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];   // header only

  // Read up to 9 cols — older sheets may only have 8, which is fine since
  // getRange will pad missing cells with empty. We'll just truncate to lastCol.
  const lastCol = Math.max(8, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const out = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const entryId = row[0];
    if (!entryId) continue;   // skip blanks
    out.push({
      entryId: String(entryId).trim(),
      weekStartDate: parseAnyDateCell(row[1]),
      dayDate: parseAnyDateCell(row[2]),
      shift: String(row[3] || '').trim(),
      sku: String(row[4] || '').trim(),
      rawQty: Number(row[5]) || 0,
      note: String(row[6] || '').trim(),
      createdAt: row[7] ? String(row[7]).trim() : '',
      fulfilledByUuid: row[8] ? String(row[8]).trim() : ''   // v10
    });
  }
  return out;
}

// Lazily creates the ProductionPlan tab if it's missing. Returns the sheet.
// Header row is written once on creation; subsequent calls are cheap no-ops.
// v10: existing sheets created by v9 will be auto-extended with column I header
// the next time setPlanEntry runs (see handleSetPlanEntry).
function ensureProductionPlanTab() {
  const ss = SpreadsheetApp.openById(SPREADSHEETID);
  let sheet = ss.getSheetByName(PRODUCTION_PLAN_TAB);
  if (sheet) return sheet;
  sheet = ss.insertSheet(PRODUCTION_PLAN_TAB);
  sheet.appendRow(['entryId', 'weekStartDate', 'dayDate', 'shift', 'sku', 'rawQty', 'note', 'createdAt', 'fulfilledByUuid']);
  sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  return sheet;
}

// ============================================================
// POST — multi-action endpoint (v9)
// Body must be JSON sent as Content-Type: text/plain (to avoid CORS preflight).
//
// v9 routing: optional `action` field selects what to do.
//   action='logInventory' (default if absent) — append to InventoryLog (v8 behaviour)
//   action='setPlanEntry'    — upsert a production plan row (by entryId)
//   action='deletePlanEntry' — remove a production plan row (by entryId)
//
// Common required fields:
//   secret — must match SYNC_SECRET above
//
// logInventory required fields:
//   uuid        — client-generated unique ID for idempotency
//   date        — DD/MM/YYYY string
//   eventType   — 'Stocktake' | 'Stock Discrepancy' | 'Receipt' | 'Production'
//                  Stocktake        = absolute physical count (signed; can be negative)
//                  Stock Discrepancy = variance from expected (signed diff)
//                  Receipt          = inbound from China shipment (positive)
//                  Production       = inhouse production output (positive)
//   treatName   — display name
//   sku         — SKU code
//   quantity    — number (signed: + for additions, - for deductions)
//
// setPlanEntry required fields:
//   entryId, weekStartDate, dayDate, shift, sku, rawQty (note optional)
// deletePlanEntry required fields:
//   entryId
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');

    // Auth check
    if (body.secret !== SYNC_SECRET) {
      return jsonResponse({ success: false, error: 'invalid secret' });
    }

    // Route by action — default to 'logInventory' for backward compatibility with
    // dashboards on v8 or earlier that don't send the action field.
    const action = body.action || 'logInventory';
    if (action === 'logInventory') return handleLogInventory(body);
    if (action === 'setPlanEntry') return handleSetPlanEntry(body);
    if (action === 'deletePlanEntry') return handleDeletePlanEntry(body);
    return jsonResponse({ success: false, error: 'unknown action: ' + action });

  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ---- v8 inventory log handler (extracted from old doPost) ----
// v11 (21/05/2026): added optional column G `linkedOrderId` to record which pending
// order a Receipt fulfills. Format: 's32' for seaOrder id 32, 'a107' for airOrder id 107,
// or empty for standalone receipts / non-Receipt event types.
function handleLogInventory(body) {
  // Validate required fields
  const required = ['uuid', 'date', 'eventType', 'treatName', 'sku', 'quantity'];
  for (const f of required) {
    if (body[f] === undefined || body[f] === null || body[f] === '') {
      return jsonResponse({ success: false, error: 'missing field: ' + f });
    }
  }

  const validEvents = ['Stock Discrepancy', 'Receipt', 'Production', 'Stocktake'];
  if (validEvents.indexOf(body.eventType) === -1) {
    return jsonResponse({ success: false, error: 'invalid eventType: ' + body.eventType });
  }

  const ss = SpreadsheetApp.openById(SPREADSHEETID);
  const sheet = ss.getSheetByName(INVENTORY_LOG_TAB);
  if (!sheet) return jsonResponse({ success: false, error: 'tab not found: ' + INVENTORY_LOG_TAB });

  // v11 schema-migration: ensure column G header exists. If old sheet has only 6
  // columns, add the linkedOrderId header so subsequent reads via getInventoryLogData
  // include it. Idempotent — re-running is a no-op when header already exists.
  if (sheet.getLastColumn() < 7) {
    sheet.getRange(1, 7).setValue('linkedOrderId');
  }

  // Idempotency check — scan last 200 rows of column F (uuid) for a match
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const startRow = Math.max(2, lastRow - 199);
    const numRows = lastRow - startRow + 1;
    const existing = sheet.getRange(startRow, 6, numRows, 1).getValues();
    for (let i = 0; i < existing.length; i++) {
      if (existing[i][0] === body.uuid) {
        return jsonResponse({ success: true, deduped: true, message: 'entry already logged' });
      }
    }
  }

  // Append row: A Date, B EventType, C TreatName, D SKU, E Quantity, F UUID, G LinkedOrderId
  sheet.appendRow([
    body.date,
    body.eventType,
    body.treatName,
    body.sku,
    Number(body.quantity),
    body.uuid,
    body.linkedOrderId || ''   // v11: optional. Empty string for Stocktake/Production/Discrepancy/standalone-Receipt.
  ]);

  return jsonResponse({ success: true, rowAppended: sheet.getLastRow() });
}

// ---- v9 production plan handlers ----
// Upsert (insert-or-update) a plan slot. Key = entryId.
// If a row with the same entryId exists, its fields are updated in place.
// Otherwise a new row is appended.
//
// v10: fulfilledByUuid is an optional 9th field. Pass '' to clear an existing
// attribution. If the sheet was created by v9 with only 8 columns, the header
// will be auto-extended on first write here.
function handleSetPlanEntry(body) {
  const required = ['entryId', 'weekStartDate', 'dayDate', 'shift', 'sku', 'rawQty'];
  for (const f of required) {
    if (body[f] === undefined || body[f] === null || body[f] === '') {
      return jsonResponse({ success: false, error: 'missing field: ' + f });
    }
  }
  const validShifts = ['morning', 'night', 'thaw'];
  if (validShifts.indexOf(body.shift) === -1) {
    return jsonResponse({ success: false, error: 'invalid shift: ' + body.shift });
  }

  const sheet = ensureProductionPlanTab();

  // v10 schema-migration: if existing sheet only has 8 columns (created by v9),
  // add the 9th header cell so future reads work cleanly.
  if (sheet.getLastColumn() < 9) {
    sheet.getRange(1, 9).setValue('fulfilledByUuid').setFontWeight('bold');
  }

  const lastRow = sheet.getLastRow();
  const createdAt = new Date().toISOString();
  const fulfilledByUuid = body.fulfilledByUuid || '';
  const rowValues = [
    body.entryId,
    body.weekStartDate,
    body.dayDate,
    body.shift,
    body.sku,
    Number(body.rawQty),
    body.note || '',
    createdAt,
    fulfilledByUuid
  ];

  // Look for existing entryId — if found, update in place
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(body.entryId).trim()) {
        // Preserve original createdAt — only update editable fields
        const origCreatedAt = sheet.getRange(i + 2, 8).getValue();
        rowValues[7] = origCreatedAt || createdAt;
        sheet.getRange(i + 2, 1, 1, 9).setValues([rowValues]);
        return jsonResponse({ success: true, updated: true, row: i + 2 });
      }
    }
  }

  // Otherwise append
  sheet.appendRow(rowValues);
  return jsonResponse({ success: true, inserted: true, row: sheet.getLastRow() });
}

// Delete a plan slot by entryId. Returns success even if not found (idempotent).
function handleDeletePlanEntry(body) {
  if (!body.entryId) return jsonResponse({ success: false, error: 'missing field: entryId' });
  const sheet = ensureProductionPlanTab();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ success: true, deleted: false, message: 'tab empty' });

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(body.entryId).trim()) {
      sheet.deleteRow(i + 2);
      return jsonResponse({ success: true, deleted: true, row: i + 2 });
    }
  }
  return jsonResponse({ success: true, deleted: false, message: 'entryId not found' });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Manual tests — run from Apps Script editor before deploying
// ============================================================
function testOutput() {
  const result = getCombinedJSON();
  Logger.log('Success: ' + result.success);
  Logger.log('Sales rows: ' + result.rowCount);
  Logger.log('InventoryLog rows: ' + result.inventoryLogCount);
  if (result.inventoryLog && result.inventoryLog.length > 0) {
    Logger.log('First inv row: ' + JSON.stringify(result.inventoryLog[0]));
    Logger.log('Last inv row: ' + JSON.stringify(result.inventoryLog[result.inventoryLog.length - 1]));
  }
}

function testInventoryLogRead() {
  const inv = getInventoryLogData();
  Logger.log('Total inventory log rows: ' + inv.length);
  inv.slice(0, 5).forEach(r => Logger.log(JSON.stringify(r)));
}

// v12: verifies the MongoDB sales pull end-to-end. Run from the editor after setting
// TREATS_SALES_API_TOKEN in Script Properties — confirms auth, pagination, and shape.
function testSalesApiRead() {
  const result = getSalesData();
  Logger.log('Sales rows pulled: ' + result.rowCount);
  Logger.log('API total reported: ' + (result.dedupReport && result.dedupReport.total));
  Logger.log('Pages fetched: ' + (result.dedupReport && result.dedupReport.pages));
  if (result.rows.length > 0) {
    Logger.log('First row: ' + JSON.stringify(result.rows[0]));   // [date, sku, uom, qty, seq, 'mongo']
    Logger.log('Last row:  ' + JSON.stringify(result.rows[result.rows.length - 1]));
    const dates = result.rows.map(r => r[0]);
    Logger.log('Date span: ' + dates[0] + ' → ' + dates[dates.length - 1]);
  } else {
    Logger.log('No rows returned — check token, SALES_SINCE_DATE, and that the API stores dates as YYYY-MM-DD.');
  }
}

function testPost() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        secret: SYNC_SECRET,
        uuid: 'test-' + Date.now(),
        date: '15/05/2026',
        eventType: 'Receipt',
        treatName: 'Test Treat',
        sku: 'TEST_SKU',
        quantity: 100
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
