const DATA_SHEET_ID    = '1j165dsa1a-DDapOCgyBLrJQ_UBa4LzCWdWez4_obLD0';
const ACTUALS_TAB      = 'Actuals_DummyData';
const INITIATIVE_TAB   = 'Initiative_DummyData';
const HISTORICAL_TAB   = 'Historical_Orphans';
const LOADING_IMG_ID   = '1kVuDLtVu_9rIgtl7VoXVo8nn_-x6fXN4';
const CACHE_KEY        = 'orphans_data_v1';
const IMG_CACHE_KEY    = 'orphans_img_v1';
const CACHE_TTL        = 21600; // 6 hours

// ── Entry point ────────────────────────────────────────────────────────────────
function doGet() {
  try {
    const data = getOrphansData();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: e.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Called from client via google.script.run.getOrphansData() ─────────────────
function getOrphansData() {
  const cache  = CacheService.getScriptCache();
  const chunks = getChunks_(cache);
  if (chunks) return JSON.parse(chunks);

  const ss = SpreadsheetApp.openById(DATA_SHEET_ID);

  // ── Actuals ────────────────────────────────────────────────────────────────
  const actualsSheet = ss.getSheetByName(ACTUALS_TAB);
  if (!actualsSheet) throw new Error('Tab not found: ' + ACTUALS_TAB);
  const actualsVals    = actualsSheet.getDataRange().getValues();
  const actualsHeaders = actualsVals[0].map(String);
  const actualsRows    = [];

  for (let i = 1; i < actualsVals.length; i++) {
    const row = {};
    actualsHeaders.forEach((h, j) => {
      const v = actualsVals[i][j];
      row[h]  = (v !== undefined && v !== null) ? String(v) : '';
    });
    if (row['Resource Name'] && row['Resource Name'].trim()) actualsRows.push(row);
  }

  // ── Initiatives ────────────────────────────────────────────────────────────
  const initSheet = ss.getSheetByName(INITIATIVE_TAB);
  if (!initSheet) throw new Error('Tab not found: ' + INITIATIVE_TAB);
  const initVals    = initSheet.getDataRange().getValues();
  const initHeaders = initVals[0].map(String);

  const jiraKeyIdx = initHeaders.findIndex(h =>
    /jira\s*key/i.test(h) || h.toLowerCase() === 'key' || h.toLowerCase() === 'jira key'
  );
  if (jiraKeyIdx === -1) {
    throw new Error('JIRA Key column not found in ' + INITIATIVE_TAB + '. Headers: ' + initHeaders.join(', '));
  }

  const initiativeMap  = {};
  const initiativeKeys = [];
  for (let i = 1; i < initVals.length; i++) {
    const key = String(initVals[i][jiraKeyIdx] || '').trim();
    if (!key) continue;
    initiativeKeys.push(key);
    const row = {};
    initHeaders.forEach((h, j) => { row[h] = String(initVals[i][j] || ''); });
    initiativeMap[key] = row;
  }

  const result = {
    actualsRows,
    actualsHeaders,
    initiativeKeys,
    initiativeMap,
    initHeaders
  };

  putChunks_(cache, JSON.stringify(result));
  return result;
}

// ── Called from client via google.script.run.getLoadingImage() ────────────────
// Returns base64 data URI for the loading screen background image.
function getLoadingImage() {
  try {
    const cache  = CacheService.getScriptCache();
    const cached = cache.get(IMG_CACHE_KEY);
    if (cached) return cached;

    const blob    = DriveApp.getFileById(LOADING_IMG_ID).getBlob();
    const mime    = blob.getContentType() || 'image/jpeg';
    const dataUri = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());

    // Only cache if encoded string fits in one CacheService entry
    if (dataUri.length < 90000) cache.put(IMG_CACHE_KEY, dataUri, CACHE_TTL);
    return dataUri;
  } catch (e) {
    console.log('Image load failed:', e);
    return ''; // client will use CSS fallback
  }
}

// ── Cache helpers ──────────────────────────────────────────────────────────────
function putChunks_(cache, json) {
  try {
    const CHUNK = 90000;
    const total = Math.ceil(json.length / CHUNK);
    const pairs = { '__orphan_chunks__': String(total) };
    for (let i = 0; i < total; i++) {
      pairs[CACHE_KEY + '_' + i] = json.slice(i * CHUNK, (i + 1) * CHUNK);
    }
    cache.putAll(pairs, CACHE_TTL);
  } catch (e) {
    console.log('Cache write failed:', e);
  }
}

function getChunks_(cache) {
  try {
    const meta = cache.get('__orphan_chunks__');
    if (!meta) return null;
    const total  = parseInt(meta);
    const keys   = Array.from({ length: total }, (_, i) => CACHE_KEY + '_' + i);
    const stored = cache.getAll(keys);
    if (Object.keys(stored).length !== total) return null;
    return keys.map(k => stored[k]).join('');
  } catch (e) {
    return null;
  }
}

// ── Historical data (called from Trend modal) ─────────────────────────────────
function getHistoricalData() {
  const ss    = SpreadsheetApp.openById(DATA_SHEET_ID);
  const sheet = ss.getSheetByName(HISTORICAL_TAB);
  if (!sheet) return [];
  const vals    = sheet.getDataRange().getValues();
  const headers = vals[0].map(String);
  return vals.slice(1)
    .map(row => {
      const r = {};
      headers.forEach((h, j) => r[h] = String(row[j] !== undefined && row[j] !== null ? row[j] : ''));
      return r;
    })
    .filter(r => r['Quarter'] && r['Quarter'].trim());
}

// ── Cache warm-up (called by time-driven trigger) ─────────────────────────────
// Clears and rebuilds the data cache so end users never hit a cold load.
function warmCache() {
  clearOrphansCache();
  getOrphansData();
  Logger.log('Cache warmed at ' + new Date().toLocaleString());
}

// ── Run once from Script Editor to install the warm-up trigger ────────────────
// Schedules warmCache() every 4 hours so the cache never expires for end users.
function createWarmCacheTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'warmCache')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('warmCache')
    .timeBased()
    .everyHours(4)
    .create();

  Logger.log('Warm-cache trigger created — warmCache() will run every 4 hours.');
}

// ── Utilities (run from editor) ────────────────────────────────────────────────
function clearOrphansCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('__orphan_chunks__');
  cache.remove(IMG_CACHE_KEY);
  Logger.log('Cache cleared.');
}

function testDataAccess() {
  clearOrphansCache();
  const data = getOrphansData();
  Logger.log('Actuals rows: '    + data.actualsRows.length);
  Logger.log('Initiative keys: ' + data.initiativeKeys.length);
  Logger.log('Actuals headers: ' + data.actualsHeaders.join(' | '));
}

// ── Run once from the Script Editor to create a demo-friendly distribution ────
// Makes the first 2 L4 managers green (<5% orphan), next 2 yellow (5-15%),
// and leaves the rest red, by populating Roadmap Mapping for their actuals rows.
function seedDemoDistribution() {
  const ss           = SpreadsheetApp.openById(DATA_SHEET_ID);
  const actualsSheet = ss.getSheetByName(ACTUALS_TAB);
  const initSheet    = ss.getSheetByName(INITIATIVE_TAB);

  // Gather valid initiative keys to use as Roadmap Mapping values
  const initVals   = initSheet.getDataRange().getValues();
  const initHdrs   = initVals[0].map(String);
  const jiraIdx    = initHdrs.findIndex(h => /jira\s*key/i.test(h) || h.toLowerCase() === 'key');
  const initKeys   = initVals.slice(1).map(r => String(r[jiraIdx] || '').trim()).filter(k => k);
  if (!initKeys.length) throw new Error('No initiative keys found in ' + INITIATIVE_TAB);

  // Read actuals
  const vals    = actualsSheet.getDataRange().getValues();
  const headers = vals[0].map(String);
  const rmIdx   = headers.indexOf('Roadmap Mapping');
  const l4Idx   = headers.indexOf('Level 4 Mgr');
  if (rmIdx === -1)  throw new Error('Roadmap Mapping column not found');
  if (l4Idx === -1)  throw new Error('Level 4 Mgr column not found');

  // Collect L4 managers in order of first appearance
  const l4Order = [], l4Seen = new Set();
  for (let i = 1; i < vals.length; i++) {
    const l4 = String(vals[i][l4Idx] || '').trim();
    if (l4 && !l4Seen.has(l4)) { l4Order.push(l4); l4Seen.add(l4); }
  }

  // Assign link rates so the OVERALL orphan % stays under 5% (triggers confetti on load).
  // The first 2–3 managers typically hold the bulk of rows; keeping them very green
  // (≈1%) ensures the weighted overall stays below 5% even when later managers
  // show YELLOW or RED.  Visual variety targets:
  //   i=0,1  → ~1%  orphan  → GREEN (star performers)
  //   i=2    → ~4%  orphan  → GREEN (solid)
  //   i=3    → ~10% orphan  → YELLOW
  //   i=4    → ~18% orphan  → RED
  //   rest   → ~15% orphan  → RED / high YELLOW
  const linkRate = {};
  l4Order.forEach((mgr, i) => {
    if      (i < 2) linkRate[mgr] = 0.990;  // ~1%  orphan → GREEN (star performers)
    else if (i < 3) linkRate[mgr] = 0.960;  // ~4%  orphan → GREEN (solid)
    else if (i < 4) linkRate[mgr] = 0.900;  // ~10% orphan → YELLOW
    else if (i < 5) linkRate[mgr] = 0.820;  // ~18% orphan → RED
    else            linkRate[mgr] = 0.850;  // ~15% orphan → RED
  });

  // Build new Roadmap Mapping column values, tracking running link ratio per manager
  const linked = {}, total = {};
  const newColumn = vals.slice(1).map((row, rowIdx) => {
    const l4   = String(row[l4Idx] || '').trim();
    const rate = linkRate[l4] || 0.93;  // default: 93% linked if manager not in order list

    total[l4]  = (total[l4]  || 0) + 1;
    linked[l4] = (linked[l4] || 0);

    const currentLinkRatio = linked[l4] / total[l4];
    if (currentLinkRatio < rate) {
      linked[l4]++;
      return [initKeys[rowIdx % initKeys.length]];
    }
    return [''];
  });

  actualsSheet.getRange(2, rmIdx + 1, newColumn.length, 1).setValues(newColumn);
  clearOrphansCache();

  const summary = l4Order.map((m, i) => {
    const rate = i < 2 ? '~1% orphan (GREEN ⭐)' : i < 3 ? '~4% orphan (GREEN)' : i < 4 ? '~10% orphan (YELLOW)' : i < 5 ? '~18% orphan (RED)' : '~15% orphan (RED)';
    return `${m}: ${rate}`;
  }).join('\n');
  Logger.log('Done! Distribution seeded:\n' + summary + '\nOverall target: <5% → confetti on load!\nCache cleared — reload the web app.');
}

// ── Run once from the Script Editor to create the Capex Targets By Resource tab ─
// Reads unique Resource Names from Actuals_DummyData and assigns a randomised
// CapEx % target (between 20% and 80%, rounded to nearest 5%) to each.
function createCapexTargetsByResource() {
  const ss           = SpreadsheetApp.openById(DATA_SHEET_ID);
  const actualsSheet = ss.getSheetByName(ACTUALS_TAB);
  if (!actualsSheet) throw new Error('Tab not found: ' + ACTUALS_TAB);

  // Collect unique resource names
  const allVals = actualsSheet.getDataRange().getValues();
  const headers = allVals[0].map(String);
  const nameIdx = headers.indexOf('Resource Name');
  if (nameIdx === -1) throw new Error('"Resource Name" column not found in ' + ACTUALS_TAB);

  const seen = new Set();
  const resources = [];
  for (let i = 1; i < allVals.length; i++) {
    const name = String(allVals[i][nameIdx] || '').trim();
    if (name && !seen.has(name)) { seen.add(name); resources.push(name); }
  }
  if (!resources.length) throw new Error('No resource names found in ' + ACTUALS_TAB);

  // Create (or clear) the target tab
  const TAB_NAME = 'Capex Targets By Resource';
  let targetSheet = ss.getSheetByName(TAB_NAME);
  if (targetSheet) {
    targetSheet.clearContents();
  } else {
    targetSheet = ss.insertSheet(TAB_NAME);
  }

  // Build rows: header + one row per resource with a randomised CapEx % target
  const rows = [['Resource Name', 'Capex Target %']];
  resources.forEach(name => {
    // Random target between 20% and 80%, rounded to nearest 5%
    const raw    = Math.random() * (0.80 - 0.20) + 0.20;
    const target = Math.round(raw * 20) / 20;  // rounds to nearest 0.05
    rows.push([name, target]);
  });

  targetSheet.getRange(1, 1, rows.length, 2).setValues(rows);

  // Light formatting: bold header, percentage format on target column
  targetSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  targetSheet.getRange(2, 2, rows.length - 1, 1).setNumberFormat('0%');
  targetSheet.autoResizeColumns(1, 2);

  Logger.log('Done! Created "' + TAB_NAME + '" with ' + resources.length + ' resources.');
}

// ── Run once from the Script Editor to add a randomised CapEx/OpEx column ─────
// ~35% of rows → CapEx only, ~35% → OpEx only, ~30% → both (row duplicated).
// Rewrites the entire sheet (header row preserved), then clears cache.
function addCapExOpExColumn() {
  const ss    = SpreadsheetApp.openById(DATA_SHEET_ID);
  const sheet = ss.getSheetByName(ACTUALS_TAB);
  if (!sheet) throw new Error('Tab not found: ' + ACTUALS_TAB);

  const allVals = sheet.getDataRange().getValues();
  const headers = allVals[0].map(String);

  // Abort if column already exists
  if (headers.includes('CapEx/OpEx')) {
    Logger.log('CapEx/OpEx column already exists — aborting.');
    return;
  }

  const newHeader  = headers.concat(['CapEx/OpEx']);
  const outputRows = [newHeader];

  for (let i = 1; i < allVals.length; i++) {
    const row = allVals[i];
    const r   = Math.random();

    if (r < 0.35) {
      // CapEx only
      outputRows.push(row.concat(['CapEx']));
    } else if (r < 0.70) {
      // OpEx only
      outputRows.push(row.concat(['OpEx']));
    } else {
      // Both — emit two rows
      outputRows.push(row.concat(['CapEx']));
      outputRows.push(row.concat(['OpEx']));
    }
  }

  // Rewrite the sheet
  sheet.clearContents();
  sheet.getRange(1, 1, outputRows.length, newHeader.length).setValues(outputRows);
  clearOrphansCache();

  const originalCount = allVals.length - 1;
  const newCount      = outputRows.length - 1;
  Logger.log(
    `Done! Original rows: ${originalCount} → New rows: ${newCount} ` +
    `(${newCount - originalCount} rows doubled for both CapEx & OpEx). Cache cleared.`
  );
}

// ── Run once from the Script Editor to seed Q-o-Q historical orphan data ──────
// Creates (or overwrites) the Historical_Orphans tab with quarterly snapshots
// from Q1 2024 through Q1 2026, showing a realistic downward trend.
function seedHistoricalOrphans() {
  const ss = SpreadsheetApp.openById(DATA_SHEET_ID);
  let sheet = ss.getSheetByName(HISTORICAL_TAB);
  if (!sheet) sheet = ss.insertSheet(HISTORICAL_TAB);

  // Columns: Quarter | Total Hours | Orphan Hours | Orphan %
  const data = [
    ['Quarter', 'Total Hours', 'Orphan Hours', 'Orphan %'],
    ['Q1 2024', 44800,  6720,  0.150],   // 15.0% — RED, early hygiene issues
    ['Q2 2024', 46100,  5532,  0.120],   // 12.0% — RED, slight improvement
    ['Q3 2024', 47300,  4257,  0.090],   // 9.0%  — YELLOW, clean-up underway
    ['Q4 2024', 48400,  3388,  0.070],   // 7.0%  — YELLOW, trending down
    ['Q1 2025', 49600,  2976,  0.060],   // 6.0%  — YELLOW, just above goal
    ['Q2 2025', 50800,  2794,  0.055],   // 5.5%  — YELLOW, near goal
    ['Q3 2025', 52000,  2496,  0.048],   // 4.8%  — GREEN, first quarter under 5%!
    ['Q4 2025', 53200,  2234,  0.042],   // 4.2%  — GREEN, holding gains
    ['Q1 2026', 54500,  2071,  0.038],   // 3.8%  — GREEN, current period
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log(
    'Historical_Orphans tab seeded with ' + (data.length - 1) + ' quarters of data ' +
    '(Q1 2024 → Q1 2026). Open the Trend modal in the web app to view the chart.'
  );
}
