// VMM Dashboard — Google Apps Script (optimised + caching)
// Only sends the columns the dashboard actually uses — faster transfer
// Cache stores result for 6 hours. Run clearCache() after updating sheet data.

const CACHE_KEY = 'vmm_dashboard_data_v2';
const CACHE_DURATION = 21600; // 6 hours

// Only these columns are sent to the dashboard (keeps payload small & fast)
const NEEDED_COLS = new Set([
  'Log Date', 'Closer Date', 'Current Status', 'Zone',
  'Product Vendor', 'Store Name', 'Product Name', 'New Product Type',
  'Over Due TAT', 'Last Delay reason', 'Aging',
  'FM Name-U', 'FM Name', 'Store Region', 'State'
]);

function doGet(e) {
  try {
    const cache = CacheService.getScriptCache();

    // Try reading from cache chunks first
    const numChunks = cache.get(CACHE_KEY + '_chunks');
    if (numChunks) {
      const n = parseInt(numChunks);
      let json = '';
      for (let i = 0; i < n; i++) {
        const chunk = cache.get(CACHE_KEY + '_' + i);
        if (!chunk) { json = null; break; }
        json += chunk;
      }
      if (json) return out(JSON.parse(json));
    } else {
      const cached = cache.get(CACHE_KEY);
      if (cached) return out(JSON.parse(cached));
    }

    // Cache miss — read from sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    if (data.length < 2) return out({data:[], total:0});

    const allHeaders = data[0].map(h => String(h).trim());
    // Only keep indices of needed columns
    const colIdx = [];
    const colNames = [];
    allHeaders.forEach((h, i) => {
      if (NEEDED_COLS.has(h)) { colIdx.push(i); colNames.push(h); }
    });

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue;
      const obj = {};
      colIdx.forEach((ci, j) => {
        const v = row[ci];
        obj[colNames[j]] = (v instanceof Date)
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd-MM-yyyy')
          : v;
      });
      rows.push(obj);
    }

    const payload = {data: rows, total: rows.length};
    const json = JSON.stringify(payload);

    // Store in cache (split into chunks if >100KB)
    if (json.length < 100000) {
      cache.put(CACHE_KEY, json, CACHE_DURATION);
    } else {
      const chunks = [];
      let pos = 0;
      while (pos < json.length) {
        chunks.push(json.slice(pos, pos + 90000));
        pos += 90000;
      }
      cache.put(CACHE_KEY + '_chunks', String(chunks.length), CACHE_DURATION);
      chunks.forEach((chunk, i) => cache.put(CACHE_KEY + '_' + i, chunk, CACHE_DURATION));
    }

    return out(payload);
  } catch(err) {
    return out({data:[], total:0, error: err.toString()});
  }
}

function clearCache() {
  const cache = CacheService.getScriptCache();
  const numChunks = cache.get(CACHE_KEY + '_chunks');
  if (numChunks) {
    const n = parseInt(numChunks);
    for (let i = 0; i < n; i++) cache.remove(CACHE_KEY + '_' + i);
    cache.remove(CACHE_KEY + '_chunks');
  }
  cache.remove(CACHE_KEY);
  Logger.log('Cache cleared. Next request will fetch fresh data.');
}

function out(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
