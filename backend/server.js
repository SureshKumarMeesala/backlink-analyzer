// server/index.js
// CSV-only Backlink Checker server (SSE + concurrent workers + Excel output)

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');
const axios = require('axios');
const cheerio = require('cheerio');
const fs = require('fs');
const path = require('path');

const UPLOAD_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const app = express();
app.use(cors()); // dev: allow all origins (restrict in production)
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// config - tune as needed
const DEFAULT_BACKLINK = 'https://bonfireinstituteofdesign.com';
let CONCURRENCY = 6;      // parallel workers
let DELAY_MS = 200;       // delay between requests per worker (ms)
const FETCH_TIMEOUT = 20000; // axios timeout ms

// multer: save uploads to uploads/
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});
const upload = multer({ storage });

// SSE clients
let clients = [];

// ---------- SSE endpoints ----------
app.get('/api/events', (req, res) => {
  res.set({
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache, no-transform',
    'Connection': 'keep-alive',
    'Access-Control-Allow-Origin': '*'
  });
  res.flushHeaders();

  // initial info for client
  res.write('retry: 10000\n\n');
  res.write(`data: ${JSON.stringify({ type: 'sse-connected', ts: Date.now() })}\n\n`);

  const clientId = Date.now() + Math.random();
  const newClient = { id: clientId, res };
  clients.push(newClient);
  console.log('SSE client connected, total clients:', clients.length);

  // heartbeat to keep proxies from closing the connection
  const keepAlive = setInterval(() => {
    try { res.write(':\n\n'); } catch (e) {}
  }, 15000);

  req.on('close', () => {
    clearInterval(keepAlive);
    clients = clients.filter(c => c.id !== clientId);
    console.log('SSE client disconnected, total clients:', clients.length);
  });
});

// simple test SSE route (for curl/browser)
app.get('/api/test-sse', (req, res) => {
  res.set({
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache, no-transform',
    'Connection': 'keep-alive',
    'Access-Control-Allow-Origin': '*'
  });
  res.flushHeaders();
  let n = 0;
  const iv = setInterval(() => {
    res.write(`data: ${JSON.stringify({ ping: ++n, ts: Date.now() })}\n\n`);
    if (n >= 6) {
      clearInterval(iv);
      try { res.end(); } catch (e) {}
    }
  }, 1000);
});

// helper to broadcast SSE data to connected clients
function sendSse(data) {
  const json = JSON.stringify(data);
  clients.forEach(c => {
    try { c.res.write(`data: ${json}\n\n`); } catch (e) { /* ignore */ }
  });
}

// ---------- helpers ----------

// add http:// if missing and return normalized url string, or null if invalid
function normalizeAndValidateUrl(value) {
  if (!value) return null;
  let s = String(value).trim();
  if (!s) return null;
  // If CSV contains "URL,description" we take only first column (caller should pass the first column)
  // Add http if scheme missing
  if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(s)) {
    s = 'http://' + s;
  }
  try {
    // new URL will throw if invalid
    new URL(s);
    return s;
  } catch (e) {
    return null;
  }
}

// fetch page and check for backlink using axios + cheerio
async function fetchAndDetect(url, backlink) {
  const normalized = String(url).trim();
  try {
    const resp = await axios.get(normalized, {
      timeout: FETCH_TIMEOUT,
      maxRedirects: 5,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
    });
    const status = resp.status;
    const isWorking = status >= 200 && status < 400;
    const body = (typeof resp.data === 'string') ? resp.data : JSON.stringify(resp.data);

    // search for anchor first
    const $ = cheerio.load(body);
    const esc = (backlink || DEFAULT_BACKLINK).replace(/\/$/, '');
    let found = false;
    let snippet = '';

    $('a[href]').each((i, el) => {
      const href = $(el).attr('href');
      if (!href) return;
      if (href.indexOf(esc) !== -1 || href.indexOf(new URL(esc, 'http://example.com').hostname) !== -1) {
        found = true;
        const text = $(el).text().trim().replace(/\s+/g, ' ');
        snippet = `${href} | ${text}`.slice(0, 500);
        return false; // break
      }
    });

    if (!found) {
      const idx = body.indexOf(esc);
      if (idx !== -1) {
        found = true;
        const start = Math.max(0, idx - 120);
        const end = Math.min(body.length, idx + 120);
        snippet = body.slice(start, end).replace(/\s+/g, ' ').slice(0, 500);
      }
    }

    return { status, isWorking, found, snippet };
  } catch (err) {
    let msg = err.toString();
    if (err.response) msg = `HTTP ${err.response.status}`;
    return { status: err.response ? err.response.status : 'ERR', isWorking: false, found: false, snippet: msg };
  }
}

// ---------- worker queue ----------
async function runQueue(urlList, backlink, workbook, sheet, outFilePath) {
  let index = 0;

  async function worker(id) {
    while (true) {
      const i = index++;
      if (i >= urlList.length) break;
      const rowNumber = i + 2; // Excel-style rows: header row=1
      const url = String(urlList[i] || '').trim();

      console.log(`[worker ${id}] processing row=${rowNumber} url=${url}`);
      sendSse({ type: 'rowStart', row: rowNumber, url });

      const res = await fetchAndDetect(url, backlink);

      // write results
      try {
        const row = sheet.getRow(rowNumber);
        row.getCell(1).value = url;              // A (normalized URL)
        row.getCell(3).value = res.status;       // C
        row.getCell(4).value = res.isWorking ? 'Yes' : 'No'; // D
        row.getCell(5).value = res.found ? 'Yes' : 'No';     // E
        row.getCell(6).value = res.snippet || '';           // F
        row.getCell(7).value = new Date();                  // G
        row.commit();
      } catch (e) {
        console.error('Error writing to sheet row', rowNumber, e);
      }

      // periodic save
      if ((i % 20) === 0) {
        try { await workbook.xlsx.writeFile(outFilePath); } catch (e) { console.error('Save error', e); }
      }

      sendSse({ type: 'rowDone', row: rowNumber, url, result: res });
      console.log(`[worker ${id}] done row=${rowNumber} result=`, res);

      // polite delay
      await new Promise(r => setTimeout(r, DELAY_MS));
    }
  }

  const workers = [];
  for (let w = 0; w < CONCURRENCY; w++) workers.push(worker(w));
  await Promise.all(workers);
  // final save
  await workbook.xlsx.writeFile(outFilePath);
}

// ---------- upload handler for CSV only ----------
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    const backlink = req.body.backlink || DEFAULT_BACKLINK;
    if (!req.file) return res.status(400).json({ ok: false, error: 'Missing file' });

    const infile = req.file.path;
    const fname = req.file.originalname || '';
    const ext = (fname.split('.').pop() || '').toLowerCase();
    if (ext !== 'csv') {
      return res.status(400).json({ ok: false, error: 'Only CSV files are accepted by this endpoint. Use .csv' });
    }

    // read CSV text
    const txt = fs.readFileSync(infile, 'utf8');
    const lines = txt.split(/\r?\n/);

    // detect header row: if first line contains "url" (case-insensitive), skip it
    let startIdx = 0;
    if (lines.length > 0 && /url/i.test(lines[0])) startIdx = 1;

    // parse CSV first column robustly (handles quoted values and commas inside quotes)
    function parseFirstColumn(line) {
      line = line.trim();
      if (!line) return '';
      // If quoted:
      if (line[0] === '"') {
        // find the closing quote that is not doubled
        let i = 1;
        let value = '';
        while (i < line.length) {
          if (line[i] === '"') {
            if (line[i+1] === '"') { value += '"'; i += 2; continue; } // escaped quote
            i++; break; // end quote
          } else {
            value += line[i++];
          }
        }
        return value;
      }
      // not quoted, simple split on comma
      return line.split(',')[0].trim();
    }

    const urlItems = []; // {rowNumber, raw, url}
    const badRows = [];

    for (let i = startIdx; i < lines.length; i++) {
      const rowNum = i + 1; // human-friendly row number
      const rawLine = lines[i];
      if (!rawLine || !rawLine.trim()) continue;
      const firstCol = parseFirstColumn(rawLine);
      if (!firstCol) continue;
      const normalized = normalizeAndValidateUrl(firstCol);
      if (normalized) {
        urlItems.push({ rowNumber: rowNum + (startIdx === 1 ? 1 : 0), raw: firstCol, url: normalized });
      } else {
        badRows.push({ rowNumber: rowNum, raw: firstCol, reason: 'Invalid URL' });
      }
    }

    // create workbook and sheet and prefill column A with normalized URLs
    const outFileName = `results-${Date.now()}.xlsx`;
    const outFilePath = path.join(UPLOAD_DIR, outFileName);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');

    // header row
    sheet.getRow(1).values = ['URL', '', 'HTTP Status', 'Working?', 'Backlink Present?', 'Found Snippet', 'Last Checked'];

    // write URLs starting at row 2 (we'll align urlItems sequentially)
    for (let idx = 0; idx < urlItems.length; idx++) {
      const excelRow = idx + 2;
      const r = sheet.getRow(excelRow);
      r.getCell(1).value = urlItems[idx].url;
      r.commit();
    }

    await workbook.xlsx.writeFile(outFilePath);

    // immediate response: include counts and badRows preview
    res.json({
      ok: true,
      message: 'Processing started (CSV)',
      total: urlItems.length,
      badRowsCount: badRows.length,
      badRows: badRows.slice(0, 20),
      download: `/api/download?file=${outFileName}`
    });

    // SSE: announce start and any bad rows
    sendSse({ type: 'start', total: urlItems.length, backlink, badRowsCount: badRows.length });
    if (badRows.length) sendSse({ type: 'badRows', badRows: badRows.slice(0, 100) });

    // run processing in background
    (async () => {
      try {
        const urlList = urlItems.map(it => it.url);
        await runQueue(urlList, backlink, workbook, sheet, outFilePath);
        sendSse({ type: 'done', download: `/api/download?file=${outFileName}` });
        console.log('Processing complete, output:', outFilePath);
      } catch (err) {
        console.error('Processing error (background)', err);
        sendSse({ type: 'error', message: String(err) });
      }
    })();

  } catch (err) {
    console.error('Upload handler error', err);
    res.status(500).json({ ok: false, error: String(err) });
  }
});

// download endpoint
app.get('/api/download', (req, res) => {
  const file = req.query.file;
  if (!file) return res.status(400).send('Missing file query param');
  const filePath = path.join(UPLOAD_DIR, path.basename(file));
  if (!fs.existsSync(filePath)) return res.status(404).send('File not found');
  res.download(filePath);
});

// optional endpoint to tune concurrency/delay at runtime
app.post('/api/config', (req, res) => {
  const { concurrency, delay } = req.body || {};
  if (concurrency) CONCURRENCY = Math.max(1, parseInt(concurrency, 10));
  if (delay) DELAY_MS = Math.max(0, parseInt(delay, 10));
  res.json({ ok: true, CONCURRENCY, DELAY_MS });
});














// --- add after existing routes in server/index.js ---

// helper: find newest results-*.xlsx in uploads
function getLatestResultFile() {
  const files = fs.readdirSync(UPLOAD_DIR).filter(f => /^results-\d+\.xlsx$/.test(f));
  if (files.length === 0) return null;
  // sort by timestamp in filename (newest last)
  files.sort((a,b) => {
    const ta = Number(a.match(/^results-(\d+)\.xlsx$/)[1]);
    const tb = Number(b.match(/^results-(\d+)\.xlsx$/)[1]);
    return ta - tb;
  });
  return path.join(UPLOAD_DIR, files[files.length - 1]);
}

// Parse an ExcelJS cell safe
function readCellSafe(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.text) return String(v.text).trim();
    if (v.hyperlink) return String(v.hyperlink).trim();
    if (v.richText) return v.richText.map(x => x.text).join('').trim();
    if (v.result) return String(v.result).trim();
    try { return String(v).trim(); } catch (e) { return ''; }
  }
  return String(v).trim();
}

// GET /api/summary
app.get('/api/summary', async (req, res) => {
  try {
    const f = getLatestResultFile();
    if (!f) return res.json({ ok: true, total: 0, processed: 0, ok: 0, errors: 0, backlinks: 0, file: null });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(f);
    const sheet = workbook.worksheets[0];

    let total = 0, processed = 0, ok = 0, errors = 0, backlinks = 0;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber < 2) return; // skip header
      const url = readCellSafe(row.getCell(1));
      if (!url) return;
      total++;
      const statusRaw = readCellSafe(row.getCell(3));
      const status = (statusRaw === '') ? null : statusRaw;
      const working = readCellSafe(row.getCell(4)); // "Yes" / "No"
      const backlinkPresent = readCellSafe(row.getCell(5)); // "Yes" / "No"
      const lastChecked = readCellSafe(row.getCell(7));
      if (lastChecked) {
        processed++;
        // count OK
        const st = String(status || '');
        if (/^\d+$/.test(st)) {
          const n = Number(st);
          if (n >= 200 && n < 400) ok++;
          else errors++;
        } else {
          // treat 'ERR' or other non-2xx as error
          if (st.toUpperCase() === 'ERR' || st === '') errors++;
        }
        if (/^Yes$/i.test(backlinkPresent)) backlinks++;
      }
    });

    res.json({ ok: true, total, processed, ok, errors, backlinks, file: path.basename(f) });
  } catch (err) {
    console.error('Summary error', err);
    res.status(500).json({ ok: false, error: String(err) });
  }
});

// GET /api/rows?offset=0&limit=100
app.get('/api/rows', async (req, res) => {
  try {
    const offset = Math.max(0, parseInt(req.query.offset || '0', 10));
    const limit = Math.min(1000, Math.max(1, parseInt(req.query.limit || '100', 10)));
    const f = getLatestResultFile();
    if (!f) return res.json({ ok: true, rows: [], totalRows: 0, file: null });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(f);
    const sheet = workbook.worksheets[0];

    // collect all rows (rowNumber, url, status, isWorking, found, snippet, lastChecked)
    const all = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber < 2) return; // skip header
      const url = readCellSafe(row.getCell(1));
      if (!url) return;
      const status = readCellSafe(row.getCell(3));
      const isWorking = readCellSafe(row.getCell(4));
      const found = readCellSafe(row.getCell(5));
      const snippet = readCellSafe(row.getCell(6));
      const lastChecked = readCellSafe(row.getCell(7));
      all.push({
        row: rowNumber,
        url,
        status,
        isWorking,
        found,
        snippet,
        lastChecked
      });
    });

    const totalRows = all.length;
    const page = all.slice(offset, offset + limit);
    res.json({ ok: true, file: path.basename(f), totalRows, offset, limit, rows: page });
  } catch (err) {
    console.error('Rows error', err);
    res.status(500).json({ ok: false, error: String(err) });
  }
});









// requires at top of file:
// const https = require('https');

async function fetchAndDetect(url, backlink) {
  // ensure url is string and trimmed
  const raw = String(url || '').trim();
  const tried = { secure: false, insecureAgent: false, httpFallback: false };

  // helper to perform axios GET with options
  async function doGet(u, options = {}) {
    return axios.get(u, Object.assign({
      timeout: FETCH_TIMEOUT,
      maxRedirects: 5,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
    }, options));
  }

  // core detection logic (same as before)
  function detectFromBody(body) {
    try {
      const $ = cheerio.load(body || '');
      const esc = (backlink || DEFAULT_BACKLINK).replace(/\/$/, '');
      // anchor-first
      let found = false;
      let snippet = '';
      $('a[href]').each((i, el) => {
        const href = $(el).attr('href');
        if (!href) return;
        if (href.indexOf(esc) !== -1 || href.indexOf(new URL(esc, 'http://example.com').hostname) !== -1) {
          found = true;
          const text = $(el).text().trim().replace(/\s+/g, ' ');
          snippet = `${href} | ${text}`.slice(0, 500);
          return false;
        }
      });
      if (!found) {
        const idx = (body || '').indexOf(esc);
        if (idx !== -1) {
          found = true;
          const start = Math.max(0, idx - 120);
          const end = Math.min(body.length, idx + 120);
          snippet = (body.slice(start, end)).replace(/\s+/g, ' ').slice(0, 500);
        }
      }
      return { found, snippet };
    } catch (e) {
      return { found: false, snippet: '' };
    }
  }

  // 1) Try secure request first
  try {
    const resp = await doGet(raw);
    const status = resp.status;
    const isWorking = status >= 200 && status < 400;
    const body = (typeof resp.data === 'string') ? resp.data : JSON.stringify(resp.data);
    const det = detectFromBody(body);
    return { status, isWorking, found: det.found, snippet: det.snippet };
  } catch (err) {
    // check for certificate / TLS errors
    const errMsg = (err && err.toString()) || '';
    const certError = errMsg.match(/unable to get local issuer certificate/i)
      || errMsg.match(/UNABLE_TO_VERIFY_LEAF_SIGNATURE/i)
      || errMsg.match(/UNABLE_TO_GET_ISSUER_CERT/i)
      || err.code === 'DEPTH_ZERO_SELF_SIGNED_CERT'
      || err.code === 'UNABLE_TO_VERIFY_LEAF_SIGNATURE';

    if (certError) {
      console.warn('Certificate validation failed for', raw, '-> will retry with relaxed TLS (insecure). Error:', errMsg);
      tried.secure = true;

      // 2) Retry with insecure https agent (rejectUnauthorized: false)
      try {
        const resp2 = await doGet(raw, { httpsAgent: new (require('https').Agent)({ rejectUnauthorized: false }) });
        tried.insecureAgent = true;
        const status = resp2.status;
        const isWorking = status >= 200 && status < 400;
        const body = (typeof resp2.data === 'string') ? resp2.data : JSON.stringify(resp2.data);
        const det = detectFromBody(body);
        return { status, isWorking, found: det.found, snippet: det.snippet, note: 'insecure-https-agent' };
      } catch (err2) {
        // if https->http fallback desired, try that
        console.warn('Insecure HTTPS retry failed for', raw, 'Error:', err2 && err2.toString());
        if (/^https:\/\//i.test(raw)) {
          const httpUrl = raw.replace(/^https:/i, 'http:');
          try {
            const resp3 = await doGet(httpUrl);
            tried.httpFallback = true;
            const status = resp3.status;
            const isWorking = status >= 200 && status < 400;
            const body = (typeof resp3.data === 'string') ? resp3.data : JSON.stringify(resp3.data);
            const det = detectFromBody(body);
            return { status, isWorking, found: det.found, snippet: det.snippet, note: 'http-fallback' };
          } catch (err3) {
            console.error('HTTP fallback also failed for', raw, 'Error:', err3 && err3.toString());
            return { status: err3.response ? err3.response.status : 'ERR', isWorking: false, found: false, snippet: String(err3) };
          }
        }
        return { status: err2.response ? err2.response.status : 'ERR', isWorking: false, found: false, snippet: String(err2) };
      }
    }

    // Not a cert error — return original error info
    return { status: err.response ? err.response.status : 'ERR', isWorking: false, found: false, snippet: String(err) };
  }
}




const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`CSV-backlink server started on http://localhost:${PORT}`));
