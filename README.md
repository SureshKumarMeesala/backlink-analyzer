Backlink Analyzer

A lightweight, self-hosted backlink checker — upload an Excel/CSV of target pages, crawl them concurrently, detect whether each page contains a backlink to your site, get live progress updates on a Next.js dashboard (SSE), and download the results as an Excel file.

Built with:

Frontend: Next.js + Tailwind CSS

Backend API: Express.js

Scraping: axios + cheerio

Excel handling: ExcelJS

File upload: multer

Real-time updates: Server-Sent Events (SSE)

Links

Project repo / developer: https://github.com/SureshKumarMeesala

Company: https://adkode.com/

Features

Upload .xlsx or .csv (URLs in Column A).

Crawl thousands of URLs with configurable concurrency and delay.

Detect backlink presence by exact href, domain match or plain text match.

Real-time progress and per-row updates via SSE.

Export results as results-<timestamp>.xlsx.

Resume-friendly: backend saves results to output file incrementally.

Simple, mobile-friendly UI built with Tailwind.

Repo layout (recommended)
/
├─ server/
│  ├─ index.js              # Express backend (API + SSE + crawler)
│  ├─ uploads/              # uploads and output files
│  └─ package.json
├─ web/
│  ├─ app/                  # Next.js app (React pages/components)
│  ├─ package.json
│  └─ tailwind.config.js
├─ README.md

Quick start (local)

You can run frontend and backend locally on your machine.

Prereqs

Node.js 18+ (or LTS)

npm or yarn

(Optional) curl, unzip, 7z for testing

1) Backend (Express)
cd server
npm install


Create server/uploads directory if not already present (the backend will create it automatically on start but you can create it manually):

mkdir -p server/uploads


Start backend (development):

# development (auto-restart with nodemon, optional)
npx nodemon index.js

# or plain node
node index.js


Default backend URL: http://localhost:4000

Key server config (in server/index.js):

DEFAULT_BACKLINK — default backlink to check

CONCURRENCY — how many parallel workers

DELAY_MS — delay between requests per worker

FETCH_TIMEOUT — axios timeout per request

You can edit those constants or environment-ify them as needed.

2) Frontend (Next.js)
cd web
npm install
npm run dev


Default frontend URL: http://localhost:3000

Open the page, choose an .xlsx or .csv file with URLs in column A, enter the backlink (or use default), and upload. Watch the dashboard update in real-time.

API Endpoints (backend)

server/index.js exposes:

GET /api/events — SSE endpoint. Client listens for events: start, rowStart, rowDone, done, error.

POST /api/upload — accepts multipart file and backlink. Returns { ok: true, total, download } then starts processing in background.

Form fields: file (binary), backlink (string)

GET /api/download?file=<filename> — download results file from uploads/

File input format

.xlsx or .csv

URLs expected in column A (first column). If the first row looks like a header (non-URL), the uploader treats it as header and starts from row 2.

CSV delimiter detection uses common separators (, ; ; \t |).

How processing works (overview)

Upload file -> backend reads workbook (ExcelJS or CSV fallback).

Backend enumerates URLs and produces results-<timestamp>.xlsx in uploads/.

Backend spins up worker queue with CONCURRENCY parallel workers. Each worker:

Normalizes URL (adds http:// if missing).

Fetches page with axios.

Parses HTML with cheerio, searches for backlinks (exact href=, domain match, or plain text).

Backend writes per-row results back into the Excel workbook and saves the file periodically (every ~20 rows).

Backend sends SSE messages as each row starts/finishes for real-time UI updates.

Configuration options (server/index.js)

Edit these at the top of server/index.js if you want to tune:

const DEFAULT_BACKLINK = 'https://example.com';
const CONCURRENCY = 6;      // number of parallel workers
const DELAY_MS = 200;       // delay between requests (ms) per worker
const FETCH_TIMEOUT = 20000;// axios timeout (ms)


Notes:

Increase DELAY_MS to be polite to target servers and avoid IP bans.

Reduce CONCURRENCY on constrained hosts or shared hosting.

If host blocks outbound requests, you must run the worker on a machine that allows external HTTP.

Running workers / background processing

Laravel-style queue workers are not used here — the Express server handles processing itself in background (spawned async task). However if you adapt this architecture (or use a database-backed queue), you’ll need persistent workers.

On hosting that restricts long-running processes (shared hosting), use a cron to curl an endpoint that processes a small batch, or run the server on a VPS that allows background services.

Troubleshooting
SSE shows Disconnected or SSE error

Backend may have closed the connection. Check server logs. Ensure CORS is allowed from your frontend origin (app.use(cors({ origin: 'http://localhost:3000' }))).

Make sure the browser isn't blocking EventSource due to Mixed Content (https front + http backend). Use HTTPS for both or use same protocol.

Links show ERR: unable to get local issuer certificate

This indicates TLS verification failed on the server. For debugging you can disable SSL verification in axios (NOT recommended for production). Better solution: ensure the server has up-to-date CA certificates or run your worker from a host with proper CA store.

TypeError: Invalid URL in logs

Some CSV rows contain fragmented/invalid data. Use the uploader's extractUrlFromRow logic — it attempts to normalize and skip obviously bad values. Clean your CSV if many invalid rows exist.

Too many http or https rows (CSV splitting)

Some CSVs have broken formatting that results in cells like http only. The extractUrlFromRow() fallback tries to join row cells — but best is to clean the input.

Security & rate limiting

This tool performs automated requests. Respect robots.txt and target site policies.

Do not use for abusive scraping or to attempt unauthorized access.

Consider rate-limiting and throttling:

CONCURRENCY: lower value if rate-limited.

DELAY_MS: increase if targets block your IP.

Add an optional User-Agent header to represent the tool and include contact info.

Exporting results

The backend saves incremental results to server/uploads/results-<timestamp>.xlsx. Download link is returned and also sent via SSE done event. The file contains columns:

Row number

URL

HTTP Status

Working? (Yes/No)

Backlink Present? (Yes/No)

Found Snippet / Note

Last Checked (timestamp)

You can also export visible results from the frontend (client-side) as a separate .xlsx (uses SheetJS).

Deployment notes

For production, run the backend behind a process manager (pm2, systemd) and serve frontend via Vercel, Netlify, or same host.

Use HTTPS for both frontend and backend. If using different domains, configure CORS accordingly.

Ensure the server can make outbound HTTP(s) requests to external hosts (curl test: curl -I https://example.com).

If processing many URLs, consider disk space and memory. Save out files periodically and rotate older uploads.

Example pm2 usage:

# from server/
npm install -g pm2
pm2 start index.js --name backlink-server
pm2 logs backlink-server

Example: Upload a CSV via curl
curl -F "file=@urls.csv" -F "backlink=https://example.com" http://localhost:4000/api/upload


Server responds immediately with a JSON that includes total and a download path; processing continues in background and SSE will push updates.

Extending / Customizing

Ideas and paths you can extend:

Add anchor text extraction and store it in a DB.

Add a DB (Postgres/MySQL) to persist results and build historical reports.

Add authentication for multi-user dashboards.

Add scheduled scans (cron / job scheduler).

Add more robust parsing (render JavaScript with a headless browser like Puppeteer for JS-heavy pages).

Add quality scoring and link spam filters.

Development tips

Use DELAY_MS and CONCURRENCY to tune throughput vs reliability.

Keep an eye on server/uploads and remove old files routinely.

Tail logs while testing: tail -f server/logs.txt (if you log to a file) or watch pm2 logs.

If you see mass DNS errors, your host may be blocking outbound DNS or cURL.

License

This project is free to adapt. Add your preferred license file (e.g. MIT) in the repo root:

LICENSE:

MIT License
...

Credits & Contact

Built by Adkode — https://adkode.com/

Developer: Suresh Kumar M — https://github.com/SureshKumarMeesala

Repo: https://github.com/SureshKumarMeesala
