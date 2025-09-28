"use client";
import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";

export default function Home() {
  const PAGE_SIZE = 100; // rows per page (change if needed)

  const [file, setFile] = useState(null);
  const [backlink, setBacklink] = useState("https://bonfireinstituteofdesign.com");
  const [logs, setLogs] = useState([]);
  const [rows, setRows] = useState([]); // {row, url, status, isWorking, found, snippet, lastChecked, note}
  const [total, setTotal] = useState(0);
  const [downloadUrl, setDownloadUrl] = useState(null); // server-provided download path
  const evtSourceRef = useRef(null);
  const [sseConnected, setSseConnected] = useState(false);
  const [showPending, setShowPending] = useState(false); // default: hide pending

  // pagination state
  const [currentPage, setCurrentPage] = useState(1);

  useEffect(() => {
    let reconnectTimer = null;
    function connect() {
      const src = new EventSource("http://localhost:4000/api/events");
      evtSourceRef.current = src;

      src.onopen = () => {
        setLogs((l) => [...l, "SSE connected"]);
        setSseConnected(true);
      };

      src.onmessage = (e) => {
        try {
          const d = JSON.parse(e.data);
          handleEvent(d);
        } catch (err) {
          console.error("SSE parse", err, e.data);
          setLogs((l) => [...l, "SSE parse error"]);
        }
      };

      src.onerror = (err) => {
        console.error("SSE error", err);
        setLogs((l) => [...l, "SSE error — retrying in 3s"]);
        setSseConnected(false);
        try { src.close(); } catch (e) {}
        clearTimeout(reconnectTimer);
        reconnectTimer = setTimeout(connect, 3000);
      };
    }

    connect();
    return () => {
      clearTimeout(reconnectTimer);
      try { evtSourceRef.current && evtSourceRef.current.close(); } catch (e) {}
    };
  }, []);

  function handleEvent(d) {
    if (!d) return;
    if (d.type === "start") {
      setLogs((l) => [...l, `Started: ${d.total} URLs — looking for ${d.backlink || backlink}`]);
      setTotal(d.total || 0);
      setDownloadUrl(null);
      // reset to first page on new run
      setCurrentPage(1);
    } else if (d.type === "rowStart") {
      setRows((rs) => {
        const exists = rs.some((r) => r.row === d.row);
        if (exists) return rs;
        return [...rs, { row: d.row, url: d.url, status: "Pending", isWorking: "", found: "", snippet: "", lastChecked: null, note: "" }].sort((a,b)=>a.row-b.row);
      });
      setLogs((l) => [...l, `Row ${d.row} start: ${d.url || ""}`]);
    } else if (d.type === "rowDone") {
      const r = d.result || {};
      const note = r.note || "";
      setLogs((l) => [...l, `Row ${d.row} done — ${r.status} — backlink:${r.found} ${note ? "(" + note + ")" : ""}`]);
      setRows((rs) => rs.map(row => row.row === d.row ? {
        ...row,
        url: d.url || row.url,
        status: r.status,
        isWorking: r.isWorking ? "Yes" : "No",
        found: r.found ? "Yes" : "No",
        snippet: r.snippet || "",
        lastChecked: new Date(),
        note
      } : row));
    } else if (d.type === "badRows") {
      setLogs((l) => [...l, `Bad rows detected: ${JSON.stringify(d.badRows || []).slice(0,300)}`]);
    } else if (d.type === "done") {
      setLogs((l) => [...l, "Done — download: " + d.download]);
      setDownloadUrl(d.download);
    } else if (d.type === "error") {
      setLogs((l) => [...l, "Error: " + d.message]);
    }
  }

  function handleFileSelect(e) {
    const f = e.target.files[0];
    setFile(f || null);
    if (!f) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target.result;
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const arr = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      const dataRows = arr.slice(1);
      const initial = dataRows.map((r, i) => {
        const excelRow = i + 2;
        return { row: excelRow, url: r[0] ? String(r[0]).trim() : "", status: "Pending", isWorking: "", found: "", snippet: "", lastChecked: null, note: "" };
      }).filter(r => r.url);
      setRows(initial);
      setTotal(initial.length);
      setCurrentPage(1);
    };
    reader.readAsArrayBuffer(f);
  }

  async function onSubmit(e) {
    e.preventDefault();
    if (!file) return alert("Choose an Excel or CSV file with URLs in column A");
    const form = new FormData();
    form.append("file", file);
    form.append("backlink", backlink);
    setLogs((l) => [...l, "Uploading file..."]);
    try {
      const r = await fetch("http://localhost:4000/api/upload", { method: "POST", body: form });
      const j = await r.json();
      if (j.ok) {
        setLogs((l) => [...l, `Upload accepted — ${j.total || 0} URLs`]);
        setDownloadUrl(j.download || null);
        setTotal(j.total || 0);
        setCurrentPage(1);
      } else {
        setLogs((l) => [...l, "Upload error: " + (j.error || JSON.stringify(j))]);
      }
    } catch (err) {
      setLogs((l) => [...l, "Upload failed: " + String(err)]);
    }
  }





  // helper: fetch latest summary from server
async function fetchSummaryAndPage(page = currentPage, pageSize = PAGE_SIZE) {
  try {
    const s = await fetch(`http://localhost:4000/api/summary`);
    const sj = await s.json();
    if (sj && sj.ok) {
      setTotal(sj.total || 0);
      // update summary counts (optional: show in UI)
      // you can also store sj.processed, sj.ok, sj.errors, sj.backlinks to display if preferred
    }

    // compute offset for /api/rows
    const offset = (page - 1) * pageSize;
    const r = await fetch(`http://localhost:4000/api/rows?offset=${offset}&limit=${pageSize}`);
    const rj = await r.json();
    if (rj && rj.ok) {
      // replace only the visible page rows (keeps other pages untouched)
      // Map returned rows into your row shape
      const returned = (rj.rows || []).map(rr => ({
        row: rr.row,
        url: rr.url,
        status: rr.status,
        isWorking: rr.isWorking,
        found: rr.found,
        snippet: rr.snippet,
        lastChecked: rr.lastChecked ? new Date(rr.lastChecked) : null,
        note: ''
      }));
      // Update rows array: merge returned into existing rows or replace displayed page
      setRows(prev => {
        // build a map for quick replacement
        const map = new Map(prev.map(p => [p.row, p]));
        returned.forEach(rw => map.set(rw.row, rw));
        // produce sorted array
        const merged = Array.from(map.values()).sort((a,b)=>a.row-b.row);
        return merged;
      });
    }
  } catch (err) {
    console.error('fetchSummaryAndPage error', err);
  }
}



useEffect(() => {
  // initial fetch
  fetchSummaryAndPage(currentPage, PAGE_SIZE);

  const iv = setInterval(() => {
    fetchSummaryAndPage(currentPage, PAGE_SIZE);
  }, 3000); // poll every 3s

  return () => clearInterval(iv);
}, [currentPage]); // re-run when page changes





  // Derived stats (from rows)
  const processedRows = rows.filter(r => r.lastChecked);
  const processedCount = processedRows.length;
  const okCount = processedRows.filter(r => {
    const s = r.status;
    return typeof s === "number" ? (s >= 200 && s < 400) : (/^2\d\d$/.test(String(s)));
  }).length;
  const errorCount = processedRows.filter(r => {
    const s = String(r.status || "");
    return s === "ERR" || /^[45]\d\d$/.test(s) || (r.isWorking === "No" && String(r.status) !== "200");
  }).length;
  const backlinkCount = processedRows.filter(r => r.found === "Yes").length;
  const percent = total ? Math.round((processedCount/total)*100) : 0;

  // Data to display in table = either all rows or only processed rows depending on toggle
  const visibleRows = useMemo(() => (showPending ? rows : rows.filter(r => r.lastChecked)), [rows, showPending]);

  // Pagination calculations
  const pageCount = Math.max(1, Math.ceil(visibleRows.length / PAGE_SIZE));
  // Ensure currentPage in range
  useEffect(() => { if (currentPage > pageCount) setCurrentPage(pageCount); }, [pageCount, currentPage]);

  const pageStartIndex = (currentPage - 1) * PAGE_SIZE;
  const pageEndIndex = pageStartIndex + PAGE_SIZE;
  const pageRows = visibleRows.slice(pageStartIndex, pageEndIndex);

  function goToPage(n) {
    const p = Math.min(Math.max(1, Math.floor(n)), pageCount);
    setCurrentPage(p);
    // scroll to top of results table for nicer UX
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  // Export current page to Excel (.xlsx) using SheetJS
  function exportCurrentPageToExcel() {
    if (!pageRows || pageRows.length === 0) {
      alert("No rows on this page to export.");
      return;
    }
    const header = ["Row", "URL", "HTTP", "Working?", "Backlink", "Snippet / Note", "Last Checked"];
    const data = pageRows.map(r => [
      r.row,
      r.url,
      r.status,
      r.isWorking,
      r.found,
      r.snippet || r.note,
      r.lastChecked ? new Date(r.lastChecked).toLocaleString() : ""
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Results");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `backlink_results_page_${currentPage}_${Date.now()}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto space-y-6">
        <div className="bg-white shadow rounded-lg p-6">
          <div className="flex items-center justify-between mb-4">
            <h1 className="text-2xl font-semibold text-black">Backlink Checker — Results</h1>
            <div className="flex items-center space-x-3">
              <span className={`px-3 py-1 rounded-full text-sm ${sseConnected ? "bg-green-100 text-black" : "bg-red-100 text-red-800"}`}>
                {sseConnected ? "Live" : "Disconnected"}
              </span>
            </div>
          </div>

          <form onSubmit={onSubmit} className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
            <div className="md:col-span-1">
              <label className="block text-sm font-medium text-black mb-1">Excel / CSV (Column A)</label>
              <input type="file" accept=".xlsx,.csv" onChange={handleFileSelect} className="block w-full text-sm text-gray-600" />
            </div>
            <div className="md:col-span-1">
              <label className="block text-sm font-medium text-black mb-1">Backlink to check</label>
              <input value={backlink} onChange={(e)=>setBacklink(e.target.value)} className="block w-full border border-gray-200 rounded p-2 text-black" />
            </div>
            <div className="md:col-span-1 flex items-end space-x-2">
              <button type="submit" className="w-full inline-flex justify-center items-center px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700">Upload & Start</button>
            </div>
          </form>

          <div className="mb-4 flex items-center justify-between">
            <div className="flex items-center space-x-4">
              <label className="inline-flex items-center space-x-2 text-sm">
                <input type="checkbox" checked={showPending} onChange={() => { setShowPending(s => !s); setCurrentPage(1); }} className="form-checkbox text-black" />
                <span className="text-black">Show Pending rows</span>
              </label>
              <div className="text-sm text-gray-600">Showing {visibleRows.length} rows</div>
            </div>

            <div className="flex items-center space-x-2">
              {/* Server-provided download (if available) */}
              {downloadUrl && (
                <a href={`http://localhost:4000${downloadUrl}`} className="px-3 py-2 bg-green-600 text-white rounded text-sm hover:bg-green-700">Download server results</a>
              )}
              {/* Export current page to Excel */}
              <button onClick={exportCurrentPageToExcel} className="px-3 py-2 bg-indigo-600 text-white rounded text-sm hover:bg-indigo-700">Export page (.xlsx)</button>
            </div>
          </div>

          {/* summary */}
          <div className="grid grid-cols-1 sm:grid-cols-3 lg:grid-cols-6 gap-4 mb-4">
            <div className="p-3 bg-gray-50 rounded shadow-sm">
              <div className="text-sm text-black">Total</div>
              <div className="text-xl font-bold text-black">{total}</div>
            </div>
            <div className="p-3 bg-gray-50 rounded shadow-sm">
              <div className="text-sm text-black">Processed</div>
              <div className="text-xl font-bold text-black">{processedCount}</div>
            </div>
            <div className="p-3 bg-gray-50 rounded shadow-sm">
              <div className="text-sm text-black">200 / OK</div>
              <div className="text-xl font-bold text-green-600">{okCount}</div>
            </div>
            <div className="p-3 bg-gray-50 rounded shadow-sm">
              <div className="text-sm text-black">Errors</div>
              <div className="text-xl font-bold text-red-600">{errorCount}</div>
            </div>
            <div className="p-3 bg-gray-50 rounded shadow-sm">
              <div className="text-sm text-black">Backlinks found</div>
              <div className="text-xl font-bold text-indigo-600">{backlinkCount}</div>
            </div>
            <div className="p-3 bg-gray-50 rounded shadow-sm">
              <div className="text-sm text-black">Progress</div>
              <div className="text-xl font-bold text-black">{percent}%</div>
            </div>
          </div>

          {/* pagination controls */}
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center space-x-2">
              <button onClick={() => goToPage(1)} disabled={currentPage === 1} className="px-3 py-1 bg-yellow-500 rounded disabled:opacity-50">First</button>
              <button onClick={() => goToPage(currentPage - 1)} disabled={currentPage === 1} className="px-3 py-1 bg-red-500 rounded disabled:opacity-50">Prev</button>
              <span className="px-3 py-1 text-black">Page <strong>{currentPage}</strong> of {pageCount}</span>
              <button onClick={() => goToPage(currentPage + 1)} disabled={currentPage === pageCount} className="px-3 py-1 bg-red-500 rounded disabled:opacity-50">Next</button>
              <button onClick={() => goToPage(pageCount)} disabled={currentPage === pageCount} className="px-3 py-1 bg-yellow-500 rounded disabled:opacity-50">Last</button>
            </div>

            <div className="flex items-center space-x-2">
              <label className="text-sm text-gray-600">Jump to</label>
              <input
                type="number"
                min={1}
                max={pageCount}
                value={currentPage}
                onChange={(e) => {
                  const v = Number(e.target.value) || 1;
                  setCurrentPage(Math.min(Math.max(1, v), pageCount));
                }}
                className="w-20 p-1 border rounded text-sm"
              />
            </div>
          </div>

          {/* results table */}
          <div className="overflow-auto border rounded">
            <table className="min-w-full divide-y divide-gray-200">
              <thead className="bg-gray-50 sticky top-0">
                <tr>
                  <th className="px-4 py-2 text-left text-xs font-medium text-black">Row</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-black">URL</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-black">HTTP</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-black">Working?</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-black">Backlink</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-black">Snippet / Note</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-black">Last Checked</th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-100">
                {pageRows.map(r => (
                  <tr key={r.row} className="hover:bg-gray-50">
                    <td className="px-4 py-2 text-sm text-black">{r.row}</td>
                    <td className="px-4 py-2 text-sm text-black max-w-xs truncate">
                    <a
                        href={r.url}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-blue-600 hover:underline"
                        title={r.url}
                    >
                        {r.url}
                    </a>
                    </td>

                    <td className="px-4 py-2 text-sm text-black">{r.status}</td>
                    <td className="px-4 py-2 text-sm text-black">{r.isWorking}</td>
                    <td className="px-4 py-2 text-sm text-black">{r.found}</td>
                    <td className="px-4 py-2 text-sm max-w-md truncate text-black">{r.snippet || r.note}</td>
                    <td className="px-4 py-2 text-sm text-black">{r.lastChecked ? new Date(r.lastChecked).toLocaleString() : ""}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="mt-6 bg-gray-50 p-3 rounded h-32 overflow-auto">
            <div className="text-sm font-medium text-black mb-2">Logs</div>
            <div className="space-y-1 text-sm text-black">
              {logs.slice().reverse().slice(0, 200).map((l, i) => <div key={i}>{l}</div>)}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
