import { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

const PAGE_SIZE = 100;

const styles = `
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Sora:wght@400;600;700;800&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg: #F5F4F0;
    --surface: #FFFFFF;
    --surface2: #F0EEE9;
    --border: #E2DED6;
    --border2: #CCC8BF;
    --text: #1A1814;
    --text2: #6B6760;
    --text3: #A09C97;
    --accent: #2B6CB0;
    --accent-light: #EBF2FB;
    --green: #1A7A4A;
    --green-bg: #E8F5EE;
    --red: #C0392B;
    --red-bg: #FDECEA;
    --amber: #B45309;
    --amber-bg: #FEF3C7;
    --radius: 10px;
  }

  body {
    font-family: 'Sora', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }

  .topbar {
    background: var(--surface);
    border-bottom: 1.5px solid var(--border);
    padding: 13px 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-shrink: 0;
  }
  .logo-wrap { display: flex; align-items: center; gap: 11px; }
  .logo-icon {
    width: 38px; height: 38px; border-radius: 10px;
    background: var(--accent);
    display: flex; align-items: center; justify-content: center;
    font-size: 19px; color: #fff;
  }
  .logo-title { font-size: 15px; font-weight: 800; color: var(--text); letter-spacing: -0.02em; }
  .logo-sub { font-size: 10px; color: var(--text3); letter-spacing: 0.08em; text-transform: uppercase; font-family: 'IBM Plex Mono', monospace; }
  .topbar-actions { display: flex; gap: 8px; }

  .btn-primary {
    display: flex; align-items: center; gap: 6px;
    padding: 8px 16px; background: var(--accent); border: none;
    border-radius: 8px; color: #fff; font-size: 13px; font-weight: 700;
    cursor: pointer; font-family: 'Sora', sans-serif; transition: background 0.15s;
  }
  .btn-primary:hover { background: #245a96; }

  .btn-ghost {
    padding: 8px 14px; background: transparent; border: 1.5px solid var(--border2);
    border-radius: 8px; color: var(--red); font-size: 13px; cursor: pointer;
    font-family: 'Sora', sans-serif; font-weight: 600; transition: border-color 0.15s;
  }
  .btn-ghost:hover { border-color: var(--red); }

  .main { flex: 1; padding: 24px 28px; display: flex; flex-direction: column; gap: 20px; }

  .dropzone {
    border: 2px dashed var(--border2);
    border-radius: 14px; padding: 36px 20px; text-align: center; cursor: pointer;
    background: var(--surface); transition: all 0.15s;
  }
  .dropzone:hover, .dropzone.drag { border-color: var(--accent); background: var(--accent-light); }
  .dropzone-icon {
    display: inline-flex; background: var(--accent-light);
    border-radius: 12px; padding: 12px 16px; font-size: 30px; margin-bottom: 12px;
  }
  .dropzone h3 { font-size: 14px; font-weight: 700; color: var(--text); margin-bottom: 4px; }
  .dropzone p { font-size: 12px; color: var(--text2); margin-bottom: 14px; }
  .btn-browse {
    font-size: 12px; padding: 6px 16px; border: 1.5px solid var(--accent);
    border-radius: 6px; color: var(--accent); background: transparent;
    cursor: pointer; font-family: 'Sora', sans-serif; font-weight: 600;
  }

  .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; }
  .stat-card {
    background: var(--surface); border: 1.5px solid var(--border);
    border-radius: var(--radius); padding: 16px 18px;
  }
  .stat-label {
    font-size: 10px; font-weight: 700; letter-spacing: 0.07em;
    text-transform: uppercase; margin-bottom: 8px;
    font-family: 'IBM Plex Mono', monospace;
  }
  .stat-value {
    font-size: 28px; font-weight: 800;
    font-family: 'IBM Plex Mono', monospace; letter-spacing: -0.03em;
  }

  .filter-bar {
    background: var(--surface); border: 1.5px solid var(--border);
    border-radius: var(--radius); padding: 12px 16px;
    display: flex; flex-wrap: wrap; gap: 8px; align-items: center;
  }
  .ctrl {
    background: var(--bg); border: 1.5px solid var(--border2);
    border-radius: 8px; color: var(--text); font-size: 13px;
    padding: 8px 12px; outline: none; cursor: pointer;
    font-family: 'Sora', sans-serif; transition: border-color 0.15s;
  }
  .ctrl:focus { border-color: var(--accent); }
  .ctrl-count { margin-left: auto; font-size: 12px; color: var(--text3); font-family: 'IBM Plex Mono', monospace; }

  .btn-clear-filter {
    padding: 8px 14px; background: transparent; border: 1.5px solid #fca5a5;
    border-radius: 8px; color: var(--red); font-size: 13px; cursor: pointer;
    font-family: 'Sora', sans-serif; font-weight: 600;
  }

  .table-wrap {
    background: var(--surface); border: 1.5px solid var(--border);
    border-radius: 14px; overflow: hidden;
  }
  .table-scroll { overflow-x: auto; max-height: 480px; overflow-y: auto; }
  table { width: 100%; border-collapse: collapse; table-layout: fixed; }
  thead tr {
    background: var(--bg); position: sticky; top: 0; z-index: 1;
    border-bottom: 1.5px solid var(--border);
  }
  th {
    padding: 11px 16px; font-size: 10px; font-weight: 700;
    color: var(--text3); letter-spacing: 0.08em; text-transform: uppercase;
    font-family: 'IBM Plex Mono', monospace; cursor: pointer; user-select: none;
    white-space: nowrap;
  }
  th:hover { color: var(--accent); }
  td { padding: 10px 16px; font-size: 13px; border-bottom: 1px solid var(--border); }
  tr:last-child td { border-bottom: none; }
  tbody tr { transition: background 0.1s; }
  tbody tr:hover { background: var(--bg); }

  .badge {
    display: inline-block; font-size: 11px; font-weight: 700;
    padding: 3px 10px; border-radius: 999px; letter-spacing: 0.05em;
    font-family: 'IBM Plex Mono', monospace;
  }
  .badge-in  { background: var(--green-bg); color: var(--green); }
  .badge-out { background: var(--red-bg);   color: var(--red);   }
  .badge-brk { background: var(--amber-bg); color: var(--amber); }

  .pagination {
    display: flex; align-items: center; justify-content: space-between;
    padding: 12px 18px; border-top: 1.5px solid var(--border);
  }
  .btn-page {
    padding: 6px 14px; background: transparent; border: 1.5px solid var(--border2);
    border-radius: 7px; font-size: 13px; cursor: pointer;
    font-family: 'Sora', sans-serif; font-weight: 600; color: var(--accent);
    transition: all 0.12s;
  }
  .btn-page:disabled { color: var(--border2); cursor: default; }
  .btn-page:not(:disabled):hover { border-color: var(--accent); background: var(--accent-light); }
  .page-info { font-size: 12px; color: var(--text3); font-family: 'IBM Plex Mono', monospace; }

  .modal-overlay {
    position: fixed; inset: 0; background: rgba(20,18,15,0.45);
    display: flex; align-items: center; justify-content: center; z-index: 1000;
  }
  .modal {
    background: var(--surface); border-radius: 18px; padding: 32px 28px;
    width: 400px; border: 1.5px solid var(--border);
    box-shadow: 0 20px 60px rgba(0,0,0,0.15);
  }
  .modal h2 { font-size: 18px; font-weight: 800; color: var(--text); margin-bottom: 6px; }
  .modal-sub { font-size: 13px; color: var(--text2); margin-bottom: 20px; }

  .year-option {
    display: flex; align-items: center; gap: 12px;
    padding: 11px 14px; border: 1.5px solid var(--border2);
    border-radius: 10px; cursor: pointer; margin-bottom: 8px; transition: all 0.12s;
  }
  .year-option.selected { border-color: var(--accent); background: var(--accent-light); }
  .year-option label { font-size: 14px; font-weight: 700; flex: 1; cursor: pointer; }
  .year-count { font-size: 12px; color: var(--text3); font-family: 'IBM Plex Mono', monospace; }

  .modal-actions { display: flex; gap: 10px; margin-top: 20px; }
  .btn-cancel {
    flex: 1; padding: 11px 0; border-radius: 10px; border: 1.5px solid var(--border2);
    background: transparent; color: var(--text2); font-size: 14px; cursor: pointer;
    font-family: 'Sora', sans-serif; font-weight: 600;
  }
  .btn-export {
    flex: 2; padding: 11px 0; border-radius: 10px; border: none;
    background: var(--accent); color: #fff; font-size: 14px; font-weight: 700;
    cursor: pointer; font-family: 'Sora', sans-serif;
  }
  .btn-export:hover { background: #245a96; }

  /* Date Range Picker */
  .drp-wrap { position: relative; }
  .drp-btn {
    display: flex; align-items: center; gap: 8px;
    padding: 8px 14px; border-radius: 8px;
    border: 1.5px solid var(--border2); background: var(--bg);
    color: var(--text2); font-size: 13px; cursor: pointer;
    font-family: 'Sora', sans-serif; white-space: nowrap; transition: border-color 0.15s;
  }
  .drp-btn.active { border-color: var(--accent); color: var(--text); }

  .drp-dropdown {
    position: absolute; top: calc(100% + 8px); left: 0; z-index: 999;
    background: var(--surface); border: 1.5px solid var(--border2);
    border-radius: 12px; padding: 16px; width: 280px;
    box-shadow: 0 12px 40px rgba(0,0,0,0.12);
  }
  .cal-nav { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
  .cal-nav button { background: none; border: none; color: var(--text2); cursor: pointer; font-size: 17px; padding: 0 4px; }
  .cal-nav span { font-size: 13px; font-weight: 700; color: var(--text); }
  .cal-days-header { display: grid; grid-template-columns: repeat(7,1fr); margin-bottom: 4px; }
  .cal-days-header span { text-align: center; font-size: 10px; color: var(--text3); padding: 2px 0; font-weight: 700; font-family: 'IBM Plex Mono', monospace; }
  .cal-grid { display: grid; grid-template-columns: repeat(7,1fr); gap: 2px; }
  .cal-day {
    text-align: center; padding: 5px 0; border-radius: 6px;
    font-size: 12px; cursor: pointer; transition: all 0.1s; color: var(--text);
  }
  .cal-day:hover { background: var(--accent-light); color: var(--accent); }
  .cal-day.endpoint { background: var(--accent); color: #fff; font-weight: 700; }
  .cal-day.in-range { background: var(--accent-light); color: var(--accent); }
  .cal-hint { font-size: 10px; color: var(--text3); text-align: center; margin-top: 10px; }

  .badge-filtered {
    background: rgba(43,108,176,0.15); border-radius: 4px;
    font-size: 10px; padding: 1px 5px; color: var(--accent);
  }
  .empty-cell { padding: 50px 0; text-align: center; color: var(--text3); font-size: 14px; }
`;

/* ─── Helpers ─────────────────────────────────────────────────────────── */

function to12h(timeStr) {
  if (!timeStr) return "";
  const [hStr, mStr, sStr] = timeStr.split(":");
  let h = parseInt(hStr, 10);
  const ampm = h >= 12 ? "PM" : "AM";
  h = h % 12 || 12;
  return `${String(h).padStart(2, "0")}:${mStr}:${sStr} ${ampm}`;
}

function parseDat(text, sourceName) {
  return text
    .trim()
    .split("\n")
    .map((line) => {
      const parts = line.trim().split(/\s+/);
      if (parts.length < 5) return null;
      const [userId, date, time, deviceId, statusCode] = parts;
      return {
        userId,
        date,
        time,
        datetime: `${date} ${time}`,
        deviceId,
        statusCode: Number(statusCode),
        source: sourceName,
      };
    })
    .filter(Boolean);
}

function readFileText(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target.result);
    r.onerror = () => rej(new Error("Failed to read file"));
    r.readAsText(file);
  });
}

/* ─── Date Range Picker ───────────────────────────────────────────────── */

function DateRangePicker({ from, to, onChange }) {
  const [open, setOpen] = useState(false);
  const [hover, setHover] = useState(null);
  const [vy, setVy] = useState(() => (from ? new Date(from) : new Date()).getFullYear());
  const [vm, setVm] = useState(() => (from ? new Date(from) : new Date()).getMonth());

  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const DAYS   = ["Su","Mo","Tu","We","Th","Fr","Sa"];

  function prevM() { if (vm === 0) { setVm(11); setVy(y => y - 1); } else setVm(m => m - 1); }
  function nextM() { if (vm === 11) { setVm(0); setVy(y => y + 1); } else setVm(m => m + 1); }
  function dim(y, m) { return new Date(y, m + 1, 0).getDate(); }
  function fd(y, m)  { return new Date(y, m, 1).getDay(); }
  function isoStr(y, m, d) {
    return `${y}-${String(m + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }

  function pick(iso) {
    if (!from || (from && to)) { onChange(iso, ""); }
    else {
      if (iso < from) onChange(iso, from);
      else onChange(from, iso);
      setOpen(false);
    }
  }

  function inRange(iso) {
    const end = to || hover;
    if (!from || !end) return false;
    const lo = from < end ? from : end;
    const hi = from < end ? end   : from;
    return iso > lo && iso < hi;
  }

  const total = dim(vy, vm);
  const start = fd(vy, vm);
  const cells = [];
  for (let i = 0; i < start; i++) cells.push(null);
  for (let d = 1; d <= total; d++) cells.push(d);

  const label = from && to ? `${from} → ${to}` : from ? `${from} → …` : "Date range";

  return (
    <div className="drp-wrap">
      <button className={`drp-btn${from || to ? " active" : ""}`} onClick={() => setOpen(o => !o)}>
        <span>📅</span>
        {label}
        {(from || to) && (
          <span
            onClick={e => { e.stopPropagation(); onChange("", ""); }}
            style={{ opacity: 0.5, fontSize: 11, cursor: "pointer" }}
          >✕</span>
        )}
      </button>

      {open && (
        <div className="drp-dropdown">
          <div className="cal-nav">
            <button onClick={prevM}>‹</button>
            <span>{MONTHS[vm]} {vy}</span>
            <button onClick={nextM}>›</button>
          </div>
          <div className="cal-days-header">
            {DAYS.map(d => <span key={d}>{d}</span>)}
          </div>
          <div className="cal-grid">
            {cells.map((day, i) => {
              if (!day) return <div key={i} />;
              const iso = isoStr(vy, vm, day);
              const ep  = iso === from || iso === to;
              return (
                <div
                  key={i}
                  className={`cal-day${ep ? " endpoint" : inRange(iso) ? " in-range" : ""}`}
                  onClick={() => pick(iso)}
                  onMouseEnter={() => { if (from && !to) setHover(iso); }}
                  onMouseLeave={() => setHover(null)}
                >
                  {day}
                </div>
              );
            })}
          </div>
          {from && !to && <p className="cal-hint">Click a second date to finish the range</p>}
        </div>
      )}
    </div>
  );
}

/* ─── Main Component ──────────────────────────────────────────────────── */

export default function AttendanceLogViewer() {
  const [records, setRecords]         = useState([]);
  const [dragging, setDragging]       = useState(false);
  const [searchUser, setSearchUser]   = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [filterDevice, setFilterDevice] = useState("");
  const [filterFrom, setFilterFrom]   = useState("");
  const [filterTo, setFilterTo]       = useState("");
  const [sortCol, setSortCol]         = useState("datetime");
  const [sortDir, setSortDir]         = useState(-1);
  const [page, setPage]               = useState(1);
  const [showModal, setShowModal]     = useState(false);
  const [selYear, setSelYear]         = useState("all");
  const fileRef = useRef();

  /* Load files */
  const loadFiles = useCallback(async (files) => {
    const incoming = [];
    for (const f of files) {
      const text  = await readFileText(f);
      const label = f.name.replace(/\.dat$/i, "");
      incoming.push(...parseDat(text, label));
    }
    setRecords(prev => [...prev, ...incoming]);
    setPage(1);
  }, []);

  /* Derived data */
  const devices = useMemo(
    () => [...new Set(records.map(r => r.deviceId))].sort(),
    [records]
  );

  const stats = useMemo(() => ({
    total:     records.length,
    users:     new Set(records.map(r => r.userId)).size,
    checkins:  records.filter(r => r.statusCode === 1).length,
    checkouts: records.filter(r => r.statusCode === 0).length,
  }), [records]);

  const filtered = useMemo(() => {
    let rows = records.filter(r => {
      if (searchUser   && !r.userId.toLowerCase().includes(searchUser.toLowerCase())) return false;
      if (filterStatus !== "" && r.statusCode !== Number(filterStatus))               return false;
      if (filterDevice && r.deviceId !== filterDevice)                                return false;
      if (filterFrom   && r.date < filterFrom)                                        return false;
      if (filterTo     && r.date > filterTo)                                          return false;
      return true;
    });
    rows.sort((a, b) => {
      const va = a[sortCol] ?? "", vb = b[sortCol] ?? "";
      return va < vb ? -sortDir : va > vb ? sortDir : 0;
    });
    return rows;
  }, [records, searchUser, filterStatus, filterDevice, filterFrom, filterTo, sortCol, sortDir]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const slice      = filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  /* Sort handler */
  function handleSort(col) {
    if (sortCol === col) setSortDir(d => d * -1);
    else { setSortCol(col); setSortDir(col === "datetime" ? -1 : 1); }
    setPage(1);
  }

  /* Clear helpers */
  function clearFilters() {
    setSearchUser(""); setFilterStatus(""); setFilterDevice("");
    setFilterFrom(""); setFilterTo(""); setPage(1);
  }
  function clearAll() { setRecords([]); clearFilters(); }

  const hasFilters = filterStatus || searchUser || filterDevice || filterFrom || filterTo;

  /* Years available for export modal */
  const availYears = useMemo(
    () => [...new Set(records.map(r => r.date.slice(0, 4)))].sort((a, b) => b - a),
    [records]
  );

  /* Export to Excel */
  function doExport() {
    setShowModal(false);
    const base   = selYear === "all" ? filtered : filtered.filter(r => r.date.startsWith(selYear));
    const sorted = [...base].sort((a, b) => (b.datetime < a.datetime ? -1 : 1));
    const rows   = sorted.map(r => ({
      Emp_No:        Number(r.userId),
      Attend_Date:   r.date,
      Attend_Time:   to12h(r.time),
      Attend_Status: r.statusCode,
    }));
    const ws = XLSX.utils.json_to_sheet(rows, {
      header: ["Emp_No", "Attend_Date", "Attend_Time", "Attend_Status"],
    });
    ws["!cols"] = [{ wch: 12 }, { wch: 14 }, { wch: 18 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance");
    const today  = new Date();
    const dateStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
    XLSX.writeFile(wb, `Attendance_${selYear === "all" ? "All" : selYear}_${dateStr}.xlsx`);
  }

  /* Sort icon */
  const SortIcon = ({ col }) => (
    <span style={{ opacity: sortCol === col ? 1 : 0.3, marginLeft: 4, fontSize: 10 }}>
      {sortCol === col ? (sortDir === 1 ? "▲" : "▼") : "⇅"}
    </span>
  );

  const statCards = [
    { label: "Total Records", value: stats.total,     color: "#2B6CB0", borderColor: "#2B6CB0" },
    { label: "Unique Users",  value: stats.users,     color: "#0F766E", borderColor: "#0F766E" },
    { label: "IN",      value: stats.checkins,  color: "#1A7A4A", borderColor: "#1A7A4A" },
    { label: "OUT",     value: stats.checkouts, color: "#C0392B", borderColor: "#C0392B" },
  ];

  return (
    <>
      {/* Inject styles */}
      <style>{styles}</style>

      <div style={{ fontFamily: "'Sora', sans-serif", background: "#F5F4F0", minHeight: "100vh", display: "flex", flexDirection: "column" }}>

        {/* ── Year Export Modal ── */}
        {showModal && (
          <div className="modal-overlay" onClick={() => setShowModal(false)}>
            <div className="modal" onClick={e => e.stopPropagation()}>
              <div style={{ fontSize: 34, marginBottom: 12 }}>📅</div>
              <h2>Export by Year</h2>
              <p className="modal-sub">Pick which year's records to export</p>
              {[
                { label: "All years", value: "all", count: filtered.length },
                ...availYears.map(yr => ({
                  label: yr,
                  value: yr,
                  count: filtered.filter(r => r.date.startsWith(yr)).length,
                })),
              ].map(opt => (
                <div
                  key={opt.value}
                  className={`year-option${selYear === opt.value ? " selected" : ""}`}
                  onClick={() => setSelYear(opt.value)}
                >
                  <input
                    type="radio" name="yr" value={opt.value}
                    checked={selYear === opt.value}
                    onChange={() => setSelYear(opt.value)}
                    style={{ accentColor: "#2B6CB0" }}
                  />
                  <label>{opt.label}</label>
                  <span className="year-count">{opt.count.toLocaleString()} rows</span>
                </div>
              ))}
              <div className="modal-actions">
                <button className="btn-cancel" onClick={() => setShowModal(false)}>Cancel</button>
                <button className="btn-export" onClick={doExport}>Export Excel ↓</button>
              </div>
            </div>
          </div>
        )}

        {/* ── Top Bar ── */}
        <div className="topbar">
          <div className="logo-wrap">
            <div className="logo-icon">⏱</div>
            <div>
              <div className="logo-title">Attendance Reports</div>
              <div className="logo-sub">Biometric Viewer</div>
            </div>
          </div>
          {records.length > 0 && (
            <div className="topbar-actions">
              <button
                className="btn-primary"
                onClick={() => { setSelYear("all"); setShowModal(true); }}
              >
                ↓ Export Excel
                {hasFilters && <span className="badge-filtered">filtered</span>}
              </button>
              <button className="btn-ghost" onClick={clearAll}>✕ Clear</button>
            </div>
          )}
        </div>

        {/* ── Main Content ── */}
        <div className="main">

          {/* Drop Zone */}
          <div
            className={`dropzone${dragging ? " drag" : ""}`}
            onDragOver={e => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={async e => { e.preventDefault(); setDragging(false); await loadFiles([...e.dataTransfer.files]); }}
            onClick={() => fileRef.current.click()}
          >
            <div className="dropzone-icon">📂</div>
            <h3>Drop .dat files here</h3>
            <p>Supports multiple files — records merged automatically</p>
            <button className="btn-browse" onClick={e => e.stopPropagation()}>Browse files</button>
            <input
              ref={fileRef} type="file" multiple accept=".dat,.txt"
              style={{ display: "none" }}
              onChange={async e => { await loadFiles([...e.target.files]); e.target.value = ""; }}
            />
          </div>

          {records.length > 0 && (
            <>
              {/* Stats */}
              <div className="stats-grid">
                {statCards.map(s => (
                  <div
                    key={s.label}
                    className="stat-card"
                    style={{ borderLeft: `4px solid ${s.borderColor}` }}
                  >
                    <div className="stat-label" style={{ color: s.color }}>{s.label}</div>
                    <div className="stat-value" style={{ color: s.color }}>{s.value.toLocaleString()}</div>
                  </div>
                ))}
              </div>

              {/* Filter Bar */}
              <div className="filter-bar">
                <input
                  type="text" placeholder="Search emp ID…"
                  value={searchUser}
                  onChange={e => { setSearchUser(e.target.value); setPage(1); }}
                  className="ctrl"
                  style={{ width: 150 }}
                />
                <select
                  value={filterStatus}
                  onChange={e => { setFilterStatus(e.target.value); setPage(1); }}
                  className="ctrl"
                >
                  <option value="">All events</option>
                  <option value="1">IN</option>
                  <option value="0">OUT</option>
                  <option value="5">Break</option>
                </select>
                <select
                  value={filterDevice}
                  onChange={e => { setFilterDevice(e.target.value); setPage(1); }}
                  className="ctrl"
                >
                  <option value="">All devices</option>
                  {devices.map(d => <option key={d} value={d}>Device {d}</option>)}
                </select>
                <DateRangePicker
                  from={filterFrom}
                  to={filterTo}
                  onChange={(f, t) => { setFilterFrom(f); setFilterTo(t); setPage(1); }}
                />
                {hasFilters && (
                  <button className="btn-clear-filter" onClick={clearFilters}>✕ Clear</button>
                )}
                <span className="ctrl-count">{filtered.length.toLocaleString()} records</span>
              </div>

              {/* Table */}
              <div className="table-wrap">
                <div className="table-scroll">
                  <table>
                    <thead>
                      <tr>
                        {[
                          { col: "userId",   label: "Emp ID",    w: "12%", align: "right"  },
                          { col: "datetime", label: "Date & Time", w: "35%", align: "center" },
                          { col: "deviceId", label: "Device ID",  w: "23%", align: "center" },
                          { col: null,       label: "Status",     w: "15%", align: "center" },
                          { col: "source",   label: "Source",     w: "15%", align: "center" },
                        ].map(({ col, label, w, align }) => (
                          <th
                            key={label}
                            onClick={col ? () => handleSort(col) : undefined}
                            style={{ width: w, textAlign: align }}
                          >
                            {label} {col && <SortIcon col={col} />}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {slice.length === 0 ? (
                        <tr>
                          <td colSpan={5} className="empty-cell">No records match your filters.</td>
                        </tr>
                      ) : (
                        slice.map((r, i) => (
                          <tr key={i}>
                            <td style={{ textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: "#1A1814" }}>
                              {r.userId}
                            </td>
                            <td style={{ textAlign: "center", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#4B5563" }}>
                              {r.date}{" "}
                              <span style={{ color: "#2B6CB0", fontWeight: 600 }}>{to12h(r.time)}</span>
                            </td>
                            <td style={{ textAlign: "center", color: "#6B7280", fontSize: 13 }}>{r.deviceId}</td>
                            <td style={{ textAlign: "center" }}>
                              <span className={`badge ${r.statusCode === 1 ? "badge-in" : r.statusCode === 0 ? "badge-out" : "badge-brk"}`}>
                                {r.statusCode === 1 ? "IN" : r.statusCode === 0 ? "OUT" : "BRK"}
                              </span>
                            </td>
                            <td style={{ textAlign: "center", fontSize: 12, color: "#9CA3AF" }}>{r.source}</td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Pagination */}
                {filtered.length > PAGE_SIZE && (
                  <div className="pagination">
                    <button
                      className="btn-page"
                      disabled={page <= 1}
                      onClick={() => setPage(p => p - 1)}
                    >← Prev</button>
                    <span className="page-info">
                      {page} / {totalPages} &nbsp;·&nbsp;
                      {((page - 1) * PAGE_SIZE + 1).toLocaleString()}–{Math.min(page * PAGE_SIZE, filtered.length).toLocaleString()}
                    </span>
                    <button
                      className="btn-page"
                      disabled={page >= totalPages}
                      onClick={() => setPage(p => p + 1)}
                    >Next →</button>
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </div>
    </>
  );
}