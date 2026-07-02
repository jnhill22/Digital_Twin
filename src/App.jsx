import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as THREE from "three";
import * as XLSX from "xlsx";
import Papa from "papaparse";

/* ------------------------------------------------------------------ */
/*  FacilityTwin — browser-based facility digital twin                 */
/*  DXF floor plans · equipment registry & BOM · utility connections  */
/*  2D/3D views · GitHub snapshot sync · persistent storage           */
/* ------------------------------------------------------------------ */

const UTILITY_TYPES = {
  hvac:    { label: "HVAC / Air Handling", color: "#45C4B0", short: "HVAC" },
  chw:     { label: "Chilled Water",       color: "#4C9BE8", short: "CHW" },
  hw:      { label: "Hot Water / Steam",   color: "#E06552", short: "STM" },
  power:   { label: "Electrical Power",    color: "#EFCB4F", short: "PWR" },
  compair: { label: "Compressed Air",      color: "#9D8DF1", short: "CDA" },
  process: { label: "Process / Product",   color: "#62D26F", short: "PROC" },
  drain:   { label: "Drain / Waste",       color: "#8C9BA5", short: "DRN" },
};

const EQUIP_TYPES = ["Process", "Packaging", "Material Handling", "Storage Tank", "Pump", "Skid", "Instrument", "Other"];

const uid = () => Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-3);

/* ---------------------------- DXF parser --------------------------- */
/* Minimal ASCII DXF entity parser: LINE, LWPOLYLINE, CIRCLE, ARC,     */
/* TEXT, MTEXT. Y axis is negated so screen coordinates read naturally.*/

function parseDXF(text) {
  const rows = text.split(/\r\n|\r|\n/);
  const pairs = [];
  for (let i = 0; i + 1 < rows.length; i += 2) {
    const code = parseInt(rows[i], 10);
    if (!Number.isNaN(code)) pairs.push([code, rows[i + 1].trim()]);
  }
  let section = "";
  let cur = null;
  const raw = [];
  for (let i = 0; i < pairs.length; i++) {
    const [c, v] = pairs[i];
    if (c === 0 && v === "SECTION") {
      const nx = pairs[i + 1];
      section = nx && nx[0] === 2 ? nx[1] : "";
      continue;
    }
    if (c === 0 && (v === "ENDSEC" || v === "EOF")) { if (cur) raw.push(cur); cur = null; section = ""; continue; }
    if (section !== "ENTITIES") continue;
    if (c === 0) {
      if (cur) raw.push(cur);
      cur = { type: v, verts: [], codes: {} };
      continue;
    }
    if (!cur) continue;
    if (c === 10) { cur.verts.push({ x: parseFloat(v), y: 0 }); }
    else if (c === 20) { const lv = cur.verts[cur.verts.length - 1]; if (lv) lv.y = parseFloat(v); }
    else if (!(c in cur.codes)) cur.codes[c] = v;
  }
  if (cur) raw.push(cur);

  const ents = [];
  for (const e of raw) {
    const c = e.codes;
    if (e.type === "LINE" && e.verts[0] && c[11] !== undefined && c[21] !== undefined) {
      ents.push({ kind: "line", pts: [[e.verts[0].x, e.verts[0].y], [parseFloat(c[11]), parseFloat(c[21])]] });
    } else if ((e.type === "LWPOLYLINE" || e.type === "POLYLINE") && e.verts.length > 1) {
      const closed = (parseInt(c[70] || "0", 10) & 1) === 1;
      ents.push({ kind: "poly", pts: e.verts.map(p => [p.x, p.y]), closed });
    } else if (e.type === "CIRCLE" && e.verts[0] && c[40]) {
      ents.push({ kind: "circle", cx: e.verts[0].x, cy: e.verts[0].y, r: parseFloat(c[40]) });
    } else if (e.type === "ARC" && e.verts[0] && c[40]) {
      const cx = e.verts[0].x, cy = e.verts[0].y, r = parseFloat(c[40]);
      let a0 = (parseFloat(c[50] || "0") * Math.PI) / 180;
      let a1 = (parseFloat(c[51] || "360") * Math.PI) / 180;
      if (a1 <= a0) a1 += Math.PI * 2;
      const n = Math.max(8, Math.ceil(((a1 - a0) / (Math.PI * 2)) * 32));
      const pts = [];
      for (let k = 0; k <= n; k++) {
        const a = a0 + ((a1 - a0) * k) / n;
        pts.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]);
      }
      ents.push({ kind: "poly", pts, closed: false });
    } else if ((e.type === "TEXT" || e.type === "MTEXT") && e.verts[0] && c[1]) {
      const t = String(c[1]).replace(/\\P/g, " ").replace(/\{|\}|\\[A-Za-z][^;]*;/g, "");
      ents.push({ kind: "text", x: e.verts[0].x, y: e.verts[0].y, h: parseFloat(c[40] || "0.3"), text: t });
    }
  }

  // flip Y (DXF is y-up, screen is y-down)
  for (const e of ents) {
    if (e.pts) e.pts = e.pts.map(([x, y]) => [x, -y]);
    if (e.cy !== undefined) e.cy = -e.cy;
    if (e.y !== undefined && e.kind === "text") e.y = -e.y;
  }

  // bounds
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  const grow = (x, y) => { minX = Math.min(minX, x); minY = Math.min(minY, y); maxX = Math.max(maxX, x); maxY = Math.max(maxY, y); };
  for (const e of ents) {
    if (e.pts) e.pts.forEach(([x, y]) => grow(x, y));
    if (e.kind === "circle") { grow(e.cx - e.r, e.cy - e.r); grow(e.cx + e.r, e.cy + e.r); }
    if (e.kind === "text") grow(e.x, e.y);
  }
  if (!isFinite(minX)) { minX = 0; minY = 0; maxX = 40; maxY = 25; }

  // unit heuristic: drawings larger than ~500 units are assumed mm → convert to m
  let unitScale = 1, assumedUnits = "m";
  const span = Math.max(maxX - minX, maxY - minY);
  if (span > 2000) { unitScale = 0.001; assumedUnits = "mm"; }
  else if (span > 500) { unitScale = 0.01; assumedUnits = "cm"; }
  if (unitScale !== 1) {
    for (const e of ents) {
      if (e.pts) e.pts = e.pts.map(([x, y]) => [x * unitScale, y * unitScale]);
      if (e.kind === "circle") { e.cx *= unitScale; e.cy *= unitScale; e.r *= unitScale; }
      if (e.kind === "text") { e.x *= unitScale; e.y *= unitScale; e.h *= unitScale; }
    }
    minX *= unitScale; minY *= unitScale; maxX *= unitScale; maxY *= unitScale;
  }
  return { entities: ents, bounds: { minX, minY, maxX, maxY }, assumedUnits, count: ents.length };
}

/* ----------------------- equipment list import --------------------- */

const COL_ALIASES = {
  tag: ["tag", "equipment tag", "eq tag", "equipment id", "id", "item", "item no", "item number", "asset tag"],
  name: ["name", "description", "equipment name", "equipment description", "desc", "title"],
  type: ["type", "category", "equipment type", "class", "classification"],
  manufacturer: ["manufacturer", "mfr", "mfg", "vendor", "make", "supplier", "oem"],
  model: ["model", "model no", "model number", "model #"],
  power: ["power", "kw", "power (kw)", "power kw", "rated power", "load", "load (kw)", "hp"],
  x: ["x", "loc x", "x (m)", "pos x", "x coord"],
  y: ["y", "loc y", "y (m)", "pos y", "y coord"],
  w: ["width", "w", "width (m)", "w (m)"],
  d: ["depth", "d", "length", "l", "depth (m)", "length (m)"],
  h: ["height", "h", "height (m)", "h (m)"],
  area: ["area", "room", "zone", "location", "space"],
};

function mapColumns(headers) {
  const map = {};
  const norm = headers.map(h => String(h || "").trim().toLowerCase());
  for (const key of Object.keys(COL_ALIASES)) {
    const idx = norm.findIndex(h => COL_ALIASES[key].includes(h));
    if (idx >= 0) map[key] = headers[idx];
  }
  return map;
}

function rowsToEquipment(rows) {
  if (!rows.length) return [];
  const map = mapColumns(Object.keys(rows[0]));
  return rows
    .filter(r => Object.values(r).some(v => String(v || "").trim() !== ""))
    .map((r, i) => {
      const num = k => { const v = parseFloat(r[map[k]]); return Number.isFinite(v) ? v : undefined; };
      const x = num("x"), y = num("y");
      return {
        id: uid(),
        tag: String(r[map.tag] ?? `EQ-${String(i + 1).padStart(3, "0")}`).trim(),
        name: String(r[map.name] ?? "").trim() || "Unnamed equipment",
        type: String(r[map.type] ?? "Process").trim(),
        manufacturer: String(r[map.manufacturer] ?? "").trim(),
        model: String(r[map.model] ?? "").trim(),
        power: num("power") ?? "",
        area: String(r[map.area] ?? "").trim(),
        x: x, y: y, placed: x !== undefined && y !== undefined,
        w: num("w") ?? 2, d: num("d") ?? 1.5, h: num("h") ?? 1.8,
        bom: [],
      };
    });
}

function downloadFile(name, text, mime = "application/json") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name; a.click();
  URL.revokeObjectURL(url);
}

/* ----------------------------- demo data --------------------------- */

function demoProject() {
  const wall = (pts, closed = true) => ({ kind: "poly", pts, closed });
  const entities = [
    wall([[0, 0], [42, 0], [42, 26], [0, 26]]),
    wall([[16, 0], [16, 10], [28, 10], [28, 0]], false),
    wall([[0, 16], [12, 16], [12, 26]], false),
    wall([[30, 26], [30, 18], [42, 18]], false),
    { kind: "text", x: 5, y: 21.5, h: 0.8, text: "UTILITY ROOM" },
    { kind: "text", x: 20, y: 5.5, h: 0.8, text: "PROCESS HALL A" },
    { kind: "text", x: 33.5, y: 22.5, h: 0.8, text: "PACKAGING" },
    { kind: "text", x: 20, y: 14, h: 0.8, text: "PROCESS HALL B" },
    { kind: "circle", cx: 21, cy: 3, r: 0.15 },
    { kind: "circle", cx: 21, cy: 7, r: 0.15 },
  ];
  const eq = (o) => ({ id: uid(), placed: true, bom: [], w: 2, d: 1.5, h: 1.8, power: "", area: "", manufacturer: "", model: "", ...o });
  const equipment = [
    eq({ tag: "MX-101", name: "High-shear batch mixer", type: "Process", manufacturer: "Silverson", model: "AX-450", power: 45, x: 6, y: 5, w: 2.4, d: 2.4, h: 2.6, area: "Process Hall A",
      bom: [{ id: uid(), pn: "AX450-SEAL-KIT", desc: "Mechanical seal kit", qty: 1 }, { id: uid(), pn: "MTR-45KW-IE3", desc: "45 kW IE3 motor", qty: 1 }, { id: uid(), pn: "IMP-HS-300", desc: "High-shear impeller 300mm", qty: 1 }] }),
    eq({ tag: "R-201", name: "Jacketed reactor 2000L", type: "Storage Tank", manufacturer: "De Dietrich", model: "AE-2000", power: 15, x: 12, y: 5, w: 2.8, d: 2.8, h: 3.4, area: "Process Hall A",
      bom: [{ id: uid(), pn: "AGIT-PBT-2000", desc: "Pitched-blade agitator", qty: 1 }, { id: uid(), pn: "RV-DN80", desc: "Bottom outlet valve DN80", qty: 1 }] }),
    eq({ tag: "P-301", name: "Transfer pump", type: "Pump", manufacturer: "Alfa Laval", model: "LKH-25", power: 7.5, x: 12, y: 8.5, w: 1, d: 0.8, h: 0.9, area: "Process Hall A" }),
    eq({ tag: "CIP-01", name: "CIP skid, 3-tank", type: "Skid", manufacturer: "GEA", model: "Compact CIP", power: 22, x: 21, y: 13, w: 4, d: 2, h: 2.4, area: "Process Hall B",
      bom: [{ id: uid(), pn: "TK-CIP-500", desc: "500L caustic tank", qty: 3 }, { id: uid(), pn: "HX-PHE-30", desc: "Plate heat exchanger", qty: 1 }] }),
    eq({ tag: "FIL-401", name: "Rotary filler, 12-head", type: "Packaging", manufacturer: "Krones", model: "Modulfill", power: 30, x: 35, y: 22, w: 3.5, d: 3.5, h: 2.8, area: "Packaging" }),
  ];
  const util = (o) => ({ id: uid(), ...o });
  const utilities = [
    util({ tag: "AHU-01", name: "Air handling unit, 12,000 CFM", type: "hvac", x: 4, y: 19, capacity: "12,000 CFM" }),
    util({ tag: "CH-01", name: "Air-cooled chiller", type: "chw", x: 8.5, y: 19, capacity: "250 kW" }),
    util({ tag: "MCC-01", name: "Motor control center", type: "power", x: 4, y: 23, capacity: "800 A" }),
    util({ tag: "AC-01", name: "Oil-free air compressor", type: "compair", x: 8.5, y: 23, capacity: "90 CFM @ 7 bar" }),
  ];
  const find = t => equipment.find(e => e.tag === t).id;
  const findU = t => utilities.find(u => u.tag === t).id;
  const conn = (from, to, type, medium, size) => ({ id: uid(), fromId: from, toId: to, utilityType: type, medium, size, notes: "" });
  const connections = [
    conn(findU("CH-01"), find("R-201"), "chw", "Chilled water supply/return", "DN50"),
    conn(findU("MCC-01"), find("MX-101"), "power", "480V 3ph feeder", "45 kW"),
    conn(findU("MCC-01"), find("FIL-401"), "power", "480V 3ph feeder", "30 kW"),
    conn(findU("AHU-01"), find("FIL-401"), "hvac", "Supply air, ISO 8", "600x400 duct"),
    conn(findU("AC-01"), find("FIL-401"), "compair", "Instrument air", "DN25"),
    conn(find("R-201"), find("CIP-01"), "process", "CIP return", "DN40"),
    conn(find("MX-101"), find("R-201"), "process", "Product transfer", "DN50"),
  ];
  return {
    name: "Line 4 — Beverage Processing",
    site: "Fresno Plant",
    revision: "B",
    floorplan: { entities, bounds: { minX: 0, minY: 0, maxX: 42, maxY: 26 }, assumedUnits: "m", source: "demo-floorplan.dxf" },
    equipment, utilities, connections,
    github: { owner: "", repo: "", branch: "main", path: "digital-twin/project.json", token: "" },
  };
}

function emptyProject() {
  return {
    name: "Untitled facility",
    site: "",
    revision: "A",
    floorplan: null,
    equipment: [], utilities: [], connections: [],
    github: { owner: "", repo: "", branch: "main", path: "digital-twin/project.json", token: "" },
  };
}

/* ------------------------------ styles ------------------------------ */

const CSS = `
:root{
  --bg:#101418; --panel:#171C21; --panel2:#1D242B; --line:#2A333C; --line2:#38434E;
  --text:#DCE3E9; --muted:#7E8B95; --faint:#55616B; --accent:#F08C3A; --accent-dim:#7a4a22;
  --danger:#E06552; --ok:#62D26F;
  --mono:ui-monospace,'SF Mono','Cascadia Code',Consolas,Menlo,monospace;
  --sans:'Segoe UI',system-ui,-apple-system,Roboto,sans-serif;
}
.ft-root{position:fixed;inset:0;background:var(--bg);color:var(--text);font-family:var(--sans);font-size:13px;display:flex;flex-direction:column;overflow:hidden}
.ft-root *{box-sizing:border-box}
.ft-root ::-webkit-scrollbar{width:8px;height:8px}
.ft-root ::-webkit-scrollbar-thumb{background:var(--line2);border-radius:4px}
.ft-root ::-webkit-scrollbar-track{background:transparent}

.tb{display:flex;align-items:stretch;border-bottom:1px solid var(--line);background:var(--panel);height:52px;flex:none}
.tb-block{display:flex;flex-direction:column;justify-content:center;padding:0 14px;border-right:1px solid var(--line)}
.tb-eyebrow{font-family:var(--mono);font-size:9px;letter-spacing:.14em;color:var(--faint);text-transform:uppercase}
.tb-title{font-weight:600;font-size:14px;letter-spacing:.01em;white-space:nowrap}
.tb-title input{background:transparent;border:none;color:var(--text);font:inherit;outline:none;width:220px}
.tb-spacer{flex:1}
.tb-group{display:flex;align-items:center;gap:6px;padding:0 12px;border-left:1px solid var(--line)}

.btn{font-family:var(--mono);font-size:11px;letter-spacing:.04em;padding:6px 11px;border:1px solid var(--line2);background:var(--panel2);color:var(--text);border-radius:3px;cursor:pointer;white-space:nowrap}
.btn:hover{border-color:var(--accent);color:#fff}
.btn:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
.btn.active{background:var(--accent);border-color:var(--accent);color:#14181c;font-weight:700}
.btn.ghost{background:transparent}
.btn.small{padding:3px 8px;font-size:10px}
.btn.danger:hover{border-color:var(--danger);color:var(--danger)}
.btn:disabled{opacity:.4;cursor:default}

.main{flex:1;display:flex;min-height:0}
.side{width:252px;flex:none;background:var(--panel);border-right:1px solid var(--line);display:flex;flex-direction:column;min-height:0}
.inspector{width:300px;flex:none;background:var(--panel);border-left:1px solid var(--line);display:flex;flex-direction:column;min-height:0}
.side-tabs{display:flex;border-bottom:1px solid var(--line);flex:none}
.side-tab{flex:1;padding:8px 4px;text-align:center;font-family:var(--mono);font-size:10px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;background:none;border-top:none;border-left:none;border-right:none}
.side-tab.on{color:var(--text);border-bottom-color:var(--accent)}
.side-body{flex:1;overflow-y:auto;padding:12px}

.sec{font-family:var(--mono);font-size:9px;letter-spacing:.14em;text-transform:uppercase;color:var(--faint);margin:14px 0 6px;display:flex;align-items:center;gap:8px}
.sec:first-child{margin-top:0}
.sec::after{content:"";flex:1;height:1px;background:var(--line)}
.hint{color:var(--muted);font-size:11px;line-height:1.5;margin:6px 0}

.row-item{display:flex;align-items:center;gap:8px;padding:6px 8px;border:1px solid transparent;border-radius:3px;cursor:pointer;margin-bottom:2px}
.row-item:hover{background:var(--panel2)}
.row-item.on{background:var(--panel2);border-color:var(--accent-dim)}
.row-dot{width:8px;height:8px;border-radius:2px;flex:none}
.row-tag{font-family:var(--mono);font-size:11px;font-weight:700;white-space:nowrap}
.row-name{font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1}
.pill{font-family:var(--mono);font-size:9px;padding:1px 5px;border-radius:2px;border:1px solid var(--line2);color:var(--muted);flex:none}

.canvas-wrap{flex:1;position:relative;min-width:0;background:
  radial-gradient(circle at 30% 20%, #131920 0%, var(--bg) 60%)}
.canvas-svg{position:absolute;inset:0;width:100%;height:100%;display:block}
.canvas-3d{position:absolute;inset:0}
.statusbar{position:absolute;left:0;right:0;bottom:0;height:26px;display:flex;align-items:center;gap:0;background:rgba(16,20,24,.92);border-top:1px solid var(--line);font-family:var(--mono);font-size:10px;color:var(--muted);z-index:5}
.statusbar>div{padding:0 12px;border-right:1px solid var(--line);height:100%;display:flex;align-items:center;gap:6px;white-space:nowrap}
.statusbar b{color:var(--text);font-weight:600}
.mode-banner{position:absolute;top:12px;left:50%;transform:translateX(-50%);background:var(--accent);color:#14181c;font-family:var(--mono);font-size:11px;font-weight:700;padding:5px 14px;border-radius:3px;z-index:6;letter-spacing:.04em}

.fld{margin-bottom:8px}
.fld label{display:block;font-family:var(--mono);font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--faint);margin-bottom:3px}
.fld input,.fld select,.fld textarea{width:100%;background:var(--bg);border:1px solid var(--line2);color:var(--text);border-radius:3px;padding:6px 8px;font-size:12px;font-family:var(--sans);outline:none}
.fld input:focus,.fld select:focus,.fld textarea:focus{border-color:var(--accent)}
.fld-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.fld-grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px}
.mono-in input{font-family:var(--mono);font-size:11px}

.bom-table{width:100%;border-collapse:collapse;font-size:11px}
.bom-table th{font-family:var(--mono);font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:var(--faint);text-align:left;padding:4px 6px;border-bottom:1px solid var(--line)}
.bom-table td{padding:3px 4px;border-bottom:1px solid var(--line)}
.bom-table input{width:100%;background:transparent;border:none;color:var(--text);font-size:11px;outline:none;padding:2px}
.bom-table input:focus{background:var(--bg)}

.conn-chip{display:flex;align-items:center;gap:8px;padding:6px 8px;border:1px solid var(--line);border-radius:3px;margin-bottom:4px;cursor:pointer}
.conn-chip:hover{border-color:var(--line2)}
.conn-chip.on{border-color:var(--accent)}

.modal-back{position:fixed;inset:0;background:rgba(8,10,12,.7);display:flex;align-items:center;justify-content:center;z-index:50}
.modal{width:440px;max-width:92vw;max-height:86vh;overflow-y:auto;background:var(--panel);border:1px solid var(--line2);border-radius:6px;padding:18px}
.modal h3{margin:0 0 4px;font-size:15px}
.modal .sub{color:var(--muted);font-size:11px;margin-bottom:14px;line-height:1.5}

.toasts{position:absolute;bottom:36px;right:12px;display:flex;flex-direction:column;gap:6px;z-index:20}
.toast{background:var(--panel2);border:1px solid var(--line2);border-left:3px solid var(--accent);padding:8px 12px;border-radius:3px;font-size:11px;max-width:300px;box-shadow:0 4px 16px rgba(0,0,0,.4)}
.toast.err{border-left-color:var(--danger)}
.toast.ok{border-left-color:var(--ok)}

.empty-canvas{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;color:var(--muted);z-index:2;pointer-events:none}
.empty-canvas .big{font-family:var(--mono);font-size:13px;letter-spacing:.12em;color:var(--faint)}
.empty-canvas button{pointer-events:auto}

@media (max-width:900px){ .side{width:200px}.inspector{width:250px} }
`;

/* --------------------------- small pieces --------------------------- */

function Fld({ label, children }) {
  return <div className="fld"><label>{label}</label>{children}</div>;
}

function nodeAnchor(node) {
  return { x: node.x ?? 0, y: node.y ?? 0 };
}

function orthoPath(a, b) {
  // simple elbow: horizontal then vertical from a to b
  const midX = a.x;
  return `M ${a.x} ${a.y} L ${midX} ${b.y} L ${b.x} ${b.y}`;
}

/* ------------------------------- app -------------------------------- */

export default function FacilityTwin() {
  const [project, setProject] = useState(emptyProject);
  const [selected, setSelected] = useState(null);        // {kind:'equipment'|'utility'|'connection', id}
  const [mode, setMode] = useState("select");            // select | add-equip | add-utility | connect
  const [pendingUtilType, setPendingUtilType] = useState("hvac");
  const [connectFrom, setConnectFrom] = useState(null);
  const [view, setView] = useState("2d");
  const [leftTab, setLeftTab] = useState("import");
  const [ghOpen, setGhOpen] = useState(false);
  const [ghBusy, setGhBusy] = useState(false);
  const [toasts, setToasts] = useState([]);
  const [cursor, setCursor] = useState({ x: 0, y: 0 });
  const [cam, setCam] = useState({ x: -2, y: -2, w: 50 }); // world viewBox (h derived)
  const [placing, setPlacing] = useState(null);           // staged equipment id being placed
  const [loaded, setLoaded] = useState(false);
  const [aspect, setAspect] = useState(0.62);

  const svgRef = useRef(null);
  const dragRef = useRef(null);
  const threeRef = useRef(null);
  const fileDxfRef = useRef(null);
  const fileListRef = useRef(null);
  const fileJsonRef = useRef(null);

  const toast = useCallback((msg, kind = "") => {
    const id = uid();
    setToasts(t => [...t, { id, msg, kind }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 4200);
  }, []);

  useEffect(() => {
    const measure = () => {
      const el = svgRef.current;
      if (el) { const r = el.getBoundingClientRect(); if (r.width > 0) setAspect(r.height / r.width); }
    };
    measure();
    window.addEventListener("resize", measure);
    const t = setInterval(measure, 1500);
    return () => { window.removeEventListener("resize", measure); clearInterval(t); };
  }, [view]);

  /* ------------------------- persistence ------------------------- */

  useEffect(() => {
    try {
      const saved = localStorage.getItem("facility-twin:project");
      if (saved) { setProject(JSON.parse(saved)); toast("Restored saved project", "ok"); }
    } catch { /* first run — nothing saved yet */ }
    setLoaded(true);
  }, []); // eslint-disable-line

  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => {
      try { localStorage.setItem("facility-twin:project", JSON.stringify(project)); } catch { /* storage full */ }
    }, 1200);
    return () => clearTimeout(t);
  }, [project, loaded]);

  /* ------------------------- derived data ------------------------ */

  const nodes = useMemo(() => {
    const m = {};
    project.equipment.forEach(e => { m[e.id] = { ...e, kind: "equipment" }; });
    project.utilities.forEach(u => { m[u.id] = { ...u, kind: "utility" }; });
    return m;
  }, [project.equipment, project.utilities]);

  const selEquip = selected?.kind === "equipment" ? project.equipment.find(e => e.id === selected.id) : null;
  const selUtil = selected?.kind === "utility" ? project.utilities.find(u => u.id === selected.id) : null;
  const selConn = selected?.kind === "connection" ? project.connections.find(c => c.id === selected.id) : null;

  const fitView = useCallback((b) => {
    const bb = b || project.floorplan?.bounds;
    if (!bb) return;
    const pad = Math.max(2, (bb.maxX - bb.minX) * 0.08);
    setCam({ x: bb.minX - pad, y: bb.minY - pad, w: (bb.maxX - bb.minX) + pad * 2 });
  }, [project.floorplan]);

  /* --------------------------- imports --------------------------- */

  async function onDxfFile(file) {
    if (!file) return;
    const isDwg = /\.dwg$/i.test(file.name);
    if (isDwg) {
      toast("DWG is a closed binary format and can't be parsed in-browser. Export to DXF (AutoCAD: SAVEAS → DXF, or the free ODA File Converter) and re-import.", "err");
      return;
    }
    const text = await file.text();
    try {
      const fp = parseDXF(text);
      if (!fp.count) { toast("No supported entities found in this DXF (LINE, LWPOLYLINE, CIRCLE, ARC, TEXT).", "err"); return; }
      setProject(p => ({ ...p, floorplan: { ...fp, source: file.name } }));
      fitView(fp.bounds);
      toast(`Imported ${fp.count} entities from ${file.name} (units read as ${fp.assumedUnits}, shown in m)`, "ok");
    } catch (e) {
      toast("Couldn't parse this DXF: " + e.message, "err");
    }
  }

  async function onEquipmentFile(file) {
    if (!file) return;
    try {
      let rows = [];
      if (/\.(xlsx|xls|xlsm)$/i.test(file.name)) {
        const wb = XLSX.read(await file.arrayBuffer());
        rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
      } else {
        const parsed = Papa.parse(await file.text(), { header: true, skipEmptyLines: true });
        rows = parsed.data;
      }
      const eqs = rowsToEquipment(rows);
      if (!eqs.length) { toast("No rows recognized. Expected headers like Tag, Description, Manufacturer, Model…", "err"); return; }
      setProject(p => ({ ...p, equipment: [...p.equipment, ...eqs] }));
      const unplaced = eqs.filter(e => !e.placed).length;
      setLeftTab("equipment");
      toast(`Imported ${eqs.length} equipment records${unplaced ? ` — ${unplaced} without coordinates; click one, then click the plan to place it` : ""}`, "ok");
    } catch (e) {
      toast("Import failed: " + e.message, "err");
    }
  }

  function onProjectJson(file) {
    if (!file) return;
    file.text().then(t => {
      try {
        const p = JSON.parse(t);
        if (!p.equipment) throw new Error("not a FacilityTwin project file");
        setProject({ ...emptyProject(), ...p });
        toast("Project loaded from JSON", "ok");
      } catch (e) { toast("Invalid project file: " + e.message, "err"); }
    });
  }

  /* ----------------------- canvas interaction --------------------- */

  const clientToWorld = useCallback((cx, cy) => {
    const el = svgRef.current;
    if (!el) return { x: 0, y: 0 };
    const r = el.getBoundingClientRect();
    const h = cam.w * (r.height / r.width);
    return { x: cam.x + ((cx - r.left) / r.width) * cam.w, y: cam.y + ((cy - r.top) / r.height) * h };
  }, [cam]);

  function onWheel(e) {
    const pt = clientToWorld(e.clientX, e.clientY);
    const f = e.deltaY > 0 ? 1.12 : 1 / 1.12;
    setCam(c => {
      const w = Math.min(2000, Math.max(1, c.w * f));
      return { x: pt.x - (pt.x - c.x) * (w / c.w), y: pt.y - (pt.y - c.y) * (w / c.w), w };
    });
  }

  function onCanvasDown(e, hit) {
    const pt = clientToWorld(e.clientX, e.clientY);
    if (mode === "add-equip") {
      const eq = { id: uid(), tag: `EQ-${String(project.equipment.length + 1).padStart(3, "0")}`, name: "New equipment", type: "Process", manufacturer: "", model: "", power: "", area: "", x: pt.x, y: pt.y, placed: true, w: 2, d: 1.5, h: 1.8, bom: [] };
      setProject(p => ({ ...p, equipment: [...p.equipment, eq] }));
      setSelected({ kind: "equipment", id: eq.id });
      setMode("select");
      return;
    }
    if (mode === "add-utility") {
      const u = { id: uid(), tag: `${UTILITY_TYPES[pendingUtilType].short}-${String(project.utilities.length + 1).padStart(2, "0")}`, name: UTILITY_TYPES[pendingUtilType].label, type: pendingUtilType, x: pt.x, y: pt.y, capacity: "" };
      setProject(p => ({ ...p, utilities: [...p.utilities, u] }));
      setSelected({ kind: "utility", id: u.id });
      setMode("select");
      return;
    }
    if (placing) {
      setProject(p => ({ ...p, equipment: p.equipment.map(q => q.id === placing ? { ...q, x: pt.x, y: pt.y, placed: true } : q) }));
      setSelected({ kind: "equipment", id: placing });
      setPlacing(null);
      toast("Equipment placed", "ok");
      return;
    }
    if (mode === "connect") {
      if (hit) {
        if (!connectFrom) { setConnectFrom(hit.id); return; }
        if (connectFrom !== hit.id) {
          const from = nodes[connectFrom], to = nodes[hit.id];
          const utype = from?.kind === "utility" ? from.type : to?.kind === "utility" ? to.type : "process";
          const c = { id: uid(), fromId: connectFrom, toId: hit.id, utilityType: utype, medium: "", size: "", notes: "" };
          setProject(p => ({ ...p, connections: [...p.connections, c] }));
          setSelected({ kind: "connection", id: c.id });
          setConnectFrom(null);
          setMode("select");
          toast(`Connected ${from?.tag} → ${to?.tag}`, "ok");
        }
      }
      return;
    }
    // select mode
    if (hit) {
      setSelected({ kind: hit.kind, id: hit.id });
      dragRef.current = { id: hit.id, kind: hit.kind, start: pt, orig: { x: nodes[hit.id].x, y: nodes[hit.id].y }, moved: false };
    } else {
      setSelected(null);
      dragRef.current = { pan: true, start: { x: e.clientX, y: e.clientY }, cam0: cam };
    }
  }

  function onCanvasMove(e) {
    const pt = clientToWorld(e.clientX, e.clientY);
    setCursor(pt);
    const d = dragRef.current;
    if (!d) return;
    if (d.pan) {
      const el = svgRef.current; if (!el) return;
      const r = el.getBoundingClientRect();
      const dx = ((e.clientX - d.start.x) / r.width) * d.cam0.w;
      const dy = ((e.clientY - d.start.y) / r.width) * d.cam0.w;
      setCam({ ...d.cam0, x: d.cam0.x - dx, y: d.cam0.y - dy });
      return;
    }
    const dx = pt.x - d.start.x, dy = pt.y - d.start.y;
    if (Math.abs(dx) + Math.abs(dy) > 0.05) d.moved = true;
    if (!d.moved) return;
    const nx = d.orig.x + dx, ny = d.orig.y + dy;
    if (d.kind === "equipment") setProject(p => ({ ...p, equipment: p.equipment.map(q => q.id === d.id ? { ...q, x: nx, y: ny } : q) }));
    else if (d.kind === "utility") setProject(p => ({ ...p, utilities: p.utilities.map(q => q.id === d.id ? { ...q, x: nx, y: ny } : q) }));
  }

  function onCanvasUp() { dragRef.current = null; }

  /* --------------------------- mutations ------------------------- */

  const patchEquip = (id, patch) => setProject(p => ({ ...p, equipment: p.equipment.map(e => e.id === id ? { ...e, ...patch } : e) }));
  const patchUtil = (id, patch) => setProject(p => ({ ...p, utilities: p.utilities.map(u => u.id === id ? { ...u, ...patch } : u) }));
  const patchConn = (id, patch) => setProject(p => ({ ...p, connections: p.connections.map(c => c.id === id ? { ...c, ...patch } : c) }));

  function removeSelected() {
    if (!selected) return;
    setProject(p => {
      if (selected.kind === "connection") return { ...p, connections: p.connections.filter(c => c.id !== selected.id) };
      const drop = { ...p, connections: p.connections.filter(c => c.fromId !== selected.id && c.toId !== selected.id) };
      if (selected.kind === "equipment") return { ...drop, equipment: p.equipment.filter(e => e.id !== selected.id) };
      return { ...drop, utilities: p.utilities.filter(u => u.id !== selected.id) };
    });
    setSelected(null);
  }

  /* ---------------------------- GitHub ---------------------------- */

  async function ghRequest(url, opts = {}) {
    const res = await fetch(url, {
      ...opts,
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${project.github.token}`,
        "X-GitHub-Api-Version": "2022-11-28",
        ...(opts.headers || {}),
      },
    });
    return res;
  }

  async function ghTest() {
    const { owner, repo } = project.github;
    setGhBusy(true);
    try {
      const res = await ghRequest(`https://api.github.com/repos/${owner}/${repo}`);
      if (res.ok) toast(`Connected to ${owner}/${repo}`, "ok");
      else toast(`GitHub responded ${res.status} — check owner/repo and token scopes (repo → contents: read/write)`, "err");
    } catch (e) {
      toast("Network blocked in this preview environment. The same code works when the app is deployed (GitHub's API allows browser CORS).", "err");
    }
    setGhBusy(false);
  }

  async function ghPush() {
    const { owner, repo, branch, path } = project.github;
    setGhBusy(true);
    try {
      const clean = { ...project, github: { ...project.github, token: "" } };
      const json = JSON.stringify(clean, null, 2);
      const content = btoa(unescape(encodeURIComponent(json)));
      const base = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
      let sha;
      const head = await ghRequest(`${base}?ref=${branch}`);
      if (head.ok) sha = (await head.json()).sha;
      const res = await ghRequest(base, {
        method: "PUT",
        body: JSON.stringify({ message: `FacilityTwin snapshot — ${project.name} rev ${project.revision}`, content, branch, ...(sha ? { sha } : {}) }),
      });
      if (res.ok) toast(`Snapshot committed to ${owner}/${repo}@${branch}`, "ok");
      else toast(`Commit failed (${res.status}). Token needs Contents read/write on this repo.`, "err");
    } catch (e) {
      toast("Network blocked in this preview environment. Deploy the app (or run locally) and the push will work — the token stays in your browser.", "err");
    }
    setGhBusy(false);
  }

  /* ---------------------------- 3D view --------------------------- */

  useEffect(() => {
    if (view !== "3d" || !threeRef.current) return;
    const el = threeRef.current;
    const W = el.clientWidth, H = el.clientHeight;
    const scene = new THREE.Scene();
    scene.background = new THREE.Color(0x101418);
    scene.fog = new THREE.Fog(0x101418, 60, 220);

    const camera = new THREE.PerspectiveCamera(50, W / H, 0.1, 500);
    const renderer = new THREE.WebGLRenderer({ antialias: true });
    renderer.setSize(W, H);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    el.appendChild(renderer.domElement);

    scene.add(new THREE.AmbientLight(0xffffff, 0.55));
    const sun = new THREE.DirectionalLight(0xffe8c8, 0.9);
    sun.position.set(30, 50, 20);
    scene.add(sun);

    const b = project.floorplan?.bounds || { minX: 0, minY: 0, maxX: 40, maxY: 25 };
    const cx = (b.minX + b.maxX) / 2, cz = (b.minY + b.maxY) / 2;
    const span = Math.max(b.maxX - b.minX, b.maxY - b.minY);

    const floor = new THREE.Mesh(
      new THREE.PlaneGeometry(b.maxX - b.minX + 8, b.maxY - b.minY + 8),
      new THREE.MeshStandardMaterial({ color: 0x171c21, roughness: 0.95 })
    );
    floor.rotation.x = -Math.PI / 2;
    floor.position.set(cx, -0.02, cz);
    scene.add(floor);
    const grid = new THREE.GridHelper(Math.ceil(span + 8), Math.ceil(span + 8), 0x2a333c, 0x1d242b);
    grid.position.set(cx, 0, cz);
    scene.add(grid);

    // floor plan linework
    if (project.floorplan) {
      const pts = [];
      for (const e of project.floorplan.entities) {
        if (e.kind === "line" || e.kind === "poly") {
          const arr = e.kind === "poly" && e.closed ? [...e.pts, e.pts[0]] : e.pts;
          for (let i = 0; i < arr.length - 1; i++) {
            pts.push(new THREE.Vector3(arr[i][0], 0.02, arr[i][1]), new THREE.Vector3(arr[i + 1][0], 0.02, arr[i + 1][1]));
          }
        }
      }
      if (pts.length) {
        const g = new THREE.BufferGeometry().setFromPoints(pts);
        scene.add(new THREE.LineSegments(g, new THREE.LineBasicMaterial({ color: 0x55616b })));
      }
    }

    // equipment boxes
    for (const e of project.equipment.filter(q => q.placed)) {
      const isSel = selected?.id === e.id;
      const mesh = new THREE.Mesh(
        new THREE.BoxGeometry(e.w, e.h, e.d),
        new THREE.MeshStandardMaterial({ color: isSel ? 0xf08c3a : 0x3a4650, roughness: 0.6, metalness: 0.25 })
      );
      mesh.position.set(e.x, e.h / 2, e.y);
      scene.add(mesh);
      const edges = new THREE.LineSegments(new THREE.EdgesGeometry(mesh.geometry), new THREE.LineBasicMaterial({ color: isSel ? 0xffc48a : 0x55616b }));
      edges.position.copy(mesh.position);
      scene.add(edges);
    }

    // utility nodes as cylinders
    for (const u of project.utilities) {
      const col = new THREE.Color(UTILITY_TYPES[u.type]?.color || "#888");
      const cyl = new THREE.Mesh(new THREE.CylinderGeometry(0.7, 0.7, 2.2, 20), new THREE.MeshStandardMaterial({ color: col, roughness: 0.5 }));
      cyl.position.set(u.x, 1.1, u.y);
      scene.add(cyl);
    }

    // connections as elevated runs
    for (const c of project.connections) {
      const a = nodes[c.fromId], t = nodes[c.toId];
      if (!a || !t || a.x === undefined || t.x === undefined) continue;
      const col = new THREE.Color(UTILITY_TYPES[c.utilityType]?.color || "#888");
      const hRun = 3.2;
      const p = [
        new THREE.Vector3(a.x, (a.h || 2.2), a.y),
        new THREE.Vector3(a.x, hRun, a.y),
        new THREE.Vector3(a.x, hRun, t.y),
        new THREE.Vector3(t.x, hRun, t.y),
        new THREE.Vector3(t.x, (t.h || 2.2), t.y),
      ];
      const g = new THREE.BufferGeometry().setFromPoints(p);
      scene.add(new THREE.Line(g, new THREE.LineBasicMaterial({ color: col })));
    }

    // custom orbit controls
    let az = -Math.PI / 4, elv = 0.9, radius = span * 1.1 + 10;
    const target = new THREE.Vector3(cx, 1, cz);
    const applyCam = () => {
      camera.position.set(
        target.x + radius * Math.cos(elv) * Math.cos(az),
        target.y + radius * Math.sin(elv),
        target.z + radius * Math.cos(elv) * Math.sin(az)
      );
      camera.lookAt(target);
    };
    applyCam();

    let down = null;
    const onDown = (e) => { down = { x: e.clientX, y: e.clientY, az, elv }; };
    const onMove = (e) => {
      if (!down) return;
      az = down.az + (e.clientX - down.x) * 0.006;
      elv = Math.min(1.45, Math.max(0.1, down.elv + (e.clientY - down.y) * 0.005));
      applyCam();
    };
    const onUp = () => { down = null; };
    const onZoom = (e) => { e.preventDefault(); radius = Math.min(400, Math.max(6, radius * (e.deltaY > 0 ? 1.1 : 0.9))); applyCam(); };
    renderer.domElement.addEventListener("mousedown", onDown);
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    renderer.domElement.addEventListener("wheel", onZoom, { passive: false });

    let alive = true;
    const loop = () => { if (!alive) return; renderer.render(scene, camera); requestAnimationFrame(loop); };
    loop();

    const onResize = () => {
      const w2 = el.clientWidth, h2 = el.clientHeight;
      camera.aspect = w2 / h2; camera.updateProjectionMatrix(); renderer.setSize(w2, h2);
    };
    window.addEventListener("resize", onResize);

    return () => {
      alive = false;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      window.removeEventListener("resize", onResize);
      renderer.dispose();
      el.innerHTML = "";
    };
  }, [view, project, selected, nodes]);

  /* ---------------------------- render ---------------------------- */

  const camH = cam.w * aspect;
  const gridStep = cam.w > 120 ? 10 : cam.w > 30 ? 5 : 1;
  const gridLines = useMemo(() => {
    const ls = [];
    const x0 = Math.floor(cam.x / gridStep) * gridStep;
    const y0 = Math.floor(cam.y / gridStep) * gridStep;
    for (let x = x0; x < cam.x + cam.w + gridStep; x += gridStep) ls.push({ v: true, p: x });
    for (let y = y0; y < cam.y + camH + gridStep; y += gridStep) ls.push({ v: false, p: y });
    return ls;
  }, [cam, gridStep, camH]);

  const unplaced = project.equipment.filter(e => !e.placed);
  const connCountFor = id => project.connections.filter(c => c.fromId === id || c.toId === id).length;
  const strokeW = cam.w / 900;

  const modeBanner =
    mode === "add-equip" ? "CLICK PLAN TO PLACE NEW EQUIPMENT" :
    mode === "add-utility" ? `CLICK PLAN TO PLACE ${UTILITY_TYPES[pendingUtilType].label.toUpperCase()} NODE` :
    mode === "connect" ? (connectFrom ? `SOURCE: ${nodes[connectFrom]?.tag} — CLICK TARGET NODE` : "CLICK SOURCE EQUIPMENT / UTILITY") :
    placing ? "CLICK PLAN TO PLACE STAGED EQUIPMENT" : null;

  return (
    <div className="ft-root">
      <style>{CSS}</style>

      {/* ---------- top bar ---------- */}
      <div className="tb">
        <div className="tb-block">
          <div className="tb-eyebrow">FacilityTwin</div>
          <div className="tb-title">
            <input value={project.name} onChange={e => setProject(p => ({ ...p, name: e.target.value }))} aria-label="Project name" />
          </div>
        </div>
        <div className="tb-block">
          <div className="tb-eyebrow">Rev</div>
          <div className="tb-title" style={{ fontFamily: "var(--mono)" }}>
            <input style={{ width: 34 }} value={project.revision} onChange={e => setProject(p => ({ ...p, revision: e.target.value }))} aria-label="Revision" />
          </div>
        </div>
        <div className="tb-group">
          <button className={`btn ${mode === "select" ? "active" : ""}`} onClick={() => { setMode("select"); setConnectFrom(null); }}>SELECT</button>
          <button className={`btn ${mode === "add-equip" ? "active" : ""}`} onClick={() => setMode("add-equip")}>+ EQUIPMENT</button>
          <button className={`btn ${mode === "add-utility" ? "active" : ""}`} onClick={() => setMode("add-utility")}>+ UTILITY</button>
          {mode === "add-utility" && (
            <select value={pendingUtilType} onChange={e => setPendingUtilType(e.target.value)}
              style={{ background: "var(--bg)", color: "var(--text)", border: "1px solid var(--line2)", borderRadius: 3, padding: "5px 6px", fontFamily: "var(--mono)", fontSize: 11 }}>
              {Object.entries(UTILITY_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
          )}
          <button className={`btn ${mode === "connect" ? "active" : ""}`} onClick={() => { setMode("connect"); setConnectFrom(null); }}>⇄ CONNECT</button>
        </div>
        <div className="tb-spacer" />
        <div className="tb-group">
          <button className={`btn ${view === "2d" ? "active" : ""}`} onClick={() => setView("2d")}>2D PLAN</button>
          <button className={`btn ${view === "3d" ? "active" : ""}`} onClick={() => setView("3d")}>3D MODEL</button>
        </div>
        <div className="tb-group">
          <button className="btn" onClick={() => fitView()}>FIT</button>
          <button className="btn" onClick={() => setGhOpen(true)}>
            {project.github.owner ? `⎇ ${project.github.owner}/${project.github.repo}` : "⎇ GITHUB"}
          </button>
        </div>
      </div>

      <div className="main">
        {/* ---------- left panel ---------- */}
        <div className="side">
          <div className="side-tabs">
            {["import", "equipment", "utilities"].map(t => (
              <button key={t} className={`side-tab ${leftTab === t ? "on" : ""}`} onClick={() => setLeftTab(t)}>{t}</button>
            ))}
          </div>
          <div className="side-body">
            {leftTab === "import" && (
              <>
                <div className="sec">Floor plan</div>
                <p className="hint">Import a 2D CAD floor plan as <b>DXF</b>. DWG must be converted to DXF first (free: ODA File Converter).</p>
                <button className="btn" style={{ width: "100%" }} onClick={() => fileDxfRef.current.click()}>IMPORT DXF…</button>
                <input ref={fileDxfRef} type="file" accept=".dxf,.dwg" hidden onChange={e => { onDxfFile(e.target.files[0]); e.target.value = ""; }} />
                {project.floorplan && (
                  <p className="hint" style={{ fontFamily: "var(--mono)", fontSize: 10 }}>
                    ▣ {project.floorplan.source} — {project.floorplan.entities.length} entities
                  </p>
                )}

                <div className="sec">Equipment list / BOM</div>
                <p className="hint">Import CSV or Excel. Recognized headers: Tag, Description, Type, Manufacturer, Model, Power (kW), X, Y, Width, Depth, Height, Area.</p>
                <button className="btn" style={{ width: "100%" }} onClick={() => fileListRef.current.click()}>IMPORT CSV / XLSX…</button>
                <input ref={fileListRef} type="file" accept=".csv,.xlsx,.xls,.xlsm,.tsv" hidden onChange={e => { onEquipmentFile(e.target.files[0]); e.target.value = ""; }} />

                <div className="sec">Project</div>
                <button className="btn ghost" style={{ width: "100%", marginBottom: 6 }} onClick={() => { setProject(demoProject()); setSelected(null); setTimeout(() => fitView(), 0); toast("Demo facility loaded", "ok"); }}>LOAD DEMO FACILITY</button>
                <button className="btn ghost" style={{ width: "100%", marginBottom: 6 }} onClick={() => downloadFile(`${project.name.replace(/\W+/g, "-")}.facilitytwin.json`, JSON.stringify({ ...project, github: { ...project.github, token: "" } }, null, 2))}>EXPORT PROJECT JSON</button>
                <button className="btn ghost" style={{ width: "100%", marginBottom: 6 }} onClick={() => fileJsonRef.current.click()}>OPEN PROJECT JSON…</button>
                <input ref={fileJsonRef} type="file" accept=".json" hidden onChange={e => { onProjectJson(e.target.files[0]); e.target.value = ""; }} />
                <button className="btn ghost danger" style={{ width: "100%" }} onClick={() => { setProject(emptyProject()); setSelected(null); }}>NEW EMPTY PROJECT</button>
              </>
            )}

            {leftTab === "equipment" && (
              <>
                {unplaced.length > 0 && (
                  <>
                    <div className="sec">Staged — click to place</div>
                    {unplaced.map(e => (
                      <div key={e.id} className={`row-item ${placing === e.id ? "on" : ""}`} onClick={() => { setPlacing(e.id); setLeftTab("equipment"); }}>
                        <span className="row-dot" style={{ background: "var(--accent)" }} />
                        <span className="row-tag">{e.tag}</span>
                        <span className="row-name">{e.name}</span>
                      </div>
                    ))}
                  </>
                )}
                <div className="sec">Placed equipment ({project.equipment.filter(e => e.placed).length})</div>
                {project.equipment.filter(e => e.placed).map(e => (
                  <div key={e.id} className={`row-item ${selected?.id === e.id ? "on" : ""}`} onClick={() => setSelected({ kind: "equipment", id: e.id })}>
                    <span className="row-dot" style={{ background: "#3a4650", border: "1px solid var(--line2)" }} />
                    <span className="row-tag">{e.tag}</span>
                    <span className="row-name">{e.name}</span>
                    {connCountFor(e.id) > 0 && <span className="pill">{connCountFor(e.id)}⇄</span>}
                  </div>
                ))}
                {project.equipment.length === 0 && <p className="hint">No equipment yet. Import a list, or use + EQUIPMENT in the toolbar.</p>}
              </>
            )}

            {leftTab === "utilities" && (
              <>
                <div className="sec">Facility utilities ({project.utilities.length})</div>
                {project.utilities.map(u => (
                  <div key={u.id} className={`row-item ${selected?.id === u.id ? "on" : ""}`} onClick={() => setSelected({ kind: "utility", id: u.id })}>
                    <span className="row-dot" style={{ background: UTILITY_TYPES[u.type]?.color, borderRadius: 6 }} />
                    <span className="row-tag">{u.tag}</span>
                    <span className="row-name">{u.name}</span>
                  </div>
                ))}
                {project.utilities.length === 0 && <p className="hint">No utility nodes yet. Use + UTILITY to place HVAC, chilled water, power, and other services.</p>}
                <div className="sec">Connections ({project.connections.length})</div>
                {project.connections.map(c => {
                  const a = nodes[c.fromId], b2 = nodes[c.toId];
                  return (
                    <div key={c.id} className={`conn-chip ${selected?.id === c.id ? "on" : ""}`} onClick={() => setSelected({ kind: "connection", id: c.id })}>
                      <span className="row-dot" style={{ background: UTILITY_TYPES[c.utilityType]?.color }} />
                      <span className="row-tag" style={{ fontSize: 10 }}>{a?.tag} → {b2?.tag}</span>
                      <span className="row-name">{c.medium || UTILITY_TYPES[c.utilityType]?.label}</span>
                    </div>
                  );
                })}
              </>
            )}
          </div>
        </div>

        {/* ---------- canvas ---------- */}
        <div className="canvas-wrap">
          {modeBanner && <div className="mode-banner">{modeBanner}</div>}

          {view === "2d" && (
            <svg
              ref={svgRef} className="canvas-svg"
              viewBox={`${cam.x} ${cam.y} ${cam.w} ${camH}`}
              preserveAspectRatio="none"
              onWheel={onWheel}
              onMouseDown={e => onCanvasDown(e, null)}
              onMouseMove={onCanvasMove}
              onMouseUp={onCanvasUp}
              onMouseLeave={onCanvasUp}
              style={{ cursor: mode === "select" && !placing ? "default" : "crosshair" }}
            >
              {/* grid */}
              <g stroke="var(--line)" strokeWidth={strokeW * 0.6} opacity="0.5">
                {gridLines.map((l, i) => l.v
                  ? <line key={i} x1={l.p} y1={cam.y - 5} x2={l.p} y2={cam.y + camH + 5} />
                  : <line key={i} x1={cam.x - 5} y1={l.p} x2={cam.x + cam.w + 5} y2={l.p} />)}
              </g>

              {/* floor plan */}
              {project.floorplan && (
                <g stroke="#6d7c88" fill="none" strokeWidth={strokeW * 1.4} strokeLinejoin="round">
                  {project.floorplan.entities.map((e, i) => {
                    if (e.kind === "line") return <line key={i} x1={e.pts[0][0]} y1={e.pts[0][1]} x2={e.pts[1][0]} y2={e.pts[1][1]} />;
                    if (e.kind === "poly") return <polyline key={i} points={(e.closed ? [...e.pts, e.pts[0]] : e.pts).map(p => p.join(",")).join(" ")} />;
                    if (e.kind === "circle") return <circle key={i} cx={e.cx} cy={e.cy} r={e.r} />;
                    if (e.kind === "text") return <text key={i} x={e.x} y={e.y} fontSize={e.h} fill="#55616b" stroke="none" fontFamily="var(--mono)" letterSpacing={e.h * 0.15}>{e.text}</text>;
                    return null;
                  })}
                </g>
              )}

              {/* connections */}
              <g fill="none">
                {project.connections.map(c => {
                  const a = nodes[c.fromId], b2 = nodes[c.toId];
                  if (!a || !b2 || a.x === undefined || b2.x === undefined) return null;
                  const col = UTILITY_TYPES[c.utilityType]?.color || "#888";
                  const on = selected?.id === c.id;
                  return (
                    <g key={c.id} onMouseDown={e => { e.stopPropagation(); setSelected({ kind: "connection", id: c.id }); }} style={{ cursor: "pointer" }}>
                      <path d={orthoPath(nodeAnchor(a), nodeAnchor(b2))} stroke="transparent" strokeWidth={strokeW * 12} />
                      <path d={orthoPath(nodeAnchor(a), nodeAnchor(b2))} stroke={col} strokeWidth={strokeW * (on ? 4 : 2)} strokeDasharray={c.utilityType === "power" ? `${strokeW * 8} ${strokeW * 5}` : "none"} opacity={on ? 1 : 0.75} />
                      <circle cx={b2.x} cy={b2.y} r={strokeW * 5} fill={col} opacity={on ? 1 : 0.75} />
                    </g>
                  );
                })}
              </g>

              {/* utility nodes */}
              {project.utilities.map(u => {
                const col = UTILITY_TYPES[u.type]?.color || "#888";
                const on = selected?.id === u.id || connectFrom === u.id;
                const r = 1.0;
                return (
                  <g key={u.id} onMouseDown={e => { e.stopPropagation(); onCanvasDown(e, { kind: "utility", id: u.id }); }} style={{ cursor: "pointer" }}>
                    <circle cx={u.x} cy={u.y} r={r} fill="var(--panel2)" stroke={col} strokeWidth={strokeW * (on ? 4 : 2.2)} />
                    <text x={u.x} y={u.y + 0.18} fontSize={0.55} textAnchor="middle" fill={col} fontFamily="var(--mono)" fontWeight="700">{UTILITY_TYPES[u.type]?.short}</text>
                    <text x={u.x} y={u.y - r - 0.35} fontSize={0.55} textAnchor="middle" fill="var(--text)" fontFamily="var(--mono)">{u.tag}</text>
                  </g>
                );
              })}

              {/* equipment */}
              {project.equipment.filter(e => e.placed).map(e => {
                const on = selected?.id === e.id || connectFrom === e.id;
                return (
                  <g key={e.id} onMouseDown={ev => { ev.stopPropagation(); onCanvasDown(ev, { kind: "equipment", id: e.id }); }} style={{ cursor: mode === "connect" ? "pointer" : "move" }}>
                    <rect x={e.x - e.w / 2} y={e.y - e.d / 2} width={e.w} height={e.d} rx={0.08}
                      fill={on ? "rgba(240,140,58,.18)" : "rgba(58,70,80,.55)"}
                      stroke={on ? "var(--accent)" : "#8b9aa6"} strokeWidth={strokeW * (on ? 3 : 1.6)} />
                    <line x1={e.x - e.w / 2} y1={e.y - e.d / 2} x2={e.x + e.w / 2} y2={e.y + e.d / 2} stroke={on ? "var(--accent)" : "#5c6a75"} strokeWidth={strokeW} />
                    <text x={e.x} y={e.y - e.d / 2 - 0.3} fontSize={0.6} textAnchor="middle" fill={on ? "var(--accent)" : "var(--text)"} fontFamily="var(--mono)" fontWeight="700">{e.tag}</text>
                  </g>
                );
              })}
            </svg>
          )}

          {view === "3d" && <div ref={threeRef} className="canvas-3d" />}

          {view === "2d" && !project.floorplan && project.equipment.length === 0 && (
            <div className="empty-canvas">
              <div className="big">NO FLOOR PLAN LOADED</div>
              <div>Import a DXF, or explore the demo facility.</div>
              <button className="btn" onClick={() => { setProject(demoProject()); setTimeout(() => fitView(), 0); }}>LOAD DEMO FACILITY</button>
            </div>
          )}

          {/* status bar */}
          <div className="statusbar">
            <div>X <b>{cursor.x.toFixed(2)}</b> Y <b>{cursor.y.toFixed(2)}</b> m</div>
            <div>GRID <b>{gridStep} m</b></div>
            <div>EQUIPMENT <b>{project.equipment.length}</b></div>
            <div>UTILITIES <b>{project.utilities.length}</b></div>
            <div>CONNECTIONS <b>{project.connections.length}</b></div>
            <div style={{ marginLeft: "auto", borderLeft: "1px solid var(--line)", borderRight: "none" }}>
              {project.floorplan ? `PLAN: ${project.floorplan.source}` : "PLAN: —"} · AUTOSAVE ON
            </div>
          </div>

          <div className="toasts">
            {toasts.map(t => <div key={t.id} className={`toast ${t.kind}`}>{t.msg}</div>)}
          </div>
        </div>

        {/* ---------- inspector ---------- */}
        <div className="inspector">
          <div className="side-tabs">
            <button className="side-tab on" style={{ cursor: "default" }}>
              {selEquip ? "Equipment" : selUtil ? "Utility" : selConn ? "Connection" : "Inspector"}
            </button>
          </div>
          <div className="side-body">
            {!selected && (
              <>
                <div className="sec">Nothing selected</div>
                <p className="hint">Select equipment, a utility node, or a connection line to edit its details here.</p>
                <p className="hint">Workflow: import a DXF plan → import or add equipment → place utility nodes → CONNECT to route HVAC, power, chilled water and other services → push snapshots to GitHub for version history.</p>
              </>
            )}

            {selEquip && (
              <EquipmentInspector eq={selEquip} patch={p => patchEquip(selEquip.id, p)}
                connections={project.connections.filter(c => c.fromId === selEquip.id || c.toId === selEquip.id)}
                nodes={nodes} onSelectConn={id => setSelected({ kind: "connection", id })}
                onRemove={removeSelected} />
            )}

            {selUtil && (
              <>
                <div className="sec">Utility node</div>
                <div className="fld-grid">
                  <Fld label="Tag"><input className="mono-in" value={selUtil.tag} onChange={e => patchUtil(selUtil.id, { tag: e.target.value })} /></Fld>
                  <Fld label="Service">
                    <select value={selUtil.type} onChange={e => patchUtil(selUtil.id, { type: e.target.value })}>
                      {Object.entries(UTILITY_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                    </select>
                  </Fld>
                </div>
                <Fld label="Name"><input value={selUtil.name} onChange={e => patchUtil(selUtil.id, { name: e.target.value })} /></Fld>
                <Fld label="Capacity / rating"><input value={selUtil.capacity || ""} placeholder="e.g. 250 kW, 12,000 CFM" onChange={e => patchUtil(selUtil.id, { capacity: e.target.value })} /></Fld>
                <div className="fld-grid">
                  <Fld label="X (m)"><input type="number" step="0.1" value={round2(selUtil.x)} onChange={e => patchUtil(selUtil.id, { x: parseFloat(e.target.value) || 0 })} /></Fld>
                  <Fld label="Y (m)"><input type="number" step="0.1" value={round2(selUtil.y)} onChange={e => patchUtil(selUtil.id, { y: parseFloat(e.target.value) || 0 })} /></Fld>
                </div>
                <div className="sec">Serves</div>
                {project.connections.filter(c => c.fromId === selUtil.id || c.toId === selUtil.id).map(c => {
                  const other = nodes[c.fromId === selUtil.id ? c.toId : c.fromId];
                  return (
                    <div key={c.id} className="conn-chip" onClick={() => setSelected({ kind: "connection", id: c.id })}>
                      <span className="row-dot" style={{ background: UTILITY_TYPES[c.utilityType]?.color }} />
                      <span className="row-tag" style={{ fontSize: 10 }}>{other?.tag}</span>
                      <span className="row-name">{c.medium || "—"}</span>
                    </div>
                  );
                })}
                <button className="btn danger" style={{ width: "100%", marginTop: 12 }} onClick={removeSelected}>DELETE UTILITY NODE</button>
              </>
            )}

            {selConn && (
              <>
                <div className="sec">Utility connection</div>
                <p className="hint" style={{ fontFamily: "var(--mono)" }}>
                  {nodes[selConn.fromId]?.tag} → {nodes[selConn.toId]?.tag}
                </p>
                <Fld label="Service">
                  <select value={selConn.utilityType} onChange={e => patchConn(selConn.id, { utilityType: e.target.value })}>
                    {Object.entries(UTILITY_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
                  </select>
                </Fld>
                <Fld label="Medium / description"><input value={selConn.medium} placeholder="e.g. Chilled water supply/return" onChange={e => patchConn(selConn.id, { medium: e.target.value })} /></Fld>
                <Fld label="Size / rating"><input value={selConn.size} placeholder="e.g. DN50 · 600x400 duct · 45 kW" onChange={e => patchConn(selConn.id, { size: e.target.value })} /></Fld>
                <Fld label="Notes"><textarea rows={3} value={selConn.notes} onChange={e => patchConn(selConn.id, { notes: e.target.value })} /></Fld>
                <button className="btn danger" style={{ width: "100%" }} onClick={removeSelected}>DELETE CONNECTION</button>
              </>
            )}
          </div>
        </div>
      </div>

      {/* ---------- GitHub modal ---------- */}
      {ghOpen && (
        <div className="modal-back" onMouseDown={e => { if (e.target === e.currentTarget) setGhOpen(false); }}>
          <div className="modal">
            <h3>GitHub repository</h3>
            <p className="sub">Version your digital twin by committing project snapshots (JSON) to a repo. Create a fine-grained personal access token with <b>Contents: read &amp; write</b> on the target repository. The token is kept only in this browser and is stripped from committed files and exports.</p>
            <div className="fld-grid">
              <Fld label="Owner"><input className="mono-in" placeholder="acme-eng" value={project.github.owner} onChange={e => setProject(p => ({ ...p, github: { ...p.github, owner: e.target.value.trim() } }))} /></Fld>
              <Fld label="Repository"><input className="mono-in" placeholder="plant-digital-twin" value={project.github.repo} onChange={e => setProject(p => ({ ...p, github: { ...p.github, repo: e.target.value.trim() } }))} /></Fld>
            </div>
            <div className="fld-grid">
              <Fld label="Branch"><input className="mono-in" value={project.github.branch} onChange={e => setProject(p => ({ ...p, github: { ...p.github, branch: e.target.value.trim() } }))} /></Fld>
              <Fld label="File path"><input className="mono-in" value={project.github.path} onChange={e => setProject(p => ({ ...p, github: { ...p.github, path: e.target.value.trim() } }))} /></Fld>
            </div>
            <Fld label="Personal access token"><input className="mono-in" type="password" placeholder="github_pat_…" value={project.github.token} onChange={e => setProject(p => ({ ...p, github: { ...p.github, token: e.target.value.trim() } }))} /></Fld>
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <button className="btn" disabled={ghBusy || !project.github.owner || !project.github.repo || !project.github.token} onClick={ghTest}>TEST CONNECTION</button>
              <button className="btn active" disabled={ghBusy || !project.github.owner || !project.github.repo || !project.github.token} onClick={ghPush}>{ghBusy ? "WORKING…" : "PUSH SNAPSHOT"}</button>
              <div style={{ flex: 1 }} />
              <button className="btn ghost" onClick={() => setGhOpen(false)}>CLOSE</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function round2(v) { return Math.round((v ?? 0) * 100) / 100; }

/* --------------------- equipment inspector ------------------------- */

function EquipmentInspector({ eq, patch, connections, nodes, onSelectConn, onRemove }) {
  const [tab, setTab] = useState("details");
  const setBom = bom => patch({ bom });
  const patchRow = (id, p) => setBom(eq.bom.map(r => r.id === id ? { ...r, ...p } : r));
  return (
    <>
      <div style={{ display: "flex", gap: 4, marginBottom: 12 }}>
        {["details", "bom", "utilities"].map(t => (
          <button key={t} className={`btn small ${tab === t ? "active" : "ghost"}`} onClick={() => setTab(t)}>
            {t === "bom" ? `BOM (${eq.bom.length})` : t === "utilities" ? `UTIL (${connections.length})` : "DETAILS"}
          </button>
        ))}
      </div>

      {tab === "details" && (
        <>
          <div className="fld-grid">
            <Fld label="Tag"><input className="mono-in" value={eq.tag} onChange={e => patch({ tag: e.target.value })} /></Fld>
            <Fld label="Type">
              <select value={eq.type} onChange={e => patch({ type: e.target.value })}>
                {[...new Set([eq.type, ...EQUIP_TYPES])].map(t => <option key={t}>{t}</option>)}
              </select>
            </Fld>
          </div>
          <Fld label="Name / description"><input value={eq.name} onChange={e => patch({ name: e.target.value })} /></Fld>
          <div className="fld-grid">
            <Fld label="Manufacturer"><input value={eq.manufacturer} onChange={e => patch({ manufacturer: e.target.value })} /></Fld>
            <Fld label="Model"><input value={eq.model} onChange={e => patch({ model: e.target.value })} /></Fld>
          </div>
          <div className="fld-grid">
            <Fld label="Power (kW)"><input type="number" value={eq.power} onChange={e => patch({ power: e.target.value === "" ? "" : parseFloat(e.target.value) })} /></Fld>
            <Fld label="Area / room"><input value={eq.area} onChange={e => patch({ area: e.target.value })} /></Fld>
          </div>
          <div className="sec">Footprint &amp; position (m)</div>
          <div className="fld-grid3">
            <Fld label="Width"><input type="number" step="0.1" value={eq.w} onChange={e => patch({ w: parseFloat(e.target.value) || 0.2 })} /></Fld>
            <Fld label="Depth"><input type="number" step="0.1" value={eq.d} onChange={e => patch({ d: parseFloat(e.target.value) || 0.2 })} /></Fld>
            <Fld label="Height"><input type="number" step="0.1" value={eq.h} onChange={e => patch({ h: parseFloat(e.target.value) || 0.2 })} /></Fld>
          </div>
          <div className="fld-grid">
            <Fld label="X"><input type="number" step="0.1" value={round2(eq.x)} onChange={e => patch({ x: parseFloat(e.target.value) || 0, placed: true })} /></Fld>
            <Fld label="Y"><input type="number" step="0.1" value={round2(eq.y)} onChange={e => patch({ y: parseFloat(e.target.value) || 0, placed: true })} /></Fld>
          </div>
          <button className="btn danger" style={{ width: "100%", marginTop: 8 }} onClick={onRemove}>DELETE EQUIPMENT</button>
        </>
      )}

      {tab === "bom" && (
        <>
          <div className="sec">Bill of materials — {eq.tag}</div>
          <table className="bom-table">
            <thead><tr><th style={{ width: "34%" }}>Part no.</th><th>Description</th><th style={{ width: 42 }}>Qty</th><th style={{ width: 22 }} /></tr></thead>
            <tbody>
              {eq.bom.map(r => (
                <tr key={r.id}>
                  <td><input style={{ fontFamily: "var(--mono)" }} value={r.pn} onChange={e => patchRow(r.id, { pn: e.target.value })} /></td>
                  <td><input value={r.desc} onChange={e => patchRow(r.id, { desc: e.target.value })} /></td>
                  <td><input type="number" value={r.qty} onChange={e => patchRow(r.id, { qty: parseFloat(e.target.value) || 0 })} /></td>
                  <td><button className="btn small ghost danger" style={{ padding: "1px 5px" }} onClick={() => setBom(eq.bom.filter(x => x.id !== r.id))}>✕</button></td>
                </tr>
              ))}
            </tbody>
          </table>
          {eq.bom.length === 0 && <p className="hint">No parts recorded for this equipment.</p>}
          <button className="btn small" style={{ marginTop: 8 }} onClick={() => setBom([...eq.bom, { id: uid(), pn: "", desc: "", qty: 1 }])}>+ ADD PART</button>
          <button className="btn small ghost" style={{ marginTop: 8, marginLeft: 6 }} disabled={!eq.bom.length}
            onClick={() => downloadFile(`${eq.tag}-BOM.csv`, "part_no,description,qty\n" + eq.bom.map(r => `"${r.pn}","${r.desc}",${r.qty}`).join("\n"), "text/csv")}>
            EXPORT CSV
          </button>
        </>
      )}

      {tab === "utilities" && (
        <>
          <div className="sec">Utility connections</div>
          {connections.map(c => {
            const other = nodes[c.fromId === eq.id ? c.toId : c.fromId];
            const dir = c.fromId === eq.id ? "→" : "←";
            return (
              <div key={c.id} className="conn-chip" onClick={() => onSelectConn(c.id)}>
                <span className="row-dot" style={{ background: UTILITY_TYPES[c.utilityType]?.color }} />
                <span className="row-tag" style={{ fontSize: 10 }}>{dir} {other?.tag}</span>
                <span className="row-name">{c.medium || UTILITY_TYPES[c.utilityType]?.label}{c.size ? ` · ${c.size}` : ""}</span>
              </div>
            );
          })}
          {connections.length === 0 && <p className="hint">Not connected to any utility. Use ⇄ CONNECT in the toolbar, click this equipment, then click a utility node (or another equipment for process lines).</p>}
        </>
      )}
    </>
  );
}
