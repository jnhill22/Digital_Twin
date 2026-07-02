import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as THREE from "three";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { LineChart, Line, AreaChart, Area, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from "recharts";

/* ================================================================== */
/*  FacilityTwin v2 — Industrial Digital Twin Platform                 */
/*  Real-time sensors · Predictive maintenance · Anomaly detection    */
/*  Analytics dashboard · DXF plans · Equipment BOM · Utility routing */
/*  2D/3D views · GitHub version control · Persistent storage         */
/* ================================================================== */

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

const SEVERITY = { ok: { label: "Normal", color: "#62D26F" }, warn: { label: "Warning", color: "#EFCB4F" }, critical: { label: "Critical", color: "#E06552" }, offline: { label: "Offline", color: "#55616B" } };

const uid = () => Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-3);

/* ====================== SENSOR SIMULATION ENGINE =================== */

const SENSOR_PROFILES = {
  Process:            { temp: [55, 85, 95],  vibration: [0.5, 4.5, 7],  pressure: [1, 6, 8],   current: [20, 80, 100], rpm: [800, 1800, 2200] },
  Packaging:          { temp: [25, 45, 55],  vibration: [0.2, 3, 5],    pressure: [0.5, 4, 6],  current: [15, 60, 85],  rpm: [400, 1200, 1600] },
  "Material Handling": { temp: [20, 40, 50], vibration: [0.3, 5, 8],    pressure: [0, 0, 0],    current: [10, 50, 70],  rpm: [0, 0, 0] },
  "Storage Tank":     { temp: [2, 25, 35],   vibration: [0, 0.5, 1],    pressure: [0.2, 2, 3],  current: [0, 5, 10],    rpm: [0, 0, 0] },
  Pump:               { temp: [30, 65, 80],  vibration: [0.5, 6, 9],    pressure: [2, 8, 12],   current: [25, 90, 110], rpm: [1400, 3000, 3600] },
  Skid:               { temp: [30, 60, 75],  vibration: [0.3, 3.5, 6],  pressure: [1, 5, 7],    current: [20, 70, 95],  rpm: [600, 1500, 2000] },
  Instrument:         { temp: [18, 30, 40],  vibration: [0, 0.2, 0.5],  pressure: [0, 0, 0],    current: [1, 5, 8],     rpm: [0, 0, 0] },
  Other:              { temp: [20, 50, 65],  vibration: [0.2, 3, 5],    pressure: [0, 3, 5],    current: [5, 40, 60],   rpm: [0, 0, 0] },
};

const SENSOR_UNITS = { temp: "°C", vibration: "mm/s", pressure: "bar", current: "A", rpm: "RPM" };
const SENSOR_LABELS = { temp: "Temperature", vibration: "Vibration", pressure: "Pressure", current: "Current Draw", rpm: "Motor Speed" };

function generateSensorReading(profile, key, healthFactor, t) {
  const [, nominal, max] = profile[key];
  if (nominal === 0 && max === 0) return null;
  const degradation = 1 + (1 - healthFactor) * 0.6;
  const base = nominal * degradation;
  const noise = Math.sin(t * 0.7 + key.length) * nominal * 0.08 + Math.sin(t * 2.3) * nominal * 0.03;
  const spike = Math.random() < (1 - healthFactor) * 0.05 ? nominal * 0.3 * (1 - healthFactor) : 0;
  return Math.max(0, base + noise + spike);
}

function getSeverity(value, profile, key) {
  if (value === null) return "ok";
  const [, nominal, max] = profile[key];
  if (value >= max) return "critical";
  if (value >= nominal + (max - nominal) * 0.6) return "warn";
  return "ok";
}

function computeHealth(sensors, profile) {
  let worst = 1;
  for (const key of Object.keys(sensors)) {
    if (sensors[key] === null) continue;
    const [min, nominal, max] = profile[key];
    const v = sensors[key];
    const ratio = Math.max(0, Math.min(1, (v - min) / (max - min)));
    worst = Math.min(worst, 1 - ratio);
  }
  return Math.max(0, Math.min(100, Math.round(worst * 100)));
}

function predictFailureDays(health, trend) {
  if (health > 85) return null;
  const rate = Math.max(0.1, -trend || 0.3);
  return Math.max(1, Math.round(health / rate));
}

function generateMaintenanceAlerts(equipment, sensorData) {
  const alerts = [];
  for (const eq of equipment) {
    const sd = sensorData[eq.id];
    if (!sd) continue;
    const profile = SENSOR_PROFILES[eq.type] || SENSOR_PROFILES.Other;
    for (const key of Object.keys(sd.current)) {
      const sev = getSeverity(sd.current[key], profile, key);
      if (sev === "critical") {
        alerts.push({ id: uid(), equipId: eq.id, tag: eq.tag, sensor: key, severity: "critical", value: sd.current[key], message: `${SENSOR_LABELS[key]} critically high on ${eq.tag} (${sd.current[key]?.toFixed(1)} ${SENSOR_UNITS[key]})`, ts: Date.now() });
      } else if (sev === "warn") {
        alerts.push({ id: uid(), equipId: eq.id, tag: eq.tag, sensor: key, severity: "warn", value: sd.current[key], message: `${SENSOR_LABELS[key]} elevated on ${eq.tag} (${sd.current[key]?.toFixed(1)} ${SENSOR_UNITS[key]})`, ts: Date.now() });
      }
    }
    if (sd.health < 40) {
      alerts.push({ id: uid(), equipId: eq.id, tag: eq.tag, sensor: "health", severity: "critical", value: sd.health, message: `${eq.tag} health critically low (${sd.health}%) — schedule maintenance`, ts: Date.now() });
    }
  }
  return alerts.sort((a, b) => (a.severity === "critical" ? 0 : 1) - (b.severity === "critical" ? 0 : 1));
}

/* ========================== DXF PARSER ============================== */

function parseDXF(text) {
  const rows = text.split(/\r\n|\r|\n/);
  const pairs = [];
  for (let i = 0; i + 1 < rows.length; i += 2) {
    const code = parseInt(rows[i], 10);
    if (!Number.isNaN(code)) pairs.push([code, rows[i + 1].trim()]);
  }
  let section = "", cur = null;
  const raw = [];
  for (let i = 0; i < pairs.length; i++) {
    const [c, v] = pairs[i];
    if (c === 0 && v === "SECTION") { const nx = pairs[i + 1]; section = nx && nx[0] === 2 ? nx[1] : ""; continue; }
    if (c === 0 && (v === "ENDSEC" || v === "EOF")) { if (cur) raw.push(cur); cur = null; section = ""; continue; }
    if (section !== "ENTITIES") continue;
    if (c === 0) { if (cur) raw.push(cur); cur = { type: v, verts: [], codes: {} }; continue; }
    if (!cur) continue;
    if (c === 10) cur.verts.push({ x: parseFloat(v), y: 0 });
    else if (c === 20) { const lv = cur.verts[cur.verts.length - 1]; if (lv) lv.y = parseFloat(v); }
    else if (!(c in cur.codes)) cur.codes[c] = v;
  }
  if (cur) raw.push(cur);
  const ents = [];
  for (const e of raw) {
    const c = e.codes;
    if (e.type === "LINE" && e.verts[0] && c[11] !== undefined && c[21] !== undefined)
      ents.push({ kind: "line", pts: [[e.verts[0].x, e.verts[0].y], [parseFloat(c[11]), parseFloat(c[21])]] });
    else if ((e.type === "LWPOLYLINE" || e.type === "POLYLINE") && e.verts.length > 1)
      ents.push({ kind: "poly", pts: e.verts.map(p => [p.x, p.y]), closed: (parseInt(c[70] || "0", 10) & 1) === 1 });
    else if (e.type === "CIRCLE" && e.verts[0] && c[40])
      ents.push({ kind: "circle", cx: e.verts[0].x, cy: e.verts[0].y, r: parseFloat(c[40]) });
    else if (e.type === "ARC" && e.verts[0] && c[40]) {
      const cx = e.verts[0].x, cy = e.verts[0].y, r = parseFloat(c[40]);
      let a0 = (parseFloat(c[50] || "0") * Math.PI) / 180, a1 = (parseFloat(c[51] || "360") * Math.PI) / 180;
      if (a1 <= a0) a1 += Math.PI * 2;
      const n = Math.max(8, Math.ceil(((a1 - a0) / (Math.PI * 2)) * 32)), pts = [];
      for (let k = 0; k <= n; k++) { const a = a0 + ((a1 - a0) * k) / n; pts.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]); }
      ents.push({ kind: "poly", pts, closed: false });
    } else if ((e.type === "TEXT" || e.type === "MTEXT") && e.verts[0] && c[1])
      ents.push({ kind: "text", x: e.verts[0].x, y: e.verts[0].y, h: parseFloat(c[40] || "0.3"), text: String(c[1]).replace(/\\P/g, " ").replace(/\{|\}|\\[A-Za-z][^;]*;/g, "") });
  }
  for (const e of ents) {
    if (e.pts) e.pts = e.pts.map(([x, y]) => [x, -y]);
    if (e.cy !== undefined) e.cy = -e.cy;
    if (e.y !== undefined && e.kind === "text") e.y = -e.y;
  }
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  const grow = (x, y) => { minX = Math.min(minX, x); minY = Math.min(minY, y); maxX = Math.max(maxX, x); maxY = Math.max(maxY, y); };
  for (const e of ents) {
    if (e.pts) e.pts.forEach(([x, y]) => grow(x, y));
    if (e.kind === "circle") { grow(e.cx - e.r, e.cy - e.r); grow(e.cx + e.r, e.cy + e.r); }
    if (e.kind === "text") grow(e.x, e.y);
  }
  if (!isFinite(minX)) { minX = 0; minY = 0; maxX = 40; maxY = 25; }
  let unitScale = 1, assumedUnits = "m";
  const span = Math.max(maxX - minX, maxY - minY);
  if (span > 2000) { unitScale = 0.001; assumedUnits = "mm"; } else if (span > 500) { unitScale = 0.01; assumedUnits = "cm"; }
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

/* ===================== EQUIPMENT LIST IMPORT ======================== */

const COL_ALIASES = {
  tag: ["tag", "equipment tag", "eq tag", "equipment id", "id", "item", "item no", "item number", "asset tag"],
  name: ["name", "description", "equipment name", "equipment description", "desc", "title"],
  type: ["type", "category", "equipment type", "class", "classification"],
  manufacturer: ["manufacturer", "mfr", "mfg", "vendor", "make", "supplier", "oem"],
  model: ["model", "model no", "model number", "model #"],
  power: ["power", "kw", "power (kw)", "power kw", "rated power", "load", "load (kw)", "hp"],
  x: ["x", "loc x", "x (m)", "pos x", "x coord"], y: ["y", "loc y", "y (m)", "pos y", "y coord"],
  w: ["width", "w", "width (m)", "w (m)"], d: ["depth", "d", "length", "l", "depth (m)", "length (m)"],
  h: ["height", "h", "height (m)", "h (m)"], area: ["area", "room", "zone", "location", "space"],
};

function mapColumns(headers) {
  const map = {};
  const norm = headers.map(h => String(h || "").trim().toLowerCase());
  for (const key of Object.keys(COL_ALIASES)) { const idx = norm.findIndex(h => COL_ALIASES[key].includes(h)); if (idx >= 0) map[key] = headers[idx]; }
  return map;
}

function rowsToEquipment(rows) {
  if (!rows.length) return [];
  const map = mapColumns(Object.keys(rows[0]));
  return rows.filter(r => Object.values(r).some(v => String(v || "").trim() !== "")).map((r, i) => {
    const num = k => { const v = parseFloat(r[map[k]]); return Number.isFinite(v) ? v : undefined; };
    const x = num("x"), y = num("y");
    return { id: uid(), tag: String(r[map.tag] ?? `EQ-${String(i + 1).padStart(3, "0")}`).trim(),
      name: String(r[map.name] ?? "").trim() || "Unnamed equipment", type: String(r[map.type] ?? "Process").trim(),
      manufacturer: String(r[map.manufacturer] ?? "").trim(), model: String(r[map.model] ?? "").trim(),
      power: num("power") ?? "", area: String(r[map.area] ?? "").trim(),
      x, y, placed: x !== undefined && y !== undefined, w: num("w") ?? 2, d: num("d") ?? 1.5, h: num("h") ?? 1.8, bom: [],
      installDate: "", lastMaintenance: "", maintenanceLog: [], operatingHours: Math.floor(Math.random() * 8000 + 1000),
    };
  });
}

function downloadFile(name, text, mime = "application/json") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = name; a.click(); URL.revokeObjectURL(url);
}

/* ============================ DEMO DATA ============================= */

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
  ];
  const eq = (o) => ({ id: uid(), placed: true, bom: [], w: 2, d: 1.5, h: 1.8, power: "", area: "", manufacturer: "", model: "",
    installDate: "", lastMaintenance: "", maintenanceLog: [], operatingHours: 0, ...o });
  const equipment = [
    eq({ tag: "MX-101", name: "High-shear batch mixer", type: "Process", manufacturer: "Silverson", model: "AX-450", power: 45, x: 6, y: 5, w: 2.4, d: 2.4, h: 2.6, area: "Process Hall A", operatingHours: 12400, installDate: "2021-03-15", lastMaintenance: "2026-05-20",
      bom: [{ id: uid(), pn: "AX450-SEAL-KIT", desc: "Mechanical seal kit", qty: 1 }, { id: uid(), pn: "MTR-45KW-IE3", desc: "45 kW IE3 motor", qty: 1 }, { id: uid(), pn: "IMP-HS-300", desc: "High-shear impeller 300mm", qty: 1 }],
      maintenanceLog: [{ id: uid(), date: "2026-05-20", type: "Preventive", desc: "Seal replacement, bearing inspection", cost: 2400 }, { id: uid(), date: "2025-11-10", type: "Corrective", desc: "Motor overhaul — bearing failure", cost: 8500 }] }),
    eq({ tag: "R-201", name: "Jacketed reactor 2000L", type: "Storage Tank", manufacturer: "De Dietrich", model: "AE-2000", power: 15, x: 12, y: 5, w: 2.8, d: 2.8, h: 3.4, area: "Process Hall A", operatingHours: 9800, installDate: "2022-01-10", lastMaintenance: "2026-04-15",
      bom: [{ id: uid(), pn: "AGIT-PBT-2000", desc: "Pitched-blade agitator", qty: 1 }, { id: uid(), pn: "RV-DN80", desc: "Bottom outlet valve DN80", qty: 1 }],
      maintenanceLog: [{ id: uid(), date: "2026-04-15", type: "Preventive", desc: "Agitator seal, jacket pressure test", cost: 1800 }] }),
    eq({ tag: "P-301", name: "Transfer pump", type: "Pump", manufacturer: "Alfa Laval", model: "LKH-25", power: 7.5, x: 12, y: 8.5, w: 1, d: 0.8, h: 0.9, area: "Process Hall A", operatingHours: 15200, installDate: "2020-07-22", lastMaintenance: "2026-06-01",
      maintenanceLog: [{ id: uid(), date: "2026-06-01", type: "Preventive", desc: "Impeller inspection, mechanical seal", cost: 950 }, { id: uid(), date: "2026-01-15", type: "Corrective", desc: "Cavitation damage repair", cost: 3200 }] }),
    eq({ tag: "CIP-01", name: "CIP skid, 3-tank", type: "Skid", manufacturer: "GEA", model: "Compact CIP", power: 22, x: 21, y: 13, w: 4, d: 2, h: 2.4, area: "Process Hall B", operatingHours: 8600, installDate: "2022-06-30", lastMaintenance: "2026-03-28",
      bom: [{ id: uid(), pn: "TK-CIP-500", desc: "500L caustic tank", qty: 3 }, { id: uid(), pn: "HX-PHE-30", desc: "Plate heat exchanger", qty: 1 }],
      maintenanceLog: [{ id: uid(), date: "2026-03-28", type: "Preventive", desc: "Heat exchanger descaling, valve overhaul", cost: 3100 }] }),
    eq({ tag: "FIL-401", name: "Rotary filler, 12-head", type: "Packaging", manufacturer: "Krones", model: "Modulfill", power: 30, x: 35, y: 22, w: 3.5, d: 3.5, h: 2.8, area: "Packaging", operatingHours: 6200, installDate: "2023-09-01", lastMaintenance: "2026-06-10",
      maintenanceLog: [{ id: uid(), date: "2026-06-10", type: "Preventive", desc: "Fill valve calibration, servo check", cost: 1500 }] }),
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
    name: "Line 4 — Beverage Processing", site: "Fresno Plant", revision: "B",
    floorplan: { entities, bounds: { minX: 0, minY: 0, maxX: 42, maxY: 26 }, assumedUnits: "m", source: "demo-floorplan.dxf" },
    equipment, utilities, connections,
    github: { owner: "", repo: "", branch: "main", path: "digital-twin/project.json", token: "" },
  };
}

function emptyProject() {
  return { name: "Untitled facility", site: "", revision: "A", floorplan: null,
    equipment: [], utilities: [], connections: [],
    github: { owner: "", repo: "", branch: "main", path: "digital-twin/project.json", token: "" },
  };
}

/* ============================== STYLES ============================== */

const CSS = `
:root{
  --bg:#101418;--panel:#171C21;--panel2:#1D242B;--line:#2A333C;--line2:#38434E;
  --text:#DCE3E9;--muted:#7E8B95;--faint:#55616B;--accent:#F08C3A;--accent-dim:#7a4a22;
  --danger:#E06552;--ok:#62D26F;--warn:#EFCB4F;
  --mono:ui-monospace,'SF Mono','Cascadia Code',Consolas,monospace;
  --sans:'Segoe UI',system-ui,-apple-system,Roboto,sans-serif;
}
.ft *{box-sizing:border-box}
.ft{position:fixed;inset:0;background:var(--bg);color:var(--text);font-family:var(--sans);font-size:13px;display:flex;flex-direction:column;overflow:hidden}
.ft ::-webkit-scrollbar{width:7px;height:7px}
.ft ::-webkit-scrollbar-thumb{background:var(--line2);border-radius:4px}
.ft ::-webkit-scrollbar-track{background:transparent}

.tb{display:flex;align-items:stretch;border-bottom:1px solid var(--line);background:var(--panel);height:48px;flex:none}
.tb-block{display:flex;flex-direction:column;justify-content:center;padding:0 14px;border-right:1px solid var(--line)}
.tb-ey{font-family:var(--mono);font-size:9px;letter-spacing:.14em;color:var(--faint);text-transform:uppercase}
.tb-tl{font-weight:600;font-size:14px;white-space:nowrap}
.tb-tl input{background:transparent;border:none;color:var(--text);font:inherit;outline:none;width:200px}
.tb-sp{flex:1}
.tb-g{display:flex;align-items:center;gap:5px;padding:0 10px;border-left:1px solid var(--line)}

.btn{font-family:var(--mono);font-size:10px;letter-spacing:.04em;padding:5px 10px;border:1px solid var(--line2);background:var(--panel2);color:var(--text);border-radius:3px;cursor:pointer;white-space:nowrap}
.btn:hover{border-color:var(--accent);color:#fff}
.btn:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
.btn.on{background:var(--accent);border-color:var(--accent);color:#14181c;font-weight:700}
.btn.gh{background:transparent}
.btn.sm{padding:3px 7px;font-size:9px}
.btn.dng:hover{border-color:var(--danger);color:var(--danger)}
.btn:disabled{opacity:.35;cursor:default}

.main{flex:1;display:flex;min-height:0}
.side{width:240px;flex:none;background:var(--panel);border-right:1px solid var(--line);display:flex;flex-direction:column;min-height:0}
.insp{width:290px;flex:none;background:var(--panel);border-left:1px solid var(--line);display:flex;flex-direction:column;min-height:0}
.stabs{display:flex;border-bottom:1px solid var(--line);flex:none}
.stab{flex:1;padding:7px 4px;text-align:center;font-family:var(--mono);font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);cursor:pointer;border:none;border-bottom:2px solid transparent;background:none}
.stab.on{color:var(--text);border-bottom-color:var(--accent)}
.sbody{flex:1;overflow-y:auto;padding:10px}

.sec{font-family:var(--mono);font-size:9px;letter-spacing:.14em;text-transform:uppercase;color:var(--faint);margin:12px 0 5px;display:flex;align-items:center;gap:8px}
.sec:first-child{margin-top:0}
.sec::after{content:"";flex:1;height:1px;background:var(--line)}
.hint{color:var(--muted);font-size:11px;line-height:1.5;margin:5px 0}

.ri{display:flex;align-items:center;gap:7px;padding:5px 7px;border:1px solid transparent;border-radius:3px;cursor:pointer;margin-bottom:2px}
.ri:hover{background:var(--panel2)}
.ri.on{background:var(--panel2);border-color:var(--accent-dim)}
.rdot{width:8px;height:8px;border-radius:2px;flex:none}
.rtag{font-family:var(--mono);font-size:11px;font-weight:700;white-space:nowrap}
.rname{font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1}
.pill{font-family:var(--mono);font-size:9px;padding:1px 5px;border-radius:2px;border:1px solid var(--line2);color:var(--muted);flex:none}

.cwrap{flex:1;position:relative;min-width:0;background:radial-gradient(circle at 30% 20%,#131920 0%,var(--bg) 60%)}
.csvg{position:absolute;inset:0;width:100%;height:100%;display:block}
.c3d{position:absolute;inset:0}
.sbar{position:absolute;left:0;right:0;bottom:0;height:24px;display:flex;align-items:center;background:rgba(16,20,24,.92);border-top:1px solid var(--line);font-family:var(--mono);font-size:10px;color:var(--muted);z-index:5}
.sbar>div{padding:0 10px;border-right:1px solid var(--line);height:100%;display:flex;align-items:center;gap:5px;white-space:nowrap}
.sbar b{color:var(--text);font-weight:600}
.mbanner{position:absolute;top:10px;left:50%;transform:translateX(-50%);background:var(--accent);color:#14181c;font-family:var(--mono);font-size:10px;font-weight:700;padding:4px 12px;border-radius:3px;z-index:6}

.fld{margin-bottom:7px}
.fld label{display:block;font-family:var(--mono);font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--faint);margin-bottom:2px}
.fld input,.fld select,.fld textarea{width:100%;background:var(--bg);border:1px solid var(--line2);color:var(--text);border-radius:3px;padding:5px 7px;font-size:12px;font-family:var(--sans);outline:none}
.fld input:focus,.fld select:focus,.fld textarea:focus{border-color:var(--accent)}
.fg2{display:grid;grid-template-columns:1fr 1fr;gap:7px}
.fg3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:7px}

.bt{width:100%;border-collapse:collapse;font-size:11px}
.bt th{font-family:var(--mono);font-size:9px;letter-spacing:.06em;text-transform:uppercase;color:var(--faint);text-align:left;padding:3px 5px;border-bottom:1px solid var(--line)}
.bt td{padding:3px 4px;border-bottom:1px solid var(--line)}
.bt input{width:100%;background:transparent;border:none;color:var(--text);font-size:11px;outline:none;padding:2px}
.bt input:focus{background:var(--bg)}

.cc{display:flex;align-items:center;gap:7px;padding:5px 7px;border:1px solid var(--line);border-radius:3px;margin-bottom:3px;cursor:pointer}
.cc:hover{border-color:var(--line2)}.cc.on{border-color:var(--accent)}

.mback{position:fixed;inset:0;background:rgba(8,10,12,.7);display:flex;align-items:center;justify-content:center;z-index:50}
.mdl{width:440px;max-width:92vw;max-height:86vh;overflow-y:auto;background:var(--panel);border:1px solid var(--line2);border-radius:6px;padding:16px}
.mdl h3{margin:0 0 4px;font-size:15px}
.mdl .sub{color:var(--muted);font-size:11px;margin-bottom:12px;line-height:1.5}

.toasts{position:absolute;bottom:34px;right:10px;display:flex;flex-direction:column;gap:5px;z-index:20}
.toast{background:var(--panel2);border:1px solid var(--line2);border-left:3px solid var(--accent);padding:7px 10px;border-radius:3px;font-size:11px;max-width:280px;box-shadow:0 4px 16px rgba(0,0,0,.4)}
.toast.err{border-left-color:var(--danger)}.toast.ok{border-left-color:var(--ok)}

.empty{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;color:var(--muted);z-index:2;pointer-events:none}
.empty .big{font-family:var(--mono);font-size:12px;letter-spacing:.12em;color:var(--faint)}
.empty button{pointer-events:auto}

/* ---- dashboard ---- */
.dash{flex:1;overflow-y:auto;padding:20px 24px;display:flex;flex-direction:column;gap:16px}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:4px;padding:14px 16px}
.kpi-label{font-family:var(--mono);font-size:9px;letter-spacing:.12em;text-transform:uppercase;color:var(--faint);margin-bottom:4px}
.kpi-val{font-size:28px;font-weight:700;font-family:var(--mono);letter-spacing:-.02em}
.kpi-sub{font-size:10px;color:var(--muted);margin-top:2px}

.card{background:var(--panel);border:1px solid var(--line);border-radius:4px;padding:14px 16px}
.card-h{font-family:var(--mono);font-size:10px;letter-spacing:.1em;text-transform:uppercase;color:var(--faint);margin-bottom:10px;display:flex;align-items:center;justify-content:space-between}

.ht-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:8px}
.ht-item{display:flex;align-items:center;gap:10px;padding:8px 10px;background:var(--panel2);border-radius:3px;border:1px solid var(--line);cursor:pointer}
.ht-item:hover{border-color:var(--line2)}
.ht-bar{width:48px;height:6px;border-radius:3px;background:var(--line);overflow:hidden}
.ht-fill{height:100%;border-radius:3px}

.alert-row{display:flex;align-items:flex-start;gap:8px;padding:6px 0;border-bottom:1px solid var(--line);font-size:11px}
.alert-dot{width:8px;height:8px;border-radius:4px;flex:none;margin-top:3px}
.alert-msg{flex:1;color:var(--text);line-height:1.4}
.alert-ts{font-family:var(--mono);font-size:9px;color:var(--faint);flex:none}

.gauge{position:relative;width:90px;height:90px}
.gauge-ring{fill:none;stroke:var(--line);stroke-width:6}
.gauge-val{fill:none;stroke-width:6;stroke-linecap:round;transition:stroke-dashoffset .6s}
.gauge-text{font-family:var(--mono);font-size:18px;font-weight:700;fill:var(--text);text-anchor:middle;dominant-baseline:central}
.gauge-label{font-family:var(--mono);font-size:8px;fill:var(--faint);text-anchor:middle;letter-spacing:.08em}

.sensor-row{display:flex;align-items:center;gap:10px;padding:5px 0;border-bottom:1px solid var(--line)}
.sensor-name{font-size:11px;color:var(--muted);width:90px;flex:none}
.sensor-val{font-family:var(--mono);font-size:13px;font-weight:600;width:70px;flex:none}
.sensor-bar{flex:1;height:5px;background:var(--line);border-radius:3px;overflow:hidden}
.sensor-fill{height:100%;border-radius:3px;transition:width .3s}
`;

/* ============================ HELPERS =============================== */

function Fld({ label, children }) { return <div className="fld"><label>{label}</label>{children}</div>; }
function nodeAnchor(n) { return { x: n.x ?? 0, y: n.y ?? 0 }; }
function orthoPath(a, b) { return `M ${a.x} ${a.y} L ${a.x} ${b.y} L ${b.x} ${b.y}`; }
function round2(v) { return Math.round((v ?? 0) * 100) / 100; }

function HealthGauge({ value, size = 90 }) {
  const r = (size - 12) / 2, circ = 2 * Math.PI * r;
  const pct = Math.max(0, Math.min(100, value));
  const color = pct > 70 ? "var(--ok)" : pct > 40 ? "var(--warn)" : "var(--danger)";
  return (
    <svg width={size} height={size} className="gauge">
      <circle cx={size/2} cy={size/2} r={r} className="gauge-ring" />
      <circle cx={size/2} cy={size/2} r={r} className="gauge-val" stroke={color}
        strokeDasharray={circ} strokeDashoffset={circ * (1 - pct / 100)}
        transform={`rotate(-90 ${size/2} ${size/2})`} />
      <text x={size/2} y={size/2 - 4} className="gauge-text">{pct}</text>
      <text x={size/2} y={size/2 + 14} className="gauge-label">HEALTH %</text>
    </svg>
  );
}

function SensorBar({ label, value, unit, min, nominal, max }) {
  if (value === null) return null;
  const pct = Math.min(100, Math.max(0, ((value - min) / (max - min)) * 100));
  const color = pct > 85 ? "var(--danger)" : pct > 65 ? "var(--warn)" : "var(--ok)";
  return (
    <div className="sensor-row">
      <span className="sensor-name">{label}</span>
      <span className="sensor-val" style={{ color }}>{value.toFixed(1)} <span style={{ fontSize: 9, color: "var(--faint)" }}>{unit}</span></span>
      <div className="sensor-bar"><div className="sensor-fill" style={{ width: pct + "%", background: color }} /></div>
    </div>
  );
}

/* ============================== APP ================================= */

export default function FacilityTwin() {
  const [project, setProject] = useState(emptyProject);
  const [selected, setSelected] = useState(null);
  const [mode, setMode] = useState("select");
  const [pendingUtilType, setPendingUtilType] = useState("hvac");
  const [connectFrom, setConnectFrom] = useState(null);
  const [topView, setTopView] = useState("plan");   // plan | dashboard | analytics
  const [planMode, setPlanMode] = useState("2d");
  const [leftTab, setLeftTab] = useState("import");
  const [ghOpen, setGhOpen] = useState(false);
  const [ghBusy, setGhBusy] = useState(false);
  const [toasts, setToasts] = useState([]);
  const [cursor, setCursor] = useState({ x: 0, y: 0 });
  const [cam, setCam] = useState({ x: -2, y: -2, w: 50 });
  const [placing, setPlacing] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [aspect, setAspect] = useState(0.62);

  // sensor simulation state
  const [sensorData, setSensorData] = useState({});
  const [sensorHistory, setSensorHistory] = useState({});
  const [alerts, setAlerts] = useState([]);
  const [simTime, setSimTime] = useState(0);
  const [simRunning, setSimRunning] = useState(true);

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

  /* ---- aspect measurement ---- */
  useEffect(() => {
    const measure = () => { const el = svgRef.current; if (el) { const r = el.getBoundingClientRect(); if (r.width > 0) setAspect(r.height / r.width); } };
    measure(); window.addEventListener("resize", measure);
    const t = setInterval(measure, 1500);
    return () => { window.removeEventListener("resize", measure); clearInterval(t); };
  }, [topView, planMode]);

  /* ---- persistence ---- */
  useEffect(() => {
    try {
      const saved = localStorage.getItem("facility-twin-v2:project");
      if (saved) { setProject(JSON.parse(saved)); toast("Restored saved project", "ok"); }
    } catch {}
    setLoaded(true);
  }, []);

  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => {
      try { localStorage.setItem("facility-twin-v2:project", JSON.stringify(project)); } catch {}
    }, 1200);
    return () => clearTimeout(t);
  }, [project, loaded]);

  /* ===================== SENSOR SIMULATION LOOP ==================== */
  useEffect(() => {
    if (!simRunning || project.equipment.length === 0) return;
    const iv = setInterval(() => {
      setSimTime(t => t + 1);
    }, 2000);
    return () => clearInterval(iv);
  }, [simRunning, project.equipment.length]);

  useEffect(() => {
    if (project.equipment.length === 0) return;
    const newData = {};
    for (const eq of project.equipment) {
      const profile = SENSOR_PROFILES[eq.type] || SENSOR_PROFILES.Other;
      // degradation factor based on operating hours (more hours = lower health baseline)
      const ageFactor = Math.max(0.55, 1 - (eq.operatingHours || 0) / 40000);
      const current = {};
      for (const key of Object.keys(profile)) {
        current[key] = generateSensorReading(profile, key, ageFactor, simTime);
      }
      const health = computeHealth(current, profile);
      const prevHealth = sensorData[eq.id]?.health ?? health;
      const trend = (health - prevHealth);
      newData[eq.id] = { current, health, trend, failureDays: predictFailureDays(health, trend) };
    }
    setSensorData(newData);

    // append to history (keep last 60 points)
    setSensorHistory(prev => {
      const h = { ...prev };
      for (const eq of project.equipment) {
        const sd = newData[eq.id];
        if (!sd) continue;
        const arr = h[eq.id] || [];
        const entry = { t: simTime, health: sd.health };
        for (const k of Object.keys(sd.current)) { entry[k] = sd.current[k]; }
        h[eq.id] = [...arr.slice(-59), entry];
      }
      return h;
    });

    // generate alerts
    const newAlerts = generateMaintenanceAlerts(project.equipment, newData);
    setAlerts(newAlerts);
  }, [simTime, project.equipment]);

  /* ---- derived ---- */
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

  /* ---- facility-wide KPIs ---- */
  const facilityKPIs = useMemo(() => {
    if (project.equipment.length === 0) return null;
    const healths = project.equipment.map(e => sensorData[e.id]?.health ?? 100);
    const avgHealth = Math.round(healths.reduce((a, b) => a + b, 0) / healths.length);
    const critical = healths.filter(h => h < 40).length;
    const warning = healths.filter(h => h >= 40 && h < 70).length;
    const ok = healths.filter(h => h >= 70).length;
    const totalPower = project.equipment.reduce((s, e) => s + (parseFloat(e.power) || 0), 0);
    const totalHours = project.equipment.reduce((s, e) => s + (e.operatingHours || 0), 0);
    const totalMainCost = project.equipment.reduce((s, e) => s + (e.maintenanceLog || []).reduce((ss, m) => ss + (m.cost || 0), 0), 0);
    const uptime = Math.round((ok / healths.length) * 100);
    return { avgHealth, critical, warning, ok, totalPower, totalHours, totalMainCost, uptime, count: healths.length };
  }, [project.equipment, sensorData]);

  /* ---- file imports ---- */
  async function onDxfFile(file) {
    if (!file) return;
    if (/\.dwg$/i.test(file.name)) { toast("DWG is a closed binary format — export to DXF first (ODA File Converter is free).", "err"); return; }
    try {
      const fp = parseDXF(await file.text());
      if (!fp.count) { toast("No supported entities in this DXF.", "err"); return; }
      setProject(p => ({ ...p, floorplan: { ...fp, source: file.name } })); fitView(fp.bounds);
      toast(`Imported ${fp.count} entities (${fp.assumedUnits})`, "ok");
    } catch (e) { toast("DXF parse error: " + e.message, "err"); }
  }
  async function onEquipmentFile(file) {
    if (!file) return;
    try {
      let rows;
      if (/\.(xlsx|xls|xlsm)$/i.test(file.name)) { const wb = XLSX.read(await file.arrayBuffer()); rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }); }
      else { rows = Papa.parse(await file.text(), { header: true, skipEmptyLines: true }).data; }
      const eqs = rowsToEquipment(rows);
      if (!eqs.length) { toast("No rows recognized.", "err"); return; }
      setProject(p => ({ ...p, equipment: [...p.equipment, ...eqs] }));
      setLeftTab("equipment"); toast(`Imported ${eqs.length} equipment`, "ok");
    } catch (e) { toast("Import failed: " + e.message, "err"); }
  }
  function onProjectJson(file) {
    if (!file) return;
    file.text().then(t => {
      try { const p = JSON.parse(t); if (!p.equipment) throw new Error("not a project file"); setProject({ ...emptyProject(), ...p }); toast("Project loaded", "ok"); }
      catch (e) { toast("Invalid project: " + e.message, "err"); }
    });
  }

  /* ---- canvas ---- */
  const clientToWorld = useCallback((cx, cy) => {
    const el = svgRef.current; if (!el) return { x: 0, y: 0 };
    const r = el.getBoundingClientRect();
    const h = cam.w * (r.height / r.width);
    return { x: cam.x + ((cx - r.left) / r.width) * cam.w, y: cam.y + ((cy - r.top) / r.height) * h };
  }, [cam]);
  function onWheel(e) {
    const pt = clientToWorld(e.clientX, e.clientY);
    const f = e.deltaY > 0 ? 1.12 : 1 / 1.12;
    setCam(c => { const w = Math.min(2000, Math.max(1, c.w * f)); return { x: pt.x - (pt.x - c.x) * (w / c.w), y: pt.y - (pt.y - c.y) * (w / c.w), w }; });
  }
  function onCanvasDown(e, hit) {
    const pt = clientToWorld(e.clientX, e.clientY);
    if (mode === "add-equip") {
      const eq = { id: uid(), tag: `EQ-${String(project.equipment.length + 1).padStart(3, "0")}`, name: "New equipment", type: "Process", manufacturer: "", model: "", power: "", area: "", x: pt.x, y: pt.y, placed: true, w: 2, d: 1.5, h: 1.8, bom: [], installDate: "", lastMaintenance: "", maintenanceLog: [], operatingHours: 0 };
      setProject(p => ({ ...p, equipment: [...p.equipment, eq] })); setSelected({ kind: "equipment", id: eq.id }); setMode("select"); return;
    }
    if (mode === "add-utility") {
      const u = { id: uid(), tag: `${UTILITY_TYPES[pendingUtilType].short}-${String(project.utilities.length + 1).padStart(2, "0")}`, name: UTILITY_TYPES[pendingUtilType].label, type: pendingUtilType, x: pt.x, y: pt.y, capacity: "" };
      setProject(p => ({ ...p, utilities: [...p.utilities, u] })); setSelected({ kind: "utility", id: u.id }); setMode("select"); return;
    }
    if (placing) { setProject(p => ({ ...p, equipment: p.equipment.map(q => q.id === placing ? { ...q, x: pt.x, y: pt.y, placed: true } : q) })); setSelected({ kind: "equipment", id: placing }); setPlacing(null); return; }
    if (mode === "connect") {
      if (hit) {
        if (!connectFrom) { setConnectFrom(hit.id); return; }
        if (connectFrom !== hit.id) {
          const from = nodes[connectFrom], to = nodes[hit.id];
          const utype = from?.kind === "utility" ? from.type : to?.kind === "utility" ? to.type : "process";
          setProject(p => ({ ...p, connections: [...p.connections, { id: uid(), fromId: connectFrom, toId: hit.id, utilityType: utype, medium: "", size: "", notes: "" }] }));
          setConnectFrom(null); setMode("select"); toast(`Connected ${from?.tag} → ${to?.tag}`, "ok");
        }
      }
      return;
    }
    if (hit) { setSelected({ kind: hit.kind, id: hit.id }); dragRef.current = { id: hit.id, kind: hit.kind, start: pt, orig: { x: nodes[hit.id].x, y: nodes[hit.id].y }, moved: false }; }
    else { setSelected(null); dragRef.current = { pan: true, start: { x: e.clientX, y: e.clientY }, cam0: cam }; }
  }
  function onCanvasMove(e) {
    const pt = clientToWorld(e.clientX, e.clientY); setCursor(pt);
    const d = dragRef.current; if (!d) return;
    if (d.pan) { const el = svgRef.current; if (!el) return; const r = el.getBoundingClientRect(); setCam({ ...d.cam0, x: d.cam0.x - ((e.clientX - d.start.x) / r.width) * d.cam0.w, y: d.cam0.y - ((e.clientY - d.start.y) / r.width) * d.cam0.w }); return; }
    const dx = pt.x - d.start.x, dy = pt.y - d.start.y;
    if (Math.abs(dx) + Math.abs(dy) > 0.05) d.moved = true; if (!d.moved) return;
    const nx = d.orig.x + dx, ny = d.orig.y + dy;
    if (d.kind === "equipment") setProject(p => ({ ...p, equipment: p.equipment.map(q => q.id === d.id ? { ...q, x: nx, y: ny } : q) }));
    else if (d.kind === "utility") setProject(p => ({ ...p, utilities: p.utilities.map(q => q.id === d.id ? { ...q, x: nx, y: ny } : q) }));
  }
  function onCanvasUp() { dragRef.current = null; }

  /* ---- mutations ---- */
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

  /* ---- GitHub ---- */
  async function ghRequest(url, opts = {}) {
    return fetch(url, { ...opts, headers: { Accept: "application/vnd.github+json", Authorization: `Bearer ${project.github.token}`, "X-GitHub-Api-Version": "2022-11-28", ...(opts.headers || {}) } });
  }
  async function ghPush() {
    const { owner, repo, branch, path } = project.github; setGhBusy(true);
    try {
      const clean = { ...project, github: { ...project.github, token: "" } };
      const content = btoa(unescape(encodeURIComponent(JSON.stringify(clean, null, 2))));
      const base = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
      let sha; const head = await ghRequest(`${base}?ref=${branch}`); if (head.ok) sha = (await head.json()).sha;
      const res = await ghRequest(base, { method: "PUT", body: JSON.stringify({ message: `FacilityTwin snapshot — ${project.name} rev ${project.revision}`, content, branch, ...(sha ? { sha } : {}) }) });
      if (res.ok) toast(`Committed to ${owner}/${repo}@${branch}`, "ok"); else toast(`Commit failed (${res.status})`, "err");
    } catch { toast("Network error — deploy the app for GitHub access.", "err"); }
    setGhBusy(false);
  }

  /* ---- 3D view ---- */
  useEffect(() => {
    if (topView !== "plan" || planMode !== "3d" || !threeRef.current) return;
    const el = threeRef.current, W = el.clientWidth, H = el.clientHeight;
    const scene = new THREE.Scene(); scene.background = new THREE.Color(0x101418); scene.fog = new THREE.Fog(0x101418, 60, 220);
    const camera = new THREE.PerspectiveCamera(50, W / H, 0.1, 500);
    const renderer = new THREE.WebGLRenderer({ antialias: true }); renderer.setSize(W, H); renderer.setPixelRatio(Math.min(devicePixelRatio, 2)); el.appendChild(renderer.domElement);
    scene.add(new THREE.AmbientLight(0xffffff, 0.55));
    const sun = new THREE.DirectionalLight(0xffe8c8, 0.9); sun.position.set(30, 50, 20); scene.add(sun);
    const b = project.floorplan?.bounds || { minX: 0, minY: 0, maxX: 40, maxY: 25 };
    const cx = (b.minX + b.maxX) / 2, cz = (b.minY + b.maxY) / 2, span = Math.max(b.maxX - b.minX, b.maxY - b.minY);
    const floor = new THREE.Mesh(new THREE.PlaneGeometry(b.maxX - b.minX + 8, b.maxY - b.minY + 8), new THREE.MeshStandardMaterial({ color: 0x171c21, roughness: 0.95 }));
    floor.rotation.x = -Math.PI / 2; floor.position.set(cx, -0.02, cz); scene.add(floor);
    const grid = new THREE.GridHelper(Math.ceil(span + 8), Math.ceil(span + 8), 0x2a333c, 0x1d242b); grid.position.set(cx, 0, cz); scene.add(grid);
    if (project.floorplan) {
      const pts = [];
      for (const e of project.floorplan.entities) {
        if (e.kind === "line" || e.kind === "poly") {
          const arr = e.kind === "poly" && e.closed ? [...e.pts, e.pts[0]] : e.pts;
          for (let i = 0; i < arr.length - 1; i++) pts.push(new THREE.Vector3(arr[i][0], 0.02, arr[i][1]), new THREE.Vector3(arr[i + 1][0], 0.02, arr[i + 1][1]));
        }
      }
      if (pts.length) scene.add(new THREE.LineSegments(new THREE.BufferGeometry().setFromPoints(pts), new THREE.LineBasicMaterial({ color: 0x55616b })));
    }
    for (const e of project.equipment.filter(q => q.placed)) {
      const isSel = selected?.id === e.id;
      const health = sensorData[e.id]?.health ?? 100;
      const col = health < 40 ? 0xe06552 : health < 70 ? 0xefcb4f : isSel ? 0xf08c3a : 0x3a4650;
      const mesh = new THREE.Mesh(new THREE.BoxGeometry(e.w, e.h, e.d), new THREE.MeshStandardMaterial({ color: col, roughness: 0.6, metalness: 0.25 }));
      mesh.position.set(e.x, e.h / 2, e.y); scene.add(mesh);
      scene.add(Object.assign(new THREE.LineSegments(new THREE.EdgesGeometry(mesh.geometry), new THREE.LineBasicMaterial({ color: isSel ? 0xffc48a : 0x55616b })), { position: mesh.position.clone() }));
    }
    for (const u of project.utilities) {
      const cyl = new THREE.Mesh(new THREE.CylinderGeometry(0.7, 0.7, 2.2, 20), new THREE.MeshStandardMaterial({ color: new THREE.Color(UTILITY_TYPES[u.type]?.color || "#888"), roughness: 0.5 }));
      cyl.position.set(u.x, 1.1, u.y); scene.add(cyl);
    }
    for (const c of project.connections) {
      const a = nodes[c.fromId], t2 = nodes[c.toId];
      if (!a || !t2 || a.x === undefined || t2.x === undefined) continue;
      const col = new THREE.Color(UTILITY_TYPES[c.utilityType]?.color || "#888");
      const hRun = 3.2;
      scene.add(new THREE.Line(new THREE.BufferGeometry().setFromPoints([
        new THREE.Vector3(a.x, a.h || 2.2, a.y), new THREE.Vector3(a.x, hRun, a.y),
        new THREE.Vector3(a.x, hRun, t2.y), new THREE.Vector3(t2.x, hRun, t2.y),
        new THREE.Vector3(t2.x, t2.h || 2.2, t2.y)
      ]), new THREE.LineBasicMaterial({ color: col })));
    }
    let az = -Math.PI / 4, elv = 0.9, radius = span * 1.1 + 10;
    const target = new THREE.Vector3(cx, 1, cz);
    const applyCam = () => { camera.position.set(target.x + radius * Math.cos(elv) * Math.cos(az), target.y + radius * Math.sin(elv), target.z + radius * Math.cos(elv) * Math.sin(az)); camera.lookAt(target); };
    applyCam();
    let down = null;
    const onD = e => { down = { x: e.clientX, y: e.clientY, az, elv }; };
    const onM = e => { if (!down) return; az = down.az + (e.clientX - down.x) * 0.006; elv = Math.min(1.45, Math.max(0.1, down.elv + (e.clientY - down.y) * 0.005)); applyCam(); };
    const onU = () => { down = null; };
    const onZ = e => { e.preventDefault(); radius = Math.min(400, Math.max(6, radius * (e.deltaY > 0 ? 1.1 : 0.9))); applyCam(); };
    renderer.domElement.addEventListener("mousedown", onD); window.addEventListener("mousemove", onM); window.addEventListener("mouseup", onU);
    renderer.domElement.addEventListener("wheel", onZ, { passive: false });
    let alive = true; const loop = () => { if (!alive) return; renderer.render(scene, camera); requestAnimationFrame(loop); }; loop();
    const onR = () => { camera.aspect = el.clientWidth / el.clientHeight; camera.updateProjectionMatrix(); renderer.setSize(el.clientWidth, el.clientHeight); };
    window.addEventListener("resize", onR);
    return () => { alive = false; window.removeEventListener("mousemove", onM); window.removeEventListener("mouseup", onU); window.removeEventListener("resize", onR); renderer.dispose(); el.innerHTML = ""; };
  }, [topView, planMode, project, selected, nodes, sensorData]);

  /* ============================ RENDER ============================== */
  const camH = cam.w * aspect;
  const gridStep = cam.w > 120 ? 10 : cam.w > 30 ? 5 : 1;
  const strokeW = cam.w / 900;
  const connCountFor = id => project.connections.filter(c => c.fromId === id || c.toId === id).length;
  const unplaced = project.equipment.filter(e => !e.placed);

  const modeBanner =
    mode === "add-equip" ? "CLICK TO PLACE EQUIPMENT" :
    mode === "add-utility" ? `CLICK TO PLACE ${UTILITY_TYPES[pendingUtilType].short}` :
    mode === "connect" ? (connectFrom ? `FROM: ${nodes[connectFrom]?.tag} — CLICK TARGET` : "CLICK SOURCE NODE") :
    placing ? "CLICK TO PLACE" : null;

  return (
    <div className="ft">
      <style>{CSS}</style>

      {/* ============ TOOLBAR ============ */}
      <div className="tb">
        <div className="tb-block">
          <div className="tb-ey">FacilityTwin</div>
          <div className="tb-tl"><input value={project.name} onChange={e => setProject(p => ({ ...p, name: e.target.value }))} /></div>
        </div>
        <div className="tb-block" style={{ minWidth: 50 }}>
          <div className="tb-ey">Rev</div>
          <div className="tb-tl" style={{ fontFamily: "var(--mono)" }}><input style={{ width: 32 }} value={project.revision} onChange={e => setProject(p => ({ ...p, revision: e.target.value }))} /></div>
        </div>

        <div className="tb-g">
          {["plan", "dashboard", "analytics"].map(v => (
            <button key={v} className={`btn ${topView === v ? "on" : ""}`} onClick={() => setTopView(v)}>{v.toUpperCase()}</button>
          ))}
        </div>

        {topView === "plan" && (
          <div className="tb-g">
            <button className={`btn ${mode === "select" ? "on" : ""}`} onClick={() => { setMode("select"); setConnectFrom(null); }}>SELECT</button>
            <button className={`btn ${mode === "add-equip" ? "on" : ""}`} onClick={() => setMode("add-equip")}>+ EQUIP</button>
            <button className={`btn ${mode === "add-utility" ? "on" : ""}`} onClick={() => setMode("add-utility")}>+ UTIL</button>
            {mode === "add-utility" && (
              <select value={pendingUtilType} onChange={e => setPendingUtilType(e.target.value)} style={{ background: "var(--bg)", color: "var(--text)", border: "1px solid var(--line2)", borderRadius: 3, padding: "4px 5px", fontFamily: "var(--mono)", fontSize: 10 }}>
                {Object.entries(UTILITY_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
              </select>
            )}
            <button className={`btn ${mode === "connect" ? "on" : ""}`} onClick={() => { setMode("connect"); setConnectFrom(null); }}>⇄ CONNECT</button>
          </div>
        )}

        <div className="tb-sp" />

        {topView === "plan" && (
          <div className="tb-g">
            <button className={`btn ${planMode === "2d" ? "on" : ""}`} onClick={() => setPlanMode("2d")}>2D</button>
            <button className={`btn ${planMode === "3d" ? "on" : ""}`} onClick={() => setPlanMode("3d")}>3D</button>
            <button className="btn" onClick={() => fitView()}>FIT</button>
          </div>
        )}

        <div className="tb-g">
          <button className={`btn sm ${simRunning ? "" : "on"}`} onClick={() => setSimRunning(!simRunning)} style={{ fontSize: 9 }}>
            {simRunning ? "⏸ PAUSE SIM" : "▶ RUN SIM"}
          </button>
          {alerts.filter(a => a.severity === "critical").length > 0 && (
            <span style={{ background: "var(--danger)", color: "#fff", fontFamily: "var(--mono)", fontSize: 9, padding: "2px 6px", borderRadius: 2 }}>
              {alerts.filter(a => a.severity === "critical").length} CRITICAL
            </span>
          )}
        </div>

        <div className="tb-g">
          <button className="btn" onClick={() => setGhOpen(true)}>
            {project.github.owner ? `⎇ ${project.github.owner}/${project.github.repo}` : "⎇ GITHUB"}
          </button>
        </div>
      </div>

      <div className="main">
        {/* ============ LEFT SIDEBAR ============ */}
        <div className="side">
          <div className="stabs">
            {["import", "equipment", "utilities"].map(t => (
              <button key={t} className={`stab ${leftTab === t ? "on" : ""}`} onClick={() => setLeftTab(t)}>{t}</button>
            ))}
          </div>
          <div className="sbody">
            {leftTab === "import" && (
              <>
                <div className="sec">Floor plan</div>
                <p className="hint">Import a 2D floor plan as <b>DXF</b>.</p>
                <button className="btn" style={{ width: "100%" }} onClick={() => fileDxfRef.current.click()}>IMPORT DXF…</button>
                <input ref={fileDxfRef} type="file" accept=".dxf,.dwg" hidden onChange={e => { onDxfFile(e.target.files[0]); e.target.value = ""; }} />
                {project.floorplan && <p className="hint" style={{ fontFamily: "var(--mono)", fontSize: 10 }}>▣ {project.floorplan.source} — {project.floorplan.entities.length} entities</p>}
                <div className="sec">Equipment list</div>
                <p className="hint">Import CSV or Excel with equipment data.</p>
                <button className="btn" style={{ width: "100%" }} onClick={() => fileListRef.current.click()}>IMPORT CSV / XLSX…</button>
                <input ref={fileListRef} type="file" accept=".csv,.xlsx,.xls,.xlsm,.tsv" hidden onChange={e => { onEquipmentFile(e.target.files[0]); e.target.value = ""; }} />
                <div className="sec">Project</div>
                <button className="btn gh" style={{ width: "100%", marginBottom: 5 }} onClick={() => { setProject(demoProject()); setSelected(null); setTimeout(() => fitView(), 0); toast("Demo loaded", "ok"); }}>LOAD DEMO FACILITY</button>
                <button className="btn gh" style={{ width: "100%", marginBottom: 5 }} onClick={() => downloadFile(`${project.name.replace(/\W+/g, "-")}.json`, JSON.stringify({ ...project, github: { ...project.github, token: "" } }, null, 2))}>EXPORT JSON</button>
                <button className="btn gh" style={{ width: "100%", marginBottom: 5 }} onClick={() => fileJsonRef.current.click()}>OPEN JSON…</button>
                <input ref={fileJsonRef} type="file" accept=".json" hidden onChange={e => { onProjectJson(e.target.files[0]); e.target.value = ""; }} />
                <button className="btn gh dng" style={{ width: "100%" }} onClick={() => { setProject(emptyProject()); setSelected(null); }}>NEW EMPTY</button>
              </>
            )}
            {leftTab === "equipment" && (
              <>
                {unplaced.length > 0 && (<><div className="sec">Staged</div>{unplaced.map(e => (
                  <div key={e.id} className={`ri ${placing === e.id ? "on" : ""}`} onClick={() => setPlacing(e.id)}>
                    <span className="rdot" style={{ background: "var(--accent)" }} /><span className="rtag">{e.tag}</span><span className="rname">{e.name}</span>
                  </div>
                ))}</>)}
                <div className="sec">Equipment ({project.equipment.filter(e => e.placed).length})</div>
                {project.equipment.filter(e => e.placed).map(e => {
                  const h = sensorData[e.id]?.health;
                  const hCol = h !== undefined ? (h < 40 ? "var(--danger)" : h < 70 ? "var(--warn)" : "var(--ok)") : "var(--line2)";
                  return (
                    <div key={e.id} className={`ri ${selected?.id === e.id ? "on" : ""}`} onClick={() => { setSelected({ kind: "equipment", id: e.id }); if (topView !== "plan") setTopView("plan"); }}>
                      <span className="rdot" style={{ background: hCol }} />
                      <span className="rtag">{e.tag}</span>
                      <span className="rname">{e.name}</span>
                      {h !== undefined && <span className="pill" style={{ color: hCol, borderColor: hCol }}>{h}%</span>}
                    </div>
                  );
                })}
              </>
            )}
            {leftTab === "utilities" && (
              <>
                <div className="sec">Utilities ({project.utilities.length})</div>
                {project.utilities.map(u => (
                  <div key={u.id} className={`ri ${selected?.id === u.id ? "on" : ""}`} onClick={() => setSelected({ kind: "utility", id: u.id })}>
                    <span className="rdot" style={{ background: UTILITY_TYPES[u.type]?.color, borderRadius: 6 }} /><span className="rtag">{u.tag}</span><span className="rname">{u.name}</span>
                  </div>
                ))}
                <div className="sec">Connections ({project.connections.length})</div>
                {project.connections.map(c => (
                  <div key={c.id} className={`cc ${selected?.id === c.id ? "on" : ""}`} onClick={() => setSelected({ kind: "connection", id: c.id })}>
                    <span className="rdot" style={{ background: UTILITY_TYPES[c.utilityType]?.color }} />
                    <span className="rtag" style={{ fontSize: 10 }}>{nodes[c.fromId]?.tag} → {nodes[c.toId]?.tag}</span>
                    <span className="rname">{c.medium || UTILITY_TYPES[c.utilityType]?.label}</span>
                  </div>
                ))}
              </>
            )}
          </div>
        </div>

        {/* ============ CENTER CONTENT ============ */}
        {topView === "plan" && (
          <div className="cwrap">
            {modeBanner && <div className="mbanner">{modeBanner}</div>}
            {planMode === "2d" && (
              <svg ref={svgRef} className="csvg" viewBox={`${cam.x} ${cam.y} ${cam.w} ${camH}`} preserveAspectRatio="none"
                onWheel={onWheel} onMouseDown={e => onCanvasDown(e, null)} onMouseMove={onCanvasMove} onMouseUp={onCanvasUp} onMouseLeave={onCanvasUp}
                style={{ cursor: mode === "select" && !placing ? "default" : "crosshair" }}>
                {/* grid */}
                <g stroke="var(--line)" strokeWidth={strokeW * 0.6} opacity="0.45">
                  {Array.from({ length: Math.ceil(cam.w / gridStep) + 2 }, (_, i) => {
                    const x = Math.floor(cam.x / gridStep) * gridStep + i * gridStep;
                    return <line key={"v" + i} x1={x} y1={cam.y - 5} x2={x} y2={cam.y + camH + 5} />;
                  })}
                  {Array.from({ length: Math.ceil(camH / gridStep) + 2 }, (_, i) => {
                    const y = Math.floor(cam.y / gridStep) * gridStep + i * gridStep;
                    return <line key={"h" + i} x1={cam.x - 5} y1={y} x2={cam.x + cam.w + 5} y2={y} />;
                  })}
                </g>
                {/* floor plan */}
                {project.floorplan && (
                  <g stroke="#6d7c88" fill="none" strokeWidth={strokeW * 1.4} strokeLinejoin="round">
                    {project.floorplan.entities.map((e, i) => {
                      if (e.kind === "line") return <line key={i} x1={e.pts[0][0]} y1={e.pts[0][1]} x2={e.pts[1][0]} y2={e.pts[1][1]} />;
                      if (e.kind === "poly") return <polyline key={i} points={(e.closed ? [...e.pts, e.pts[0]] : e.pts).map(p => p.join(",")).join(" ")} />;
                      if (e.kind === "circle") return <circle key={i} cx={e.cx} cy={e.cy} r={e.r} />;
                      if (e.kind === "text") return <text key={i} x={e.x} y={e.y} fontSize={e.h} fill="#55616b" stroke="none" fontFamily="var(--mono)">{e.text}</text>;
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
                    return (<g key={c.id} onMouseDown={e => { e.stopPropagation(); setSelected({ kind: "connection", id: c.id }); }} style={{ cursor: "pointer" }}>
                      <path d={orthoPath(nodeAnchor(a), nodeAnchor(b2))} stroke="transparent" strokeWidth={strokeW * 12} />
                      <path d={orthoPath(nodeAnchor(a), nodeAnchor(b2))} stroke={col} strokeWidth={strokeW * (on ? 4 : 2)} strokeDasharray={c.utilityType === "power" ? `${strokeW * 8} ${strokeW * 5}` : "none"} opacity={on ? 1 : 0.7} />
                    </g>);
                  })}
                </g>
                {/* utilities */}
                {project.utilities.map(u => {
                  const col = UTILITY_TYPES[u.type]?.color || "#888";
                  const on = selected?.id === u.id || connectFrom === u.id;
                  return (<g key={u.id} onMouseDown={e => { e.stopPropagation(); onCanvasDown(e, { kind: "utility", id: u.id }); }} style={{ cursor: "pointer" }}>
                    <circle cx={u.x} cy={u.y} r={1} fill="var(--panel2)" stroke={col} strokeWidth={strokeW * (on ? 4 : 2.2)} />
                    <text x={u.x} y={u.y + 0.18} fontSize={0.5} textAnchor="middle" fill={col} fontFamily="var(--mono)" fontWeight="700">{UTILITY_TYPES[u.type]?.short}</text>
                    <text x={u.x} y={u.y - 1.3} fontSize={0.5} textAnchor="middle" fill="var(--text)" fontFamily="var(--mono)">{u.tag}</text>
                  </g>);
                })}
                {/* equipment with health coloring */}
                {project.equipment.filter(e => e.placed).map(e => {
                  const on = selected?.id === e.id || connectFrom === e.id;
                  const h = sensorData[e.id]?.health;
                  const fill = h !== undefined ? (h < 40 ? "rgba(224,101,82,.25)" : h < 70 ? "rgba(239,203,79,.18)" : "rgba(98,210,111,.1)") : "rgba(58,70,80,.55)";
                  const stroke = on ? "var(--accent)" : h !== undefined ? (h < 40 ? "var(--danger)" : h < 70 ? "var(--warn)" : "#8b9aa6") : "#8b9aa6";
                  return (<g key={e.id} onMouseDown={ev => { ev.stopPropagation(); onCanvasDown(ev, { kind: "equipment", id: e.id }); }} style={{ cursor: mode === "connect" ? "pointer" : "move" }}>
                    <rect x={e.x - e.w / 2} y={e.y - e.d / 2} width={e.w} height={e.d} rx={0.08} fill={fill} stroke={stroke} strokeWidth={strokeW * (on ? 3 : 1.6)} />
                    <line x1={e.x - e.w / 2} y1={e.y - e.d / 2} x2={e.x + e.w / 2} y2={e.y + e.d / 2} stroke={stroke} strokeWidth={strokeW * 0.8} opacity={0.5} />
                    <text x={e.x} y={e.y - e.d / 2 - 0.3} fontSize={0.55} textAnchor="middle" fill={on ? "var(--accent)" : "var(--text)"} fontFamily="var(--mono)" fontWeight="700">{e.tag}</text>
                    {h !== undefined && <text x={e.x} y={e.y + e.d / 2 + 0.7} fontSize={0.4} textAnchor="middle" fill={h < 40 ? "var(--danger)" : h < 70 ? "var(--warn)" : "var(--ok)"} fontFamily="var(--mono)">{h}%</text>}
                  </g>);
                })}
              </svg>
            )}
            {planMode === "3d" && <div ref={threeRef} className="c3d" />}
            {planMode === "2d" && !project.floorplan && project.equipment.length === 0 && (
              <div className="empty"><div className="big">NO FLOOR PLAN LOADED</div><div>Import a DXF, or explore the demo.</div>
                <button className="btn" onClick={() => { setProject(demoProject()); setTimeout(() => fitView(), 0); }}>LOAD DEMO FACILITY</button></div>
            )}
            <div className="sbar">
              <div>X <b>{cursor.x.toFixed(2)}</b> Y <b>{cursor.y.toFixed(2)}</b></div>
              <div>EQUIP <b>{project.equipment.length}</b></div><div>UTIL <b>{project.utilities.length}</b></div><div>CONN <b>{project.connections.length}</b></div>
              <div style={{ marginLeft: "auto", borderRight: "none" }}>SIM {simRunning ? "RUNNING" : "PAUSED"} · t={simTime}</div>
            </div>
            <div className="toasts">{toasts.map(t => <div key={t.id} className={`toast ${t.kind}`}>{t.msg}</div>)}</div>
          </div>
        )}

        {/* ============ DASHBOARD VIEW ============ */}
        {topView === "dashboard" && (
          <div className="dash">
            {!facilityKPIs ? (
              <div style={{ textAlign: "center", padding: 60, color: "var(--muted)" }}>
                <div style={{ fontFamily: "var(--mono)", fontSize: 12, letterSpacing: ".12em", color: "var(--faint)", marginBottom: 8 }}>NO EQUIPMENT DATA</div>
                <div>Load a project or the demo facility to see live analytics.</div>
                <button className="btn" style={{ marginTop: 12 }} onClick={() => { setProject(demoProject()); setTimeout(() => fitView(), 0); }}>LOAD DEMO</button>
              </div>
            ) : (
              <>
                <div className="kpi-row">
                  <div className="kpi"><div className="kpi-label">Avg Health</div><div className="kpi-val" style={{ color: facilityKPIs.avgHealth > 70 ? "var(--ok)" : facilityKPIs.avgHealth > 40 ? "var(--warn)" : "var(--danger)" }}>{facilityKPIs.avgHealth}%</div><div className="kpi-sub">Across {facilityKPIs.count} assets</div></div>
                  <div className="kpi"><div className="kpi-label">Uptime</div><div className="kpi-val" style={{ color: "var(--ok)" }}>{facilityKPIs.uptime}%</div><div className="kpi-sub">{facilityKPIs.ok} assets healthy</div></div>
                  <div className="kpi"><div className="kpi-label">Total Power</div><div className="kpi-val">{facilityKPIs.totalPower}<span style={{ fontSize: 14, color: "var(--muted)" }}> kW</span></div><div className="kpi-sub">Connected load</div></div>
                  <div className="kpi"><div className="kpi-label">Operating Hours</div><div className="kpi-val">{(facilityKPIs.totalHours / 1000).toFixed(1)}<span style={{ fontSize: 14, color: "var(--muted)" }}>k</span></div><div className="kpi-sub">Fleet total</div></div>
                  <div className="kpi"><div className="kpi-label">Maintenance Cost</div><div className="kpi-val">${(facilityKPIs.totalMainCost / 1000).toFixed(1)}<span style={{ fontSize: 14, color: "var(--muted)" }}>k</span></div><div className="kpi-sub">Recorded to date</div></div>
                  <div className="kpi"><div className="kpi-label">Alerts</div><div className="kpi-val" style={{ color: facilityKPIs.critical > 0 ? "var(--danger)" : "var(--text)" }}>{facilityKPIs.critical + facilityKPIs.warning}</div><div className="kpi-sub">{facilityKPIs.critical} critical · {facilityKPIs.warning} warning</div></div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                  <div className="card">
                    <div className="card-h">Fleet Health Distribution</div>
                    <div style={{ height: 200 }}>
                      <ResponsiveContainer>
                        <PieChart>
                          <Pie data={[{ name: "Healthy", value: facilityKPIs.ok }, { name: "Warning", value: facilityKPIs.warning }, { name: "Critical", value: facilityKPIs.critical }]}
                            cx="50%" cy="50%" innerRadius={50} outerRadius={80} dataKey="value" strokeWidth={0}>
                            <Cell fill="var(--ok)" /><Cell fill="var(--warn)" /><Cell fill="var(--danger)" />
                          </Pie>
                          <Tooltip contentStyle={{ background: "var(--panel)", border: "1px solid var(--line)", fontSize: 11, fontFamily: "var(--mono)" }} />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="card">
                    <div className="card-h">Power Load by Equipment</div>
                    <div style={{ height: 200 }}>
                      <ResponsiveContainer>
                        <BarChart data={project.equipment.filter(e => e.power).map(e => ({ name: e.tag, kW: parseFloat(e.power) || 0 }))} margin={{ top: 5, right: 10, left: -15, bottom: 5 }}>
                          <CartesianGrid stroke="var(--line)" strokeDasharray="3 3" />
                          <XAxis dataKey="name" tick={{ fill: "var(--muted)", fontSize: 10, fontFamily: "var(--mono)" }} />
                          <YAxis tick={{ fill: "var(--muted)", fontSize: 10 }} />
                          <Tooltip contentStyle={{ background: "var(--panel)", border: "1px solid var(--line)", fontSize: 11 }} />
                          <Bar dataKey="kW" fill="var(--accent)" radius={[2, 2, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>

                <div className="card">
                  <div className="card-h">Asset Health Overview<span style={{ fontSize: 9, color: "var(--faint)" }}>LIVE · CLICK TO INSPECT</span></div>
                  <div className="ht-grid">
                    {project.equipment.map(e => {
                      const h = sensorData[e.id]?.health ?? 100;
                      const col = h < 40 ? "var(--danger)" : h < 70 ? "var(--warn)" : "var(--ok)";
                      const fd = sensorData[e.id]?.failureDays;
                      return (
                        <div key={e.id} className="ht-item" onClick={() => { setSelected({ kind: "equipment", id: e.id }); setTopView("analytics"); }}>
                          <HealthGauge value={h} size={52} />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontFamily: "var(--mono)", fontWeight: 700, fontSize: 12 }}>{e.tag}</div>
                            <div style={{ fontSize: 10, color: "var(--muted)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{e.name}</div>
                            <div style={{ display: "flex", gap: 6, marginTop: 3 }}>
                              <div className="ht-bar"><div className="ht-fill" style={{ width: h + "%", background: col }} /></div>
                              {fd && <span style={{ fontFamily: "var(--mono)", fontSize: 9, color: "var(--danger)" }}>{fd}d</span>}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                <div className="card">
                  <div className="card-h">Live Alerts Feed<span className="pill">{alerts.length}</span></div>
                  {alerts.length === 0 && <p className="hint">No active alerts — all systems nominal.</p>}
                  {alerts.slice(0, 20).map(a => (
                    <div key={a.id} className="alert-row" style={{ cursor: "pointer" }} onClick={() => { setSelected({ kind: "equipment", id: a.equipId }); setTopView("analytics"); }}>
                      <div className="alert-dot" style={{ background: SEVERITY[a.severity]?.color }} />
                      <div className="alert-msg">{a.message}</div>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        )}

        {/* ============ ANALYTICS VIEW ============ */}
        {topView === "analytics" && (
          <div className="dash">
            {!selEquip ? (
              <div style={{ textAlign: "center", padding: 60, color: "var(--muted)" }}>
                <div style={{ fontFamily: "var(--mono)", fontSize: 12, letterSpacing: ".12em", color: "var(--faint)", marginBottom: 8 }}>SELECT AN ASSET</div>
                <div>Click any equipment in the left sidebar to see its sensor data, trend analysis, and predictive maintenance.</div>
              </div>
            ) : (
              <>
                {/* header */}
                <div style={{ display: "flex", alignItems: "center", gap: 16, padding: "4px 0 8px" }}>
                  <HealthGauge value={sensorData[selEquip.id]?.health ?? 100} size={80} />
                  <div>
                    <div style={{ fontFamily: "var(--mono)", fontSize: 18, fontWeight: 700 }}>{selEquip.tag}</div>
                    <div style={{ color: "var(--muted)", fontSize: 13 }}>{selEquip.name}</div>
                    <div style={{ fontFamily: "var(--mono)", fontSize: 10, color: "var(--faint)", marginTop: 4 }}>
                      {selEquip.manufacturer} {selEquip.model} · {selEquip.operatingHours?.toLocaleString() || "—"} hrs · {selEquip.area || "—"}
                    </div>
                    {sensorData[selEquip.id]?.failureDays && (
                      <div style={{ fontFamily: "var(--mono)", fontSize: 11, color: "var(--danger)", marginTop: 4 }}>
                        ⚠ PREDICTED MAINTENANCE IN {sensorData[selEquip.id].failureDays} DAYS
                      </div>
                    )}
                  </div>
                </div>

                {/* live sensors */}
                <div className="card">
                  <div className="card-h">Live Sensor Readings<span style={{ fontSize: 9, color: simRunning ? "var(--ok)" : "var(--faint)" }}>{simRunning ? "● LIVE" : "○ PAUSED"}</span></div>
                  {(() => {
                    const sd = sensorData[selEquip.id];
                    const profile = SENSOR_PROFILES[selEquip.type] || SENSOR_PROFILES.Other;
                    if (!sd) return <p className="hint">Waiting for sensor data…</p>;
                    return Object.keys(profile).map(key => {
                      const [min, , max] = profile[key];
                      if (max === 0) return null;
                      return <SensorBar key={key} label={SENSOR_LABELS[key]} value={sd.current[key]} unit={SENSOR_UNITS[key]} min={min} nominal={profile[key][1]} max={max} />;
                    });
                  })()}
                </div>

                {/* trend charts */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                  <div className="card">
                    <div className="card-h">Health Trend</div>
                    <div style={{ height: 180 }}>
                      <ResponsiveContainer>
                        <AreaChart data={sensorHistory[selEquip.id] || []} margin={{ top: 5, right: 10, left: -15, bottom: 5 }}>
                          <CartesianGrid stroke="var(--line)" strokeDasharray="3 3" />
                          <XAxis dataKey="t" tick={{ fill: "var(--muted)", fontSize: 9 }} />
                          <YAxis domain={[0, 100]} tick={{ fill: "var(--muted)", fontSize: 9 }} />
                          <Tooltip contentStyle={{ background: "var(--panel)", border: "1px solid var(--line)", fontSize: 11 }} />
                          <Area type="monotone" dataKey="health" stroke="var(--ok)" fill="var(--ok)" fillOpacity={0.15} strokeWidth={2} />
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="card">
                    <div className="card-h">Temperature / Vibration</div>
                    <div style={{ height: 180 }}>
                      <ResponsiveContainer>
                        <LineChart data={sensorHistory[selEquip.id] || []} margin={{ top: 5, right: 10, left: -15, bottom: 5 }}>
                          <CartesianGrid stroke="var(--line)" strokeDasharray="3 3" />
                          <XAxis dataKey="t" tick={{ fill: "var(--muted)", fontSize: 9 }} />
                          <YAxis tick={{ fill: "var(--muted)", fontSize: 9 }} />
                          <Tooltip contentStyle={{ background: "var(--panel)", border: "1px solid var(--line)", fontSize: 11 }} />
                          <Line type="monotone" dataKey="temp" stroke="var(--danger)" strokeWidth={1.5} dot={false} name="Temp (°C)" />
                          <Line type="monotone" dataKey="vibration" stroke="var(--accent)" strokeWidth={1.5} dot={false} name="Vibr (mm/s)" />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="card">
                    <div className="card-h">Pressure / Current</div>
                    <div style={{ height: 180 }}>
                      <ResponsiveContainer>
                        <LineChart data={sensorHistory[selEquip.id] || []} margin={{ top: 5, right: 10, left: -15, bottom: 5 }}>
                          <CartesianGrid stroke="var(--line)" strokeDasharray="3 3" />
                          <XAxis dataKey="t" tick={{ fill: "var(--muted)", fontSize: 9 }} />
                          <YAxis tick={{ fill: "var(--muted)", fontSize: 9 }} />
                          <Tooltip contentStyle={{ background: "var(--panel)", border: "1px solid var(--line)", fontSize: 11 }} />
                          <Line type="monotone" dataKey="pressure" stroke="#4C9BE8" strokeWidth={1.5} dot={false} name="Press (bar)" />
                          <Line type="monotone" dataKey="current" stroke="var(--warn)" strokeWidth={1.5} dot={false} name="Current (A)" />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  <div className="card">
                    <div className="card-h">Maintenance History</div>
                    <div style={{ maxHeight: 180, overflow: "auto" }}>
                      {(selEquip.maintenanceLog || []).length === 0 && <p className="hint">No maintenance records.</p>}
                      <table className="bt">
                        <thead><tr><th>Date</th><th>Type</th><th>Description</th><th>Cost</th></tr></thead>
                        <tbody>
                          {(selEquip.maintenanceLog || []).map(m => (
                            <tr key={m.id}>
                              <td style={{ fontFamily: "var(--mono)", fontSize: 10, whiteSpace: "nowrap" }}>{m.date}</td>
                              <td><span style={{ fontFamily: "var(--mono)", fontSize: 9, padding: "1px 4px", borderRadius: 2, background: m.type === "Corrective" ? "rgba(224,101,82,.15)" : "rgba(98,210,111,.1)", color: m.type === "Corrective" ? "var(--danger)" : "var(--ok)" }}>{m.type}</span></td>
                              <td style={{ fontSize: 11 }}>{m.desc}</td>
                              <td style={{ fontFamily: "var(--mono)", fontSize: 11 }}>${m.cost?.toLocaleString()}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </>
            )}
          </div>
        )}

        {/* ============ INSPECTOR ============ */}
        <div className="insp">
          <div className="stabs"><button className="stab on" style={{ cursor: "default" }}>{selEquip ? "Equipment" : selUtil ? "Utility" : selConn ? "Connection" : "Inspector"}</button></div>
          <div className="sbody">
            {!selected && (
              <>
                <div className="sec">Nothing selected</div>
                <p className="hint">Select equipment to see live sensor data, health score, and predictive maintenance info. Use the Dashboard for a facility overview or Analytics for deep-dive per-asset.</p>
              </>
            )}

            {selEquip && (
              <EquipInspector eq={selEquip} patch={p => patchEquip(selEquip.id, p)}
                connections={project.connections.filter(c => c.fromId === selEquip.id || c.toId === selEquip.id)}
                nodes={nodes} onSelectConn={id => setSelected({ kind: "connection", id })}
                onRemove={removeSelected} sensorData={sensorData[selEquip.id]} />
            )}

            {selUtil && (
              <>
                <div className="sec">Utility node</div>
                <div className="fg2">
                  <Fld label="Tag"><input value={selUtil.tag} onChange={e => patchUtil(selUtil.id, { tag: e.target.value })} /></Fld>
                  <Fld label="Service"><select value={selUtil.type} onChange={e => patchUtil(selUtil.id, { type: e.target.value })}>{Object.entries(UTILITY_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></Fld>
                </div>
                <Fld label="Name"><input value={selUtil.name} onChange={e => patchUtil(selUtil.id, { name: e.target.value })} /></Fld>
                <Fld label="Capacity"><input value={selUtil.capacity || ""} onChange={e => patchUtil(selUtil.id, { capacity: e.target.value })} /></Fld>
                <div className="sec">Serves</div>
                {project.connections.filter(c => c.fromId === selUtil.id || c.toId === selUtil.id).map(c => {
                  const other = nodes[c.fromId === selUtil.id ? c.toId : c.fromId];
                  return <div key={c.id} className="cc" onClick={() => setSelected({ kind: "connection", id: c.id })}><span className="rdot" style={{ background: UTILITY_TYPES[c.utilityType]?.color }} /><span className="rtag" style={{ fontSize: 10 }}>{other?.tag}</span><span className="rname">{c.medium || "—"}</span></div>;
                })}
                <button className="btn dng" style={{ width: "100%", marginTop: 10 }} onClick={removeSelected}>DELETE</button>
              </>
            )}

            {selConn && (
              <>
                <div className="sec">Connection</div>
                <p className="hint" style={{ fontFamily: "var(--mono)" }}>{nodes[selConn.fromId]?.tag} → {nodes[selConn.toId]?.tag}</p>
                <Fld label="Service"><select value={selConn.utilityType} onChange={e => patchConn(selConn.id, { utilityType: e.target.value })}>{Object.entries(UTILITY_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></Fld>
                <Fld label="Medium"><input value={selConn.medium} onChange={e => patchConn(selConn.id, { medium: e.target.value })} /></Fld>
                <Fld label="Size"><input value={selConn.size} onChange={e => patchConn(selConn.id, { size: e.target.value })} /></Fld>
                <Fld label="Notes"><textarea rows={3} value={selConn.notes} onChange={e => patchConn(selConn.id, { notes: e.target.value })} /></Fld>
                <button className="btn dng" style={{ width: "100%" }} onClick={removeSelected}>DELETE</button>
              </>
            )}
          </div>
        </div>
      </div>

      {/* ============ GITHUB MODAL ============ */}
      {ghOpen && (
        <div className="mback" onMouseDown={e => { if (e.target === e.currentTarget) setGhOpen(false); }}>
          <div className="mdl">
            <h3>GitHub Repository</h3>
            <p className="sub">Version your digital twin by committing JSON snapshots. Create a fine-grained PAT with <b>Contents: read &amp; write</b>.</p>
            <div className="fg2">
              <Fld label="Owner"><input placeholder="username" value={project.github.owner} onChange={e => setProject(p => ({ ...p, github: { ...p.github, owner: e.target.value.trim() } }))} /></Fld>
              <Fld label="Repo"><input placeholder="digital-twin" value={project.github.repo} onChange={e => setProject(p => ({ ...p, github: { ...p.github, repo: e.target.value.trim() } }))} /></Fld>
            </div>
            <div className="fg2">
              <Fld label="Branch"><input value={project.github.branch} onChange={e => setProject(p => ({ ...p, github: { ...p.github, branch: e.target.value.trim() } }))} /></Fld>
              <Fld label="Path"><input value={project.github.path} onChange={e => setProject(p => ({ ...p, github: { ...p.github, path: e.target.value.trim() } }))} /></Fld>
            </div>
            <Fld label="Token"><input type="password" placeholder="github_pat_…" value={project.github.token} onChange={e => setProject(p => ({ ...p, github: { ...p.github, token: e.target.value.trim() } }))} /></Fld>
            <div style={{ display: "flex", gap: 7, marginTop: 10 }}>
              <button className="btn on" disabled={ghBusy || !project.github.owner || !project.github.repo || !project.github.token} onClick={ghPush}>{ghBusy ? "WORKING…" : "PUSH SNAPSHOT"}</button>
              <div style={{ flex: 1 }} />
              <button className="btn gh" onClick={() => setGhOpen(false)}>CLOSE</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* =================== EQUIPMENT INSPECTOR ========================== */

function EquipInspector({ eq, patch, connections, nodes, onSelectConn, onRemove, sensorData }) {
  const [tab, setTab] = useState("health");
  const setBom = bom => patch({ bom });
  const patchRow = (id, p) => setBom(eq.bom.map(r => r.id === id ? { ...r, ...p } : r));
  const addLog = () => patch({ maintenanceLog: [...(eq.maintenanceLog || []), { id: uid(), date: new Date().toISOString().slice(0, 10), type: "Preventive", desc: "", cost: 0 }] });
  const patchLog = (id, p) => patch({ maintenanceLog: (eq.maintenanceLog || []).map(m => m.id === id ? { ...m, ...p } : m) });

  const profile = SENSOR_PROFILES[eq.type] || SENSOR_PROFILES.Other;
  const sd = sensorData;

  return (
    <>
      <div style={{ display: "flex", gap: 3, marginBottom: 10, flexWrap: "wrap" }}>
        {["health", "details", "bom", "maint", "util"].map(t => (
          <button key={t} className={`btn sm ${tab === t ? "on" : "gh"}`} onClick={() => setTab(t)}>
            {t === "health" ? "HEALTH" : t === "bom" ? `BOM(${eq.bom.length})` : t === "maint" ? "MAINT" : t === "util" ? `UTIL(${connections.length})` : "DETAILS"}
          </button>
        ))}
      </div>

      {tab === "health" && (
        <>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
            <HealthGauge value={sd?.health ?? 100} size={80} />
            <div>
              <div style={{ fontFamily: "var(--mono)", fontWeight: 700, fontSize: 14 }}>{eq.tag}</div>
              <div style={{ fontSize: 11, color: "var(--muted)" }}>{eq.name}</div>
              <div style={{ fontFamily: "var(--mono)", fontSize: 10, color: "var(--faint)", marginTop: 3 }}>
                {eq.operatingHours?.toLocaleString() || 0} operating hrs
              </div>
              {sd?.failureDays && (
                <div style={{ fontFamily: "var(--mono)", fontSize: 10, color: "var(--danger)", marginTop: 3 }}>
                  ⚠ ~{sd.failureDays} days to maintenance
                </div>
              )}
            </div>
          </div>
          <div className="sec">Live sensors</div>
          {sd ? Object.keys(profile).map(key => {
            const [min, , max] = profile[key];
            if (max === 0) return null;
            return <SensorBar key={key} label={SENSOR_LABELS[key]} value={sd.current[key]} unit={SENSOR_UNITS[key]} min={min} nominal={profile[key][1]} max={max} />;
          }) : <p className="hint">Waiting for data…</p>}
        </>
      )}

      {tab === "details" && (
        <>
          <div className="fg2">
            <Fld label="Tag"><input value={eq.tag} onChange={e => patch({ tag: e.target.value })} /></Fld>
            <Fld label="Type"><select value={eq.type} onChange={e => patch({ type: e.target.value })}>{[...new Set([eq.type, ...EQUIP_TYPES])].map(t => <option key={t}>{t}</option>)}</select></Fld>
          </div>
          <Fld label="Name"><input value={eq.name} onChange={e => patch({ name: e.target.value })} /></Fld>
          <div className="fg2">
            <Fld label="Manufacturer"><input value={eq.manufacturer} onChange={e => patch({ manufacturer: e.target.value })} /></Fld>
            <Fld label="Model"><input value={eq.model} onChange={e => patch({ model: e.target.value })} /></Fld>
          </div>
          <div className="fg2">
            <Fld label="Power (kW)"><input type="number" value={eq.power} onChange={e => patch({ power: e.target.value === "" ? "" : parseFloat(e.target.value) })} /></Fld>
            <Fld label="Area"><input value={eq.area} onChange={e => patch({ area: e.target.value })} /></Fld>
          </div>
          <div className="fg2">
            <Fld label="Install date"><input type="date" value={eq.installDate || ""} onChange={e => patch({ installDate: e.target.value })} /></Fld>
            <Fld label="Operating hrs"><input type="number" value={eq.operatingHours || 0} onChange={e => patch({ operatingHours: parseInt(e.target.value) || 0 })} /></Fld>
          </div>
          <div className="sec">Footprint (m)</div>
          <div className="fg3">
            <Fld label="W"><input type="number" step="0.1" value={eq.w} onChange={e => patch({ w: parseFloat(e.target.value) || 0.2 })} /></Fld>
            <Fld label="D"><input type="number" step="0.1" value={eq.d} onChange={e => patch({ d: parseFloat(e.target.value) || 0.2 })} /></Fld>
            <Fld label="H"><input type="number" step="0.1" value={eq.h} onChange={e => patch({ h: parseFloat(e.target.value) || 0.2 })} /></Fld>
          </div>
          <div className="fg2">
            <Fld label="X"><input type="number" step="0.1" value={round2(eq.x)} onChange={e => patch({ x: parseFloat(e.target.value) || 0, placed: true })} /></Fld>
            <Fld label="Y"><input type="number" step="0.1" value={round2(eq.y)} onChange={e => patch({ y: parseFloat(e.target.value) || 0, placed: true })} /></Fld>
          </div>
          <button className="btn dng" style={{ width: "100%", marginTop: 8 }} onClick={onRemove}>DELETE</button>
        </>
      )}

      {tab === "bom" && (
        <>
          <div className="sec">Bill of materials</div>
          <table className="bt">
            <thead><tr><th>Part no.</th><th>Desc</th><th style={{ width: 36 }}>Qty</th><th style={{ width: 20 }} /></tr></thead>
            <tbody>
              {eq.bom.map(r => (
                <tr key={r.id}>
                  <td><input style={{ fontFamily: "var(--mono)" }} value={r.pn} onChange={e => patchRow(r.id, { pn: e.target.value })} /></td>
                  <td><input value={r.desc} onChange={e => patchRow(r.id, { desc: e.target.value })} /></td>
                  <td><input type="number" value={r.qty} onChange={e => patchRow(r.id, { qty: parseFloat(e.target.value) || 0 })} /></td>
                  <td><button className="btn sm gh dng" style={{ padding: "1px 4px" }} onClick={() => setBom(eq.bom.filter(x => x.id !== r.id))}>✕</button></td>
                </tr>
              ))}
            </tbody>
          </table>
          {!eq.bom.length && <p className="hint">No parts recorded.</p>}
          <button className="btn sm" style={{ marginTop: 6 }} onClick={() => setBom([...eq.bom, { id: uid(), pn: "", desc: "", qty: 1 }])}>+ ADD PART</button>
        </>
      )}

      {tab === "maint" && (
        <>
          <div className="sec">Maintenance log</div>
          {(eq.maintenanceLog || []).map(m => (
            <div key={m.id} style={{ padding: "6px 0", borderBottom: "1px solid var(--line)" }}>
              <div className="fg2">
                <Fld label="Date"><input type="date" value={m.date} onChange={e => patchLog(m.id, { date: e.target.value })} /></Fld>
                <Fld label="Type"><select value={m.type} onChange={e => patchLog(m.id, { type: e.target.value })}><option>Preventive</option><option>Corrective</option><option>Inspection</option></select></Fld>
              </div>
              <Fld label="Description"><input value={m.desc} onChange={e => patchLog(m.id, { desc: e.target.value })} /></Fld>
              <Fld label="Cost ($)"><input type="number" value={m.cost || 0} onChange={e => patchLog(m.id, { cost: parseFloat(e.target.value) || 0 })} /></Fld>
            </div>
          ))}
          {!(eq.maintenanceLog || []).length && <p className="hint">No records. Add entries to track maintenance costs and predict future needs.</p>}
          <button className="btn sm" style={{ marginTop: 6 }} onClick={addLog}>+ ADD RECORD</button>
        </>
      )}

      {tab === "util" && (
        <>
          <div className="sec">Utility connections</div>
          {connections.map(c => {
            const other = nodes[c.fromId === eq.id ? c.toId : c.fromId];
            return <div key={c.id} className="cc" onClick={() => onSelectConn(c.id)}><span className="rdot" style={{ background: UTILITY_TYPES[c.utilityType]?.color }} /><span className="rtag" style={{ fontSize: 10 }}>{c.fromId === eq.id ? "→" : "←"} {other?.tag}</span><span className="rname">{c.medium || UTILITY_TYPES[c.utilityType]?.label}</span></div>;
          })}
          {!connections.length && <p className="hint">Not connected. Use ⇄ CONNECT in toolbar.</p>}
        </>
      )}
    </>
  );
}
