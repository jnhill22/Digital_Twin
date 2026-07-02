import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as THREE from "three";
import * as XLSX from "xlsx";
import Papa from "papaparse";

/* ================================================================== */
/*  FacilityTwin — Level 1: Spatial Model + Material Flow              */
/*  Workstation types · Corner resize · Material routing · DXF plans  */
/*  Equipment BOM · Utility connections · GitHub version control       */
/* ================================================================== */

const uid = () => Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-3);

/* ==================== WORKSTATION TYPE SYSTEM ====================== */

const WS_TYPES = {
  mixer:          { label: "Mixer",                 short: "MIX",  color: "#4C9BE8", dW: 2.4, dD: 2.4, dH: 2.6, icon: "mixer",
    props: [{ key: "capacity", label: "Capacity (L)", type: "number" }, { key: "speed", label: "Max RPM", type: "number" }, { key: "batchTime", label: "Batch time (min)", type: "number" }] },
  dryVacuum:      { label: "Dry vacuum transfer",   short: "VAC",  color: "#9D8DF1", dW: 1.5, dD: 1.0, dH: 2.0, icon: "vacuum",
    props: [{ key: "sourceType", label: "Source", type: "select", opts: ["Bucket", "Silo", "Hopper", "Bag dump", "Super sack"] }, { key: "material", label: "Material", type: "text" }, { key: "flowRate", label: "Flow rate (kg/hr)", type: "number" }] },
  liquidSupply:   { label: "Liquid supply",          short: "LIQ",  color: "#45C4B0", dW: 1.2, dD: 1.0, dH: 1.5, icon: "liquid",
    props: [{ key: "sourceType", label: "Source", type: "select", opts: ["Drum", "IBC tote", "Water (infinite)", "Tank", "Pipe feed"] }, { key: "material", label: "Material", type: "text" }, { key: "flowRate", label: "Flow rate (L/min)", type: "number" }, { key: "temp", label: "Temp (°C)", type: "number" }] },
  chiller:        { label: "Chiller / Cooler",       short: "CHL",  color: "#5BC0EB", dW: 2.0, dD: 2.0, dH: 1.8, icon: "chiller",
    props: [{ key: "capacity", label: "Capacity (kW)", type: "number" }, { key: "tempRange", label: "Temp range (°C)", type: "text" }, { key: "refrigerant", label: "Refrigerant", type: "text" }] },
  dryingRack:     { label: "Drying rack",            short: "DRY",  color: "#EFCB4F", dW: 3.0, dD: 1.5, dH: 2.0, icon: "rack",
    props: [{ key: "shelfCount", label: "Shelves", type: "number" }, { key: "dryTime", label: "Dry time (hr)", type: "number" }, { key: "capacity", label: "Capacity (units)", type: "number" }] },
  vibrationTable: { label: "Vibration table",        short: "VIB",  color: "#F08C3A", dW: 2.5, dD: 1.5, dH: 1.0, icon: "vibe",
    props: [{ key: "frequency", label: "Frequency (Hz)", type: "number" }, { key: "amplitude", label: "Amplitude (mm)", type: "number" }, { key: "tableSize", label: "Table size (m)", type: "text" }] },
  silo:           { label: "Silo / Hopper",          short: "SIL",  color: "#B8A9C9", dW: 1.8, dD: 1.8, dH: 4.0, icon: "silo",
    props: [{ key: "capacity", label: "Capacity (kg)", type: "number" }, { key: "material", label: "Stored material", type: "text" }] },
  tank:           { label: "Tank / Vessel",           short: "TNK",  color: "#7EC8B0", dW: 2.0, dD: 2.0, dH: 2.5, icon: "tank",
    props: [{ key: "capacity", label: "Capacity (L)", type: "number" }, { key: "material", label: "Contents", type: "text" }, { key: "jacketed", label: "Jacketed", type: "select", opts: ["No", "Yes — heating", "Yes — cooling"] }] },
  conveyor:       { label: "Conveyor / Belt",         short: "CNV",  color: "#8C9BA5", dW: 4.0, dD: 0.8, dH: 1.0, icon: "conveyor",
    props: [{ key: "length", label: "Length (m)", type: "number" }, { key: "speed", label: "Speed (m/min)", type: "number" }, { key: "beltType", label: "Belt type", type: "text" }] },
  workbench:      { label: "Workbench / Station",     short: "WRK",  color: "#D4A574", dW: 2.0, dD: 1.0, dH: 0.9, icon: "bench",
    props: [{ key: "purpose", label: "Purpose", type: "text" }, { key: "operators", label: "Operators", type: "number" }] },
  custom:         { label: "Custom object",            short: "OBJ",  color: "#7E8B95", dW: 2.0, dD: 1.5, dH: 1.8, icon: "box",
    props: [{ key: "notes", label: "Notes", type: "text" }] },
};

const MATERIAL_TYPES = {
  dry:      { label: "Dry ingredient",  color: "#EFCB4F", dash: "" },
  liquid:   { label: "Liquid",          color: "#4C9BE8", dash: "" },
  mixed:    { label: "Mixed / batter",  color: "#62D26F", dash: "" },
  cooled:   { label: "Cooled product",  color: "#5BC0EB", dash: "6 3" },
  finished: { label: "Finished product", color: "#F08C3A", dash: "" },
  waste:    { label: "Waste / return",  color: "#8C9BA5", dash: "4 4" },
  custom:   { label: "Custom",          color: "#B8A9C9", dash: "3 3" },
};

const UTILITY_TYPES = {
  hvac:    { label: "HVAC",            color: "#45C4B0", short: "HVAC" },
  chw:     { label: "Chilled water",   color: "#4C9BE8", short: "CHW" },
  hw:      { label: "Hot water/steam", color: "#E06552", short: "STM" },
  power:   { label: "Electrical",      color: "#EFCB4F", short: "PWR" },
  compair: { label: "Compressed air",  color: "#9D8DF1", short: "CDA" },
  drain:   { label: "Drain / waste",   color: "#8C9BA5", short: "DRN" },
};

/* ====================== WORKSTATION ICONS ========================== */
/* Each draws inside a normalized [-1,-1] to [1,1] box scaled to w×d  */

function wsIcon(type, x, y, w, d, sw, selected) {
  const t = WS_TYPES[type] || WS_TYPES.custom;
  const col = selected ? "#F08C3A" : t.color;
  const o = selected ? 1 : 0.7;
  const cx = x, cy = y;
  const hw = w / 2, hd = d / 2;
  const s = Math.min(hw, hd) * 0.55;

  switch (t.icon) {
    case "mixer": return (<>
      <circle cx={cx} cy={cy} r={s} fill="none" stroke={col} strokeWidth={sw * 2} opacity={o} />
      <line x1={cx} y1={cy - s * 0.6} x2={cx} y2={cy + s * 0.6} stroke={col} strokeWidth={sw * 1.5} opacity={o} />
      <line x1={cx - s * 0.35} y1={cy + s * 0.2} x2={cx + s * 0.35} y2={cy - s * 0.2} stroke={col} strokeWidth={sw * 1.5} opacity={o} />
      <line x1={cx - s * 0.35} y1={cy - s * 0.2} x2={cx + s * 0.35} y2={cy + s * 0.2} stroke={col} strokeWidth={sw * 1.5} opacity={o} />
    </>);
    case "vacuum": return (<>
      <line x1={cx} y1={cy - s * 0.7} x2={cx} y2={cy + s * 0.4} stroke={col} strokeWidth={sw * 2} opacity={o} />
      <polygon points={`${cx - s * 0.5},${cy + s * 0.4} ${cx + s * 0.5},${cy + s * 0.4} ${cx},${cy - s * 0.2}`} fill={col} opacity={o * 0.5} />
      <line x1={cx - s * 0.3} y1={cy - s * 0.7} x2={cx + s * 0.3} y2={cy - s * 0.7} stroke={col} strokeWidth={sw * 1.5} opacity={o} />
    </>);
    case "liquid": return (<>
      <circle cx={cx} cy={cy} r={s * 0.8} fill="none" stroke={col} strokeWidth={sw * 2} opacity={o} />
      <path d={`M${cx - s * 0.4} ${cy - s * 0.1} Q${cx - s * 0.2} ${cy - s * 0.35} ${cx} ${cy - s * 0.1} Q${cx + s * 0.2} ${cy + s * 0.15} ${cx + s * 0.4} ${cy - s * 0.1}`}
        fill="none" stroke={col} strokeWidth={sw * 1.5} opacity={o} />
    </>);
    case "chiller": return (<>
      <line x1={cx} y1={cy - s * 0.6} x2={cx} y2={cy + s * 0.6} stroke={col} strokeWidth={sw * 2} opacity={o} />
      <line x1={cx - s * 0.4} y1={cy - s * 0.3} x2={cx + s * 0.4} y2={cy - s * 0.3} stroke={col} strokeWidth={sw * 1.2} opacity={o} />
      <line x1={cx - s * 0.4} y1={cy + s * 0.3} x2={cx + s * 0.4} y2={cy + s * 0.3} stroke={col} strokeWidth={sw * 1.2} opacity={o} />
      <circle cx={cx} cy={cy} r={s * 0.15} fill={col} opacity={o * 0.5} />
    </>);
    case "rack": {
      const n = 4;
      return (<>{Array.from({ length: n }, (_, i) => {
        const fy = cy - s * 0.6 + (s * 1.2 / (n - 1)) * i;
        return <line key={i} x1={cx - s * 0.5} y1={fy} x2={cx + s * 0.5} y2={fy} stroke={col} strokeWidth={sw * 1.3} opacity={o} />;
      })}</>);
    }
    case "vibe": return (<>
      <path d={`M${cx - s * 0.6} ${cy + s * 0.3} L${cx - s * 0.3} ${cy + s * 0.6} L${cx} ${cy + s * 0.3} L${cx + s * 0.3} ${cy + s * 0.6} L${cx + s * 0.6} ${cy + s * 0.3}`}
        fill="none" stroke={col} strokeWidth={sw * 1.8} opacity={o} />
      <line x1={cx - s * 0.6} y1={cy - s * 0.2} x2={cx + s * 0.6} y2={cy - s * 0.2} stroke={col} strokeWidth={sw * 2.5} opacity={o} />
    </>);
    case "silo": return (<>
      <rect x={cx - s * 0.35} y={cy - s * 0.7} width={s * 0.7} height={s * 1.1} fill="none" stroke={col} strokeWidth={sw * 2} opacity={o} rx={sw} />
      <path d={`M${cx - s * 0.35} ${cy + s * 0.4} L${cx} ${cy + s * 0.75} L${cx + s * 0.35} ${cy + s * 0.4}`} fill="none" stroke={col} strokeWidth={sw * 2} opacity={o} />
    </>);
    case "tank": return (<>
      <ellipse cx={cx} cy={cy - s * 0.5} rx={s * 0.5} ry={s * 0.2} fill="none" stroke={col} strokeWidth={sw * 2} opacity={o} />
      <line x1={cx - s * 0.5} y1={cy - s * 0.5} x2={cx - s * 0.5} y2={cy + s * 0.4} stroke={col} strokeWidth={sw * 2} opacity={o} />
      <line x1={cx + s * 0.5} y1={cy - s * 0.5} x2={cx + s * 0.5} y2={cy + s * 0.4} stroke={col} strokeWidth={sw * 2} opacity={o} />
      <ellipse cx={cx} cy={cy + s * 0.4} rx={s * 0.5} ry={s * 0.2} fill="none" stroke={col} strokeWidth={sw * 2} opacity={o} />
    </>);
    case "conveyor": return (<>
      <line x1={cx - hw * 0.7} y1={cy} x2={cx + hw * 0.7} y2={cy} stroke={col} strokeWidth={sw * 2.5} opacity={o} />
      <polygon points={`${cx + hw * 0.5},${cy - s * 0.3} ${cx + hw * 0.7},${cy} ${cx + hw * 0.5},${cy + s * 0.3}`} fill={col} opacity={o * 0.6} />
    </>);
    case "bench": return (<>
      <rect x={cx - hw * 0.6} y={cy - hd * 0.4} width={hw * 1.2} height={hd * 0.8} fill="none" stroke={col} strokeWidth={sw * 2} opacity={o} rx={sw} />
    </>);
    default: return (<>
      <line x1={cx - s * 0.4} y1={cy - s * 0.4} x2={cx + s * 0.4} y2={cy + s * 0.4} stroke={col} strokeWidth={sw * 1.5} opacity={o} />
      <line x1={cx + s * 0.4} y1={cy - s * 0.4} x2={cx - s * 0.4} y2={cy + s * 0.4} stroke={col} strokeWidth={sw * 1.5} opacity={o} />
    </>);
  }
}

/* ========================== DXF PARSER ============================== */

function parseDXF(text) {
  const rows = text.split(/\r\n|\r|\n/), pairs = [];
  for (let i = 0; i + 1 < rows.length; i += 2) { const c = parseInt(rows[i], 10); if (!isNaN(c)) pairs.push([c, rows[i + 1].trim()]); }
  let section = "", cur = null; const raw = [];
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
  for (const e of ents) { if (e.pts) e.pts.forEach(([x, y]) => grow(x, y)); if (e.kind === "circle") { grow(e.cx - e.r, e.cy - e.r); grow(e.cx + e.r, e.cy + e.r); } if (e.kind === "text") grow(e.x, e.y); }
  if (!isFinite(minX)) { minX = 0; minY = 0; maxX = 40; maxY = 25; }
  let unitScale = 1, assumedUnits = "m"; const span = Math.max(maxX - minX, maxY - minY);
  if (span > 2000) { unitScale = 0.001; assumedUnits = "mm"; } else if (span > 500) { unitScale = 0.01; assumedUnits = "cm"; }
  if (unitScale !== 1) { for (const e of ents) { if (e.pts) e.pts = e.pts.map(([x, y]) => [x * unitScale, y * unitScale]); if (e.kind === "circle") { e.cx *= unitScale; e.cy *= unitScale; e.r *= unitScale; } if (e.kind === "text") { e.x *= unitScale; e.y *= unitScale; e.h *= unitScale; } } minX *= unitScale; minY *= unitScale; maxX *= unitScale; maxY *= unitScale; }
  return { entities: ents, bounds: { minX, minY, maxX, maxY }, assumedUnits, count: ents.length };
}

/* ===================== EQUIPMENT LIST IMPORT ======================== */

const COL_ALIASES = {
  tag: ["tag", "equipment tag", "eq tag", "equipment id", "id", "item", "item no"],
  name: ["name", "description", "equipment name", "desc", "title"],
  type: ["type", "category", "equipment type", "class", "station type", "workstation"],
  manufacturer: ["manufacturer", "mfr", "mfg", "vendor", "make", "oem"],
  model: ["model", "model no", "model number"],
  power: ["power", "kw", "power (kw)", "hp"],
  x: ["x", "loc x", "pos x"], y: ["y", "loc y", "pos y"],
  w: ["width", "w"], d: ["depth", "d", "length", "l"], h: ["height", "h"],
  area: ["area", "room", "zone", "location"],
};

function mapColumns(headers) {
  const map = {}, norm = headers.map(h => String(h || "").trim().toLowerCase());
  for (const key of Object.keys(COL_ALIASES)) { const idx = norm.findIndex(h => COL_ALIASES[key].includes(h)); if (idx >= 0) map[key] = headers[idx]; }
  return map;
}

const guessWsType = (name) => {
  const n = (name || "").toLowerCase();
  if (/mix|blend|agit/i.test(n)) return "mixer";
  if (/vacuum|vac|suck|transfer/i.test(n)) return "dryVacuum";
  if (/liquid|water|pump|drum/i.test(n)) return "liquidSupply";
  if (/chill|cool|refrig/i.test(n)) return "chiller";
  if (/dry|rack|cure/i.test(n)) return "dryingRack";
  if (/vib|shak|table/i.test(n)) return "vibrationTable";
  if (/silo|hopper|bin/i.test(n)) return "silo";
  if (/tank|vessel|reactor/i.test(n)) return "tank";
  if (/conveyor|belt/i.test(n)) return "conveyor";
  if (/bench|station|table/i.test(n)) return "workbench";
  return "custom";
};

function rowsToEquipment(rows) {
  if (!rows.length) return [];
  const map = mapColumns(Object.keys(rows[0]));
  return rows.filter(r => Object.values(r).some(v => String(v || "").trim() !== "")).map((r, i) => {
    const num = k => { const v = parseFloat(r[map[k]]); return Number.isFinite(v) ? v : undefined; };
    const name = String(r[map.name] ?? "").trim() || "Unnamed";
    const rawType = String(r[map.type] ?? "").trim();
    const wsType = Object.keys(WS_TYPES).find(k => WS_TYPES[k].label.toLowerCase() === rawType.toLowerCase()) || guessWsType(name);
    const ws = WS_TYPES[wsType];
    const x = num("x"), y = num("y");
    return { id: uid(), tag: String(r[map.tag] ?? `${ws.short}-${String(i + 1).padStart(3, "0")}`).trim(),
      name, wsType, manufacturer: String(r[map.manufacturer] ?? "").trim(), model: String(r[map.model] ?? "").trim(),
      power: num("power") ?? "", area: String(r[map.area] ?? "").trim(),
      x, y, placed: x !== undefined && y !== undefined,
      w: num("w") ?? ws.dW, d: num("d") ?? ws.dD, h: num("h") ?? ws.dH, bom: [], custom: {},
    };
  });
}

function downloadFile(name, text, mime = "application/json") {
  const blob = new Blob([text], { type: mime }); const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = name; a.click(); URL.revokeObjectURL(url);
}

/* ============================ DEMO DATA ============================= */

function demoProject() {
  const wall = (pts, closed = true) => ({ kind: "poly", pts, closed });
  const entities = [
    wall([[0, 0], [36, 0], [36, 22], [0, 22]]),
    wall([[14, 0], [14, 10]], false),
    wall([[0, 12], [14, 12]], false),
    wall([[24, 22], [24, 14], [36, 14]], false),
    { kind: "text", x: 5, y: 5, h: 0.7, text: "DRY STAGING" },
    { kind: "text", x: 5, y: 16, h: 0.7, text: "LIQUID PREP" },
    { kind: "text", x: 20, y: 5, h: 0.7, text: "MIXING" },
    { kind: "text", x: 20, y: 16, h: 0.7, text: "PROCESSING" },
    { kind: "text", x: 29, y: 18, h: 0.7, text: "FINISHING" },
  ];
  const mk = (o) => ({ id: uid(), placed: true, bom: [], custom: {}, manufacturer: "", model: "", power: "", area: "", ...o });
  const equipment = [
    mk({ tag: "SIL-01", name: "Sugar silo", wsType: "silo", x: 3, y: 3, w: 1.8, d: 1.8, h: 4, area: "Dry Staging", custom: { capacity: "2000", material: "Granulated sugar" } }),
    mk({ tag: "SIL-02", name: "Flour silo", wsType: "silo", x: 7, y: 3, w: 1.8, d: 1.8, h: 4, area: "Dry Staging", custom: { capacity: "1500", material: "AP flour" } }),
    mk({ tag: "VAC-01", name: "Dry vacuum — sugar", wsType: "dryVacuum", x: 11, y: 3, w: 1.5, d: 1, h: 2, area: "Dry Staging", custom: { sourceType: "Silo", material: "Sugar", flowRate: "500" } }),
    mk({ tag: "VAC-02", name: "Dry vacuum — flour", wsType: "dryVacuum", x: 11, y: 7, w: 1.5, d: 1, h: 2, area: "Dry Staging", custom: { sourceType: "Silo", material: "Flour", flowRate: "400" } }),
    mk({ tag: "TNK-01", name: "Water supply", wsType: "liquidSupply", x: 4, y: 15, w: 1.2, d: 1, h: 1.5, area: "Liquid Prep", custom: { sourceType: "Water (infinite)", material: "Filtered water", flowRate: "20" } }),
    mk({ tag: "TNK-02", name: "Oil drum station", wsType: "liquidSupply", x: 8, y: 15, w: 1.2, d: 1, h: 1.5, area: "Liquid Prep", custom: { sourceType: "Drum", material: "Vegetable oil", flowRate: "8" } }),
    mk({ tag: "MIX-01", name: "Planetary mixer 200L", wsType: "mixer", x: 18, y: 3, w: 2.4, d: 2.4, h: 2.6, area: "Mixing", manufacturer: "Hobart", model: "HL800", power: 22, custom: { capacity: "200", speed: "450", batchTime: "15" },
      bom: [{ id: uid(), pn: "MIX-BOWL-200", desc: "200L mixing bowl", qty: 2 }, { id: uid(), pn: "MIX-WHIP-200", desc: "Wire whip attachment", qty: 1 }, { id: uid(), pn: "MIX-PADDLE-200", desc: "Flat paddle", qty: 1 }] }),
    mk({ tag: "MIX-02", name: "High-shear mixer 100L", wsType: "mixer", x: 24, y: 3, w: 2, d: 2, h: 2.2, area: "Mixing", manufacturer: "Silverson", model: "FMX-100", power: 15, custom: { capacity: "100", speed: "3000", batchTime: "8" } }),
    mk({ tag: "VIB-01", name: "Vibration table A", wsType: "vibrationTable", x: 18, y: 12, w: 2.5, d: 1.5, h: 1, area: "Processing", custom: { frequency: "50", amplitude: "2", tableSize: "2.5 x 1.5" } }),
    mk({ tag: "VIB-02", name: "Vibration table B", wsType: "vibrationTable", x: 24, y: 12, w: 2.5, d: 1.5, h: 1, area: "Processing", custom: { frequency: "50", amplitude: "2", tableSize: "2.5 x 1.5" } }),
    mk({ tag: "CHL-01", name: "Walk-in chiller", wsType: "chiller", x: 30, y: 18, w: 3, d: 3, h: 2.4, area: "Finishing", custom: { capacity: "18", tempRange: "2-6", refrigerant: "R-404A" } }),
    mk({ tag: "DRY-01", name: "Drying rack bank", wsType: "dryingRack", x: 30, y: 11, w: 3, d: 1.5, h: 2, area: "Processing", custom: { shelfCount: "12", dryTime: "4", capacity: "240" } }),
  ];
  const find = t => equipment.find(e => e.tag === t).id;
  const mflow = (from, to, mat, label, rate) => ({ id: uid(), fromId: from, toId: to, materialType: mat, label: label || "", flowRate: rate || "", notes: "" });
  const materialFlows = [
    mflow(find("SIL-01"), find("VAC-01"), "dry", "Sugar transfer", "500 kg/hr"),
    mflow(find("SIL-02"), find("VAC-02"), "dry", "Flour transfer", "400 kg/hr"),
    mflow(find("VAC-01"), find("MIX-01"), "dry", "Sugar to mixer"),
    mflow(find("VAC-02"), find("MIX-01"), "dry", "Flour to mixer"),
    mflow(find("TNK-01"), find("MIX-01"), "liquid", "Water to mixer"),
    mflow(find("TNK-02"), find("MIX-02"), "liquid", "Oil to mixer"),
    mflow(find("MIX-01"), find("VIB-01"), "mixed", "Batter to table"),
    mflow(find("MIX-02"), find("VIB-02"), "mixed", "Batter to table"),
    mflow(find("VIB-01"), find("DRY-01"), "mixed", "Molded product"),
    mflow(find("VIB-02"), find("DRY-01"), "mixed", "Molded product"),
    mflow(find("DRY-01"), find("CHL-01"), "finished", "To chiller"),
  ];
  const utilities = [
    { id: uid(), tag: "MCC-01", name: "Motor control center", type: "power", x: 12, y: 19, capacity: "400 A" },
    { id: uid(), tag: "CDA-01", name: "Compressed air manifold", type: "compair", x: 12, y: 21, capacity: "60 CFM" },
  ];
  const utilConns = [
    { id: uid(), fromId: utilities[0].id, toId: find("MIX-01"), utilityType: "power", medium: "480V 3ph", size: "22 kW", notes: "" },
    { id: uid(), fromId: utilities[0].id, toId: find("MIX-02"), utilityType: "power", medium: "480V 3ph", size: "15 kW", notes: "" },
    { id: uid(), fromId: utilities[0].id, toId: find("CHL-01"), utilityType: "power", medium: "480V 3ph", size: "8 kW", notes: "" },
  ];
  return {
    name: "Confectionery Line 2", site: "", revision: "A",
    floorplan: { entities, bounds: { minX: 0, minY: 0, maxX: 36, maxY: 22 }, assumedUnits: "m", source: "demo.dxf" },
    equipment, utilities, utilityConnections: utilConns, materialFlows,
    github: { owner: "", repo: "", branch: "main", path: "digital-twin/project.json", token: "" },
  };
}

function emptyProject() {
  return { name: "Untitled facility", site: "", revision: "A", floorplan: null,
    equipment: [], utilities: [], utilityConnections: [], materialFlows: [],
    github: { owner: "", repo: "", branch: "main", path: "digital-twin/project.json", token: "" },
  };
}

/* ============================== STYLES ============================== */

const CSS = `
:root{--bg:#101418;--panel:#171C21;--panel2:#1D242B;--line:#2A333C;--line2:#38434E;--text:#DCE3E9;--muted:#7E8B95;--faint:#55616B;--accent:#F08C3A;--accent-dim:#7a4a22;--danger:#E06552;--ok:#62D26F;--warn:#EFCB4F;--mono:ui-monospace,'SF Mono','Cascadia Code',Consolas,monospace;--sans:'Segoe UI',system-ui,-apple-system,Roboto,sans-serif}
.ft *{box-sizing:border-box}
.ft{position:fixed;inset:0;background:var(--bg);color:var(--text);font-family:var(--sans);font-size:13px;display:flex;flex-direction:column;overflow:hidden}
.ft ::-webkit-scrollbar{width:7px;height:7px}.ft ::-webkit-scrollbar-thumb{background:var(--line2);border-radius:4px}

.tb{display:flex;align-items:stretch;border-bottom:1px solid var(--line);background:var(--panel);height:48px;flex:none}
.tb-block{display:flex;flex-direction:column;justify-content:center;padding:0 14px;border-right:1px solid var(--line)}
.tb-ey{font-family:var(--mono);font-size:9px;letter-spacing:.14em;color:var(--faint);text-transform:uppercase}
.tb-tl{font-weight:600;font-size:14px;white-space:nowrap}
.tb-tl input{background:transparent;border:none;color:var(--text);font:inherit;outline:none;width:200px}
.tb-sp{flex:1}
.tb-g{display:flex;align-items:center;gap:5px;padding:0 10px;border-left:1px solid var(--line)}

.btn{font-family:var(--mono);font-size:10px;letter-spacing:.04em;padding:5px 10px;border:1px solid var(--line2);background:var(--panel2);color:var(--text);border-radius:3px;cursor:pointer;white-space:nowrap}
.btn:hover{border-color:var(--accent);color:#fff}.btn.on{background:var(--accent);border-color:var(--accent);color:#14181c;font-weight:700}
.btn.gh{background:transparent}.btn.sm{padding:3px 7px;font-size:9px}.btn.dng:hover{border-color:var(--danger);color:var(--danger)}.btn:disabled{opacity:.35;cursor:default}

.main{flex:1;display:flex;min-height:0}
.side{width:240px;flex:none;background:var(--panel);border-right:1px solid var(--line);display:flex;flex-direction:column;min-height:0}
.insp{width:290px;flex:none;background:var(--panel);border-left:1px solid var(--line);display:flex;flex-direction:column;min-height:0}
.stabs{display:flex;border-bottom:1px solid var(--line);flex:none}
.stab{flex:1;padding:7px 4px;text-align:center;font-family:var(--mono);font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);cursor:pointer;border:none;border-bottom:2px solid transparent;background:none}
.stab.on{color:var(--text);border-bottom-color:var(--accent)}
.sbody{flex:1;overflow-y:auto;padding:10px}

.sec{font-family:var(--mono);font-size:9px;letter-spacing:.14em;text-transform:uppercase;color:var(--faint);margin:12px 0 5px;display:flex;align-items:center;gap:8px}
.sec:first-child{margin-top:0}.sec::after{content:"";flex:1;height:1px;background:var(--line)}
.hint{color:var(--muted);font-size:11px;line-height:1.5;margin:5px 0}

.ri{display:flex;align-items:center;gap:7px;padding:5px 7px;border:1px solid transparent;border-radius:3px;cursor:pointer;margin-bottom:2px}
.ri:hover{background:var(--panel2)}.ri.on{background:var(--panel2);border-color:var(--accent-dim)}
.rdot{width:8px;height:8px;border-radius:2px;flex:none}
.rtag{font-family:var(--mono);font-size:11px;font-weight:700;white-space:nowrap}
.rname{font-size:11px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1}
.pill{font-family:var(--mono);font-size:9px;padding:1px 5px;border-radius:2px;border:1px solid var(--line2);color:var(--muted);flex:none}

.cwrap{flex:1;position:relative;min-width:0;background:radial-gradient(circle at 30% 20%,#131920 0%,var(--bg) 60%)}
.csvg{position:absolute;inset:0;width:100%;height:100%;display:block}
.c3d{position:absolute;inset:0}
.sbar{position:absolute;left:0;right:0;bottom:0;height:24px;display:flex;align-items:center;background:rgba(16,20,24,.92);border-top:1px solid var(--line);font-family:var(--mono);font-size:10px;color:var(--muted);z-index:5}
.sbar>div{padding:0 10px;border-right:1px solid var(--line);height:100%;display:flex;align-items:center;gap:5px;white-space:nowrap}.sbar b{color:var(--text);font-weight:600}
.mbanner{position:absolute;top:10px;left:50%;transform:translateX(-50%);background:var(--accent);color:#14181c;font-family:var(--mono);font-size:10px;font-weight:700;padding:4px 12px;border-radius:3px;z-index:6}

.fld{margin-bottom:7px}
.fld label{display:block;font-family:var(--mono);font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--faint);margin-bottom:2px}
.fld input,.fld select,.fld textarea{width:100%;background:var(--bg);border:1px solid var(--line2);color:var(--text);border-radius:3px;padding:5px 7px;font-size:12px;font-family:var(--sans);outline:none}
.fld input:focus,.fld select:focus,.fld textarea:focus{border-color:var(--accent)}
.fg2{display:grid;grid-template-columns:1fr 1fr;gap:7px}.fg3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:7px}

.bt{width:100%;border-collapse:collapse;font-size:11px}
.bt th{font-family:var(--mono);font-size:9px;letter-spacing:.06em;text-transform:uppercase;color:var(--faint);text-align:left;padding:3px 5px;border-bottom:1px solid var(--line)}
.bt td{padding:3px 4px;border-bottom:1px solid var(--line)}
.bt input{width:100%;background:transparent;border:none;color:var(--text);font-size:11px;outline:none;padding:2px}.bt input:focus{background:var(--bg)}

.cc{display:flex;align-items:center;gap:7px;padding:5px 7px;border:1px solid var(--line);border-radius:3px;margin-bottom:3px;cursor:pointer}
.cc:hover{border-color:var(--line2)}.cc.on{border-color:var(--accent)}

.mback{position:fixed;inset:0;background:rgba(8,10,12,.7);display:flex;align-items:center;justify-content:center;z-index:50}
.mdl{width:440px;max-width:92vw;max-height:86vh;overflow-y:auto;background:var(--panel);border:1px solid var(--line2);border-radius:6px;padding:16px}
.mdl h3{margin:0 0 4px;font-size:15px}.mdl .sub{color:var(--muted);font-size:11px;margin-bottom:12px;line-height:1.5}

.toasts{position:absolute;bottom:34px;right:10px;display:flex;flex-direction:column;gap:5px;z-index:20}
.toast{background:var(--panel2);border:1px solid var(--line2);border-left:3px solid var(--accent);padding:7px 10px;border-radius:3px;font-size:11px;max-width:280px}
.toast.err{border-left-color:var(--danger)}.toast.ok{border-left-color:var(--ok)}

.empty{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;color:var(--muted);z-index:2;pointer-events:none}
.empty .big{font-family:var(--mono);font-size:12px;letter-spacing:.12em;color:var(--faint)}.empty button{pointer-events:auto}

.ws-picker{display:grid;grid-template-columns:1fr 1fr;gap:4px;margin:6px 0}
.ws-opt{display:flex;align-items:center;gap:6px;padding:5px 7px;border:1px solid var(--line);border-radius:3px;cursor:pointer;font-size:10px}
.ws-opt:hover{border-color:var(--line2)}.ws-opt.on{border-color:var(--accent);background:var(--panel2)}
.ws-dot{width:10px;height:10px;border-radius:2px;flex:none}
`;

function Fld({ label, children }) { return <div className="fld"><label>{label}</label>{children}</div>; }
function round2(v) { return Math.round((v ?? 0) * 100) / 100; }

/* ============================== APP ================================= */

export default function FacilityTwin() {
  const [project, setProject] = useState(emptyProject);
  const [selected, setSelected] = useState(null);
  const [mode, setMode] = useState("select"); // select | add-ws | add-utility | material-flow | utility-conn
  const [pendingWsType, setPendingWsType] = useState("mixer");
  const [pendingUtilType, setPendingUtilType] = useState("power");
  const [connectFrom, setConnectFrom] = useState(null);
  const [pendingMatType, setPendingMatType] = useState("dry");
  const [planMode, setPlanMode] = useState("2d");
  const [leftTab, setLeftTab] = useState("import");
  const [ghOpen, setGhOpen] = useState(false);
  const [ghBusy, setGhBusy] = useState(false);
  const [toasts, setToasts] = useState([]);
  const [cursor, setCursor] = useState({ x: 0, y: 0 });
  const [cam, setCam] = useState({ x: -2, y: -2, w: 44 });
  const [placing, setPlacing] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [aspect, setAspect] = useState(0.62);

  const svgRef = useRef(null);
  const dragRef = useRef(null);
  const threeRef = useRef(null);
  const fileDxfRef = useRef(null);
  const fileListRef = useRef(null);
  const fileJsonRef = useRef(null);

  const toast = useCallback((msg, kind = "") => {
    const id = uid(); setToasts(t => [...t, { id, msg, kind }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 4200);
  }, []);

  useEffect(() => {
    const measure = () => { const el = svgRef.current; if (el) { const r = el.getBoundingClientRect(); if (r.width > 0) setAspect(r.height / r.width); } };
    measure(); window.addEventListener("resize", measure); const t = setInterval(measure, 1500);
    return () => { window.removeEventListener("resize", measure); clearInterval(t); };
  }, [planMode]);

  useEffect(() => {
    try { const s = localStorage.getItem("ft-l1:project"); if (s) { setProject(JSON.parse(s)); toast("Restored project", "ok"); } } catch {}
    setLoaded(true);
  }, []);
  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => { try { localStorage.setItem("ft-l1:project", JSON.stringify(project)); } catch {} }, 1200);
    return () => clearTimeout(t);
  }, [project, loaded]);

  const nodes = useMemo(() => {
    const m = {};
    project.equipment.forEach(e => { m[e.id] = { ...e, kind: "equipment" }; });
    project.utilities.forEach(u => { m[u.id] = { ...u, kind: "utility" }; });
    return m;
  }, [project.equipment, project.utilities]);

  const selEquip = selected?.kind === "equipment" ? project.equipment.find(e => e.id === selected.id) : null;
  const selUtil = selected?.kind === "utility" ? project.utilities.find(u => u.id === selected.id) : null;
  const selFlow = selected?.kind === "flow" ? project.materialFlows.find(f => f.id === selected.id) : null;
  const selUtilConn = selected?.kind === "utilconn" ? project.utilityConnections.find(c => c.id === selected.id) : null;

  const fitView = useCallback((b) => {
    const bb = b || project.floorplan?.bounds; if (!bb) return;
    const pad = Math.max(2, (bb.maxX - bb.minX) * 0.08);
    setCam({ x: bb.minX - pad, y: bb.minY - pad, w: (bb.maxX - bb.minX) + pad * 2 });
  }, [project.floorplan]);

  /* ---- imports ---- */
  async function onDxfFile(file) {
    if (!file) return;
    if (/\.dwg$/i.test(file.name)) { toast("DWG must be converted to DXF first.", "err"); return; }
    try { const fp = parseDXF(await file.text()); if (!fp.count) { toast("No entities found.", "err"); return; }
      setProject(p => ({ ...p, floorplan: { ...fp, source: file.name } })); fitView(fp.bounds); toast(`Imported ${fp.count} entities`, "ok");
    } catch (e) { toast("DXF error: " + e.message, "err"); }
  }
  async function onEquipmentFile(file) {
    if (!file) return;
    try { let rows;
      if (/\.(xlsx|xls|xlsm)$/i.test(file.name)) { const wb = XLSX.read(await file.arrayBuffer()); rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }); }
      else { rows = Papa.parse(await file.text(), { header: true, skipEmptyLines: true }).data; }
      const eqs = rowsToEquipment(rows); if (!eqs.length) { toast("No rows recognized.", "err"); return; }
      setProject(p => ({ ...p, equipment: [...p.equipment, ...eqs] })); setLeftTab("equipment"); toast(`Imported ${eqs.length} workstations`, "ok");
    } catch (e) { toast("Import failed: " + e.message, "err"); }
  }

  /* ---- canvas interaction ---- */
  const clientToWorld = useCallback((cx, cy) => {
    const el = svgRef.current; if (!el) return { x: 0, y: 0 };
    const r = el.getBoundingClientRect(), h = cam.w * (r.height / r.width);
    return { x: cam.x + ((cx - r.left) / r.width) * cam.w, y: cam.y + ((cy - r.top) / r.height) * h };
  }, [cam]);

  function onWheel(e) {
    const pt = clientToWorld(e.clientX, e.clientY);
    const f = e.deltaY > 0 ? 1.12 : 1 / 1.12;
    setCam(c => { const w = Math.min(2000, Math.max(1, c.w * f)); return { x: pt.x - (pt.x - c.x) * (w / c.w), y: pt.y - (pt.y - c.y) * (w / c.w), w }; });
  }

  function onCanvasDown(e, hit) {
    const pt = clientToWorld(e.clientX, e.clientY);
    if (mode === "add-ws") {
      const ws = WS_TYPES[pendingWsType];
      const eq = { id: uid(), tag: `${ws.short}-${String(project.equipment.length + 1).padStart(2, "0")}`, name: ws.label, wsType: pendingWsType,
        manufacturer: "", model: "", power: "", area: "", x: pt.x, y: pt.y, placed: true,
        w: ws.dW, d: ws.dD, h: ws.dH, bom: [], custom: {} };
      setProject(p => ({ ...p, equipment: [...p.equipment, eq] })); setSelected({ kind: "equipment", id: eq.id }); setMode("select"); return;
    }
    if (mode === "add-utility") {
      const u = { id: uid(), tag: `${UTILITY_TYPES[pendingUtilType].short}-${String(project.utilities.length + 1).padStart(2, "0")}`, name: UTILITY_TYPES[pendingUtilType].label, type: pendingUtilType, x: pt.x, y: pt.y, capacity: "" };
      setProject(p => ({ ...p, utilities: [...p.utilities, u] })); setSelected({ kind: "utility", id: u.id }); setMode("select"); return;
    }
    if (placing) { setProject(p => ({ ...p, equipment: p.equipment.map(q => q.id === placing ? { ...q, x: pt.x, y: pt.y, placed: true } : q) })); setSelected({ kind: "equipment", id: placing }); setPlacing(null); return; }
    if (mode === "material-flow" || mode === "utility-conn") {
      if (hit) {
        if (!connectFrom) { setConnectFrom(hit.id); return; }
        if (connectFrom !== hit.id) {
          if (mode === "material-flow") {
            const f = { id: uid(), fromId: connectFrom, toId: hit.id, materialType: pendingMatType, label: "", flowRate: "", notes: "" };
            setProject(p => ({ ...p, materialFlows: [...p.materialFlows, f] }));
            toast(`Material flow: ${nodes[connectFrom]?.tag} → ${nodes[hit.id]?.tag}`, "ok");
          } else {
            const from = nodes[connectFrom], to = nodes[hit.id];
            const ut = from?.kind === "utility" ? from.type : to?.kind === "utility" ? to.type : "power";
            const c = { id: uid(), fromId: connectFrom, toId: hit.id, utilityType: ut, medium: "", size: "", notes: "" };
            setProject(p => ({ ...p, utilityConnections: [...p.utilityConnections, c] }));
            toast(`Utility: ${nodes[connectFrom]?.tag} → ${nodes[hit.id]?.tag}`, "ok");
          }
          setConnectFrom(null); setMode("select");
        }
      }
      return;
    }
    // select or drag
    if (hit) {
      setSelected({ kind: hit.kind, id: hit.id });
      if (hit.kind === "equipment" || hit.kind === "utility") {
        dragRef.current = { id: hit.id, kind: hit.kind, start: pt, orig: { x: nodes[hit.id].x, y: nodes[hit.id].y }, moved: false };
      }
    } else { setSelected(null); dragRef.current = { pan: true, start: { x: e.clientX, y: e.clientY }, cam0: cam }; }
  }

  // corner resize
  function onCornerDown(e, eqId, corner) {
    e.stopPropagation();
    const eq = project.equipment.find(q => q.id === eqId); if (!eq) return;
    const pt = clientToWorld(e.clientX, e.clientY);
    dragRef.current = { resize: true, id: eqId, corner, start: pt, origX: eq.x, origY: eq.y, origW: eq.w, origD: eq.d };
  }

  function onCanvasMove(e) {
    const pt = clientToWorld(e.clientX, e.clientY); setCursor(pt);
    const d = dragRef.current; if (!d) return;
    if (d.pan) { const el = svgRef.current; if (!el) return; const r = el.getBoundingClientRect(); setCam({ ...d.cam0, x: d.cam0.x - ((e.clientX - d.start.x) / r.width) * d.cam0.w, y: d.cam0.y - ((e.clientY - d.start.y) / r.width) * d.cam0.w }); return; }
    if (d.resize) {
      const dx = pt.x - d.start.x, dy = pt.y - d.start.y;
      let { origX, origY, origW, origD } = d;
      let x = origX, y = origY, w = origW, dd = origD;
      if (d.corner === "tl") { w = Math.max(0.3, origW - dx * 2); dd = Math.max(0.3, origD - dy * 2); }
      else if (d.corner === "tr") { w = Math.max(0.3, origW + dx * 2); dd = Math.max(0.3, origD - dy * 2); }
      else if (d.corner === "bl") { w = Math.max(0.3, origW - dx * 2); dd = Math.max(0.3, origD + dy * 2); }
      else if (d.corner === "br") { w = Math.max(0.3, origW + dx * 2); dd = Math.max(0.3, origD + dy * 2); }
      setProject(p => ({ ...p, equipment: p.equipment.map(q => q.id === d.id ? { ...q, w, d: dd } : q) }));
      return;
    }
    const dx = pt.x - d.start.x, dy = pt.y - d.start.y;
    if (Math.abs(dx) + Math.abs(dy) > 0.05) d.moved = true; if (!d.moved) return;
    if (d.kind === "equipment") setProject(p => ({ ...p, equipment: p.equipment.map(q => q.id === d.id ? { ...q, x: d.orig.x + dx, y: d.orig.y + dy } : q) }));
    else if (d.kind === "utility") setProject(p => ({ ...p, utilities: p.utilities.map(q => q.id === d.id ? { ...q, x: d.orig.x + dx, y: d.orig.y + dy } : q) }));
  }
  function onCanvasUp() { dragRef.current = null; }

  /* ---- mutations ---- */
  const patchEquip = (id, patch) => setProject(p => ({ ...p, equipment: p.equipment.map(e => e.id === id ? { ...e, ...patch } : e) }));
  const patchUtil = (id, patch) => setProject(p => ({ ...p, utilities: p.utilities.map(u => u.id === id ? { ...u, ...patch } : u) }));
  const patchFlow = (id, patch) => setProject(p => ({ ...p, materialFlows: p.materialFlows.map(f => f.id === id ? { ...f, ...patch } : f) }));
  const patchUtilConn = (id, patch) => setProject(p => ({ ...p, utilityConnections: p.utilityConnections.map(c => c.id === id ? { ...c, ...patch } : c) }));

  function removeSelected() {
    if (!selected) return;
    setProject(p => {
      if (selected.kind === "flow") return { ...p, materialFlows: p.materialFlows.filter(f => f.id !== selected.id) };
      if (selected.kind === "utilconn") return { ...p, utilityConnections: p.utilityConnections.filter(c => c.id !== selected.id) };
      const drop = { ...p,
        materialFlows: p.materialFlows.filter(f => f.fromId !== selected.id && f.toId !== selected.id),
        utilityConnections: p.utilityConnections.filter(c => c.fromId !== selected.id && c.toId !== selected.id),
      };
      if (selected.kind === "equipment") return { ...drop, equipment: p.equipment.filter(e => e.id !== selected.id) };
      return { ...drop, utilities: p.utilities.filter(u => u.id !== selected.id) };
    });
    setSelected(null);
  }

  /* ---- GitHub ---- */
  async function ghPush() {
    const { owner, repo, branch, path, token } = project.github; setGhBusy(true);
    try {
      const clean = { ...project, github: { ...project.github, token: "" } };
      const content = btoa(unescape(encodeURIComponent(JSON.stringify(clean, null, 2))));
      const base = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
      const headers = { Accept: "application/vnd.github+json", Authorization: `Bearer ${token}`, "X-GitHub-Api-Version": "2022-11-28" };
      let sha; const head = await fetch(`${base}?ref=${branch}`, { headers }); if (head.ok) sha = (await head.json()).sha;
      const res = await fetch(base, { method: "PUT", headers, body: JSON.stringify({ message: `FacilityTwin — ${project.name} rev ${project.revision}`, content, branch, ...(sha ? { sha } : {}) }) });
      toast(res.ok ? `Committed to ${owner}/${repo}` : `Failed (${res.status})`, res.ok ? "ok" : "err");
    } catch { toast("Network error — deploy the app for GitHub access.", "err"); }
    setGhBusy(false);
  }

  /* ---- 3D ---- */
  useEffect(() => {
    if (planMode !== "3d" || !threeRef.current) return;
    const el = threeRef.current, W = el.clientWidth, H = el.clientHeight;
    const scene = new THREE.Scene(); scene.background = new THREE.Color(0x101418); scene.fog = new THREE.Fog(0x101418, 60, 220);
    const camera = new THREE.PerspectiveCamera(50, W / H, 0.1, 500);
    const renderer = new THREE.WebGLRenderer({ antialias: true }); renderer.setSize(W, H); renderer.setPixelRatio(Math.min(devicePixelRatio, 2)); el.appendChild(renderer.domElement);
    scene.add(new THREE.AmbientLight(0xffffff, 0.55));
    const sun = new THREE.DirectionalLight(0xffe8c8, 0.9); sun.position.set(30, 50, 20); scene.add(sun);
    const b = project.floorplan?.bounds || { minX: 0, minY: 0, maxX: 36, maxY: 22 };
    const cx = (b.minX + b.maxX) / 2, cz = (b.minY + b.maxY) / 2, span = Math.max(b.maxX - b.minX, b.maxY - b.minY);
    const floor = new THREE.Mesh(new THREE.PlaneGeometry(span + 8, span + 8), new THREE.MeshStandardMaterial({ color: 0x171c21, roughness: 0.95 }));
    floor.rotation.x = -Math.PI / 2; floor.position.set(cx, -0.02, cz); scene.add(floor);
    scene.add(Object.assign(new THREE.GridHelper(Math.ceil(span + 8), Math.ceil(span + 8), 0x2a333c, 0x1d242b), { position: new THREE.Vector3(cx, 0, cz) }));
    if (project.floorplan) { const pts = []; for (const e of project.floorplan.entities) { if (e.kind === "line" || e.kind === "poly") { const arr = e.kind === "poly" && e.closed ? [...e.pts, e.pts[0]] : e.pts; for (let i = 0; i < arr.length - 1; i++) pts.push(new THREE.Vector3(arr[i][0], 0.02, arr[i][1]), new THREE.Vector3(arr[i + 1][0], 0.02, arr[i + 1][1])); } } if (pts.length) scene.add(new THREE.LineSegments(new THREE.BufferGeometry().setFromPoints(pts), new THREE.LineBasicMaterial({ color: 0x55616b }))); }
    for (const e of project.equipment.filter(q => q.placed)) {
      const col = new THREE.Color(WS_TYPES[e.wsType]?.color || "#7E8B95");
      const isSilo = e.wsType === "silo";
      const geo = isSilo ? new THREE.CylinderGeometry(e.w / 2, e.w / 2, e.h, 16) : new THREE.BoxGeometry(e.w, e.h, e.d);
      const mesh = new THREE.Mesh(geo, new THREE.MeshStandardMaterial({ color: col, roughness: 0.6, metalness: 0.2, transparent: true, opacity: 0.85 }));
      mesh.position.set(e.x, e.h / 2, e.y); scene.add(mesh);
      scene.add(Object.assign(new THREE.LineSegments(new THREE.EdgesGeometry(geo), new THREE.LineBasicMaterial({ color: selected?.id === e.id ? 0xf08c3a : 0x55616b })), { position: mesh.position.clone() }));
    }
    for (const u of project.utilities) { const cyl = new THREE.Mesh(new THREE.CylinderGeometry(0.6, 0.6, 1.8, 16), new THREE.MeshStandardMaterial({ color: new THREE.Color(UTILITY_TYPES[u.type]?.color || "#888"), roughness: 0.5 })); cyl.position.set(u.x, 0.9, u.y); scene.add(cyl); }
    // material flow lines
    for (const f of project.materialFlows) { const a = nodes[f.fromId], t2 = nodes[f.toId]; if (!a || !t2 || a.x === undefined || t2.x === undefined) continue;
      scene.add(new THREE.Line(new THREE.BufferGeometry().setFromPoints([new THREE.Vector3(a.x, 0.5, a.y), new THREE.Vector3(a.x, 0.5, t2.y), new THREE.Vector3(t2.x, 0.5, t2.y)]), new THREE.LineBasicMaterial({ color: new THREE.Color(MATERIAL_TYPES[f.materialType]?.color || "#888") }))); }
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
  }, [planMode, project, selected, nodes]);

  /* ============================ RENDER ============================== */
  const camH = cam.w * aspect;
  const gridStep = cam.w > 120 ? 10 : cam.w > 30 ? 5 : 1;
  const strokeW = cam.w / 900;
  const handleSize = cam.w / 120;
  const unplaced = project.equipment.filter(e => !e.placed);

  const modeBanner =
    mode === "add-ws" ? `CLICK TO PLACE ${WS_TYPES[pendingWsType]?.label.toUpperCase()}` :
    mode === "add-utility" ? `CLICK TO PLACE ${UTILITY_TYPES[pendingUtilType]?.short}` :
    mode === "material-flow" ? (connectFrom ? `FROM ${nodes[connectFrom]?.tag} — CLICK DESTINATION` : "CLICK SOURCE WORKSTATION") :
    mode === "utility-conn" ? (connectFrom ? `FROM ${nodes[connectFrom]?.tag} — CLICK TARGET` : "CLICK SOURCE") :
    placing ? "CLICK TO PLACE" : null;

  const isConnecting = mode === "material-flow" || mode === "utility-conn";

  return (
    <div className="ft">
      <style>{CSS}</style>

      {/* ===== TOOLBAR ===== */}
      <div className="tb">
        <div className="tb-block"><div className="tb-ey">FacilityTwin</div><div className="tb-tl"><input value={project.name} onChange={e => setProject(p => ({ ...p, name: e.target.value }))} /></div></div>
        <div className="tb-block" style={{ minWidth: 46 }}><div className="tb-ey">Rev</div><div className="tb-tl" style={{ fontFamily: "var(--mono)" }}><input style={{ width: 28 }} value={project.revision} onChange={e => setProject(p => ({ ...p, revision: e.target.value }))} /></div></div>

        <div className="tb-g">
          <button className={`btn ${mode === "select" ? "on" : ""}`} onClick={() => { setMode("select"); setConnectFrom(null); }}>SELECT</button>
          <button className={`btn ${mode === "add-ws" ? "on" : ""}`} onClick={() => setMode("add-ws")}>+ WORKSTATION</button>
          <button className={`btn ${mode === "add-utility" ? "on" : ""}`} onClick={() => setMode("add-utility")}>+ UTILITY</button>
        </div>
        <div className="tb-g">
          <button className={`btn ${mode === "material-flow" ? "on" : ""}`} onClick={() => { setMode("material-flow"); setConnectFrom(null); }}>⇄ MATERIAL</button>
          <button className={`btn ${mode === "utility-conn" ? "on" : ""}`} onClick={() => { setMode("utility-conn"); setConnectFrom(null); }}>⚡ UTILITY</button>
        </div>

        <div className="tb-sp" />
        <div className="tb-g">
          <button className={`btn ${planMode === "2d" ? "on" : ""}`} onClick={() => setPlanMode("2d")}>2D</button>
          <button className={`btn ${planMode === "3d" ? "on" : ""}`} onClick={() => setPlanMode("3d")}>3D</button>
          <button className="btn" onClick={() => fitView()}>FIT</button>
        </div>
        <div className="tb-g">
          <button className="btn" onClick={() => setGhOpen(true)}>{project.github.owner ? `⎇ ${project.github.owner}/${project.github.repo}` : "⎇ GITHUB"}</button>
        </div>
      </div>

      <div className="main">
        {/* ===== LEFT PANEL ===== */}
        <div className="side">
          <div className="stabs">
            {["import", "equipment", "flows"].map(t => <button key={t} className={`stab ${leftTab === t ? "on" : ""}`} onClick={() => setLeftTab(t)}>{t}</button>)}
          </div>
          <div className="sbody">
            {leftTab === "import" && (<>
              <div className="sec">Floor plan</div>
              <button className="btn" style={{ width: "100%" }} onClick={() => fileDxfRef.current.click()}>IMPORT DXF…</button>
              <input ref={fileDxfRef} type="file" accept=".dxf,.dwg" hidden onChange={e => { onDxfFile(e.target.files[0]); e.target.value = ""; }} />
              {project.floorplan && <p className="hint" style={{ fontFamily: "var(--mono)", fontSize: 10 }}>▣ {project.floorplan.source}</p>}
              <div className="sec">Equipment list</div>
              <button className="btn" style={{ width: "100%" }} onClick={() => fileListRef.current.click()}>IMPORT CSV / XLSX…</button>
              <input ref={fileListRef} type="file" accept=".csv,.xlsx,.xls,.tsv" hidden onChange={e => { onEquipmentFile(e.target.files[0]); e.target.value = ""; }} />
              <div className="sec">Project</div>
              <button className="btn gh" style={{ width: "100%", marginBottom: 5 }} onClick={() => { setProject(demoProject()); setTimeout(() => fitView(), 0); toast("Demo loaded", "ok"); }}>LOAD DEMO FACILITY</button>
              <button className="btn gh" style={{ width: "100%", marginBottom: 5 }} onClick={() => downloadFile(`${project.name.replace(/\W+/g, "-")}.json`, JSON.stringify({ ...project, github: { ...project.github, token: "" } }, null, 2))}>EXPORT JSON</button>
              <button className="btn gh" style={{ width: "100%", marginBottom: 5 }} onClick={() => fileJsonRef.current.click()}>OPEN JSON…</button>
              <input ref={fileJsonRef} type="file" accept=".json" hidden onChange={e => { if (e.target.files[0]) e.target.files[0].text().then(t => { try { setProject({ ...emptyProject(), ...JSON.parse(t) }); toast("Loaded", "ok"); } catch (err) { toast(err.message, "err"); } }); e.target.value = ""; }} />
              <button className="btn gh dng" style={{ width: "100%" }} onClick={() => { setProject(emptyProject()); setSelected(null); }}>NEW EMPTY</button>

              {mode === "add-ws" && (<>
                <div className="sec">Workstation type</div>
                <div className="ws-picker">
                  {Object.entries(WS_TYPES).map(([k, v]) => (
                    <div key={k} className={`ws-opt ${pendingWsType === k ? "on" : ""}`} onClick={() => setPendingWsType(k)}>
                      <span className="ws-dot" style={{ background: v.color }} />{v.label}
                    </div>
                  ))}
                </div>
              </>)}
              {mode === "add-utility" && (<>
                <div className="sec">Utility type</div>
                {Object.entries(UTILITY_TYPES).map(([k, v]) => (
                  <div key={k} className={`ws-opt ${pendingUtilType === k ? "on" : ""}`} onClick={() => setPendingUtilType(k)} style={{ marginBottom: 3 }}>
                    <span className="ws-dot" style={{ background: v.color, borderRadius: 6 }} />{v.label}
                  </div>
                ))}
              </>)}
              {mode === "material-flow" && (<>
                <div className="sec">Material type</div>
                {Object.entries(MATERIAL_TYPES).map(([k, v]) => (
                  <div key={k} className={`ws-opt ${pendingMatType === k ? "on" : ""}`} onClick={() => setPendingMatType(k)} style={{ marginBottom: 3 }}>
                    <span className="ws-dot" style={{ background: v.color, borderRadius: 6 }} />{v.label}
                  </div>
                ))}
              </>)}
            </>)}

            {leftTab === "equipment" && (<>
              {unplaced.length > 0 && (<><div className="sec">Staged</div>{unplaced.map(e => (
                <div key={e.id} className={`ri ${placing === e.id ? "on" : ""}`} onClick={() => setPlacing(e.id)}>
                  <span className="rdot" style={{ background: WS_TYPES[e.wsType]?.color || "var(--faint)" }} /><span className="rtag">{e.tag}</span><span className="rname">{e.name}</span>
                </div>))}</>)}
              <div className="sec">Workstations ({project.equipment.filter(e => e.placed).length})</div>
              {project.equipment.filter(e => e.placed).map(e => (
                <div key={e.id} className={`ri ${selected?.id === e.id ? "on" : ""}`} onClick={() => setSelected({ kind: "equipment", id: e.id })}>
                  <span className="rdot" style={{ background: WS_TYPES[e.wsType]?.color || "var(--faint)" }} /><span className="rtag">{e.tag}</span><span className="rname">{e.name}</span>
                  <span className="pill">{WS_TYPES[e.wsType]?.short || "OBJ"}</span>
                </div>
              ))}
            </>)}

            {leftTab === "flows" && (<>
              <div className="sec">Material flows ({project.materialFlows.length})</div>
              {project.materialFlows.map(f => (
                <div key={f.id} className={`cc ${selected?.id === f.id ? "on" : ""}`} onClick={() => setSelected({ kind: "flow", id: f.id })}>
                  <span className="rdot" style={{ background: MATERIAL_TYPES[f.materialType]?.color }} />
                  <span className="rtag" style={{ fontSize: 10 }}>{nodes[f.fromId]?.tag} → {nodes[f.toId]?.tag}</span>
                  <span className="rname">{f.label || MATERIAL_TYPES[f.materialType]?.label}</span>
                </div>
              ))}
              <div className="sec">Utilities ({project.utilities.length})</div>
              {project.utilities.map(u => (
                <div key={u.id} className={`ri ${selected?.id === u.id ? "on" : ""}`} onClick={() => setSelected({ kind: "utility", id: u.id })}>
                  <span className="rdot" style={{ background: UTILITY_TYPES[u.type]?.color, borderRadius: 6 }} /><span className="rtag">{u.tag}</span><span className="rname">{u.name}</span>
                </div>
              ))}
              <div className="sec">Utility connections ({project.utilityConnections.length})</div>
              {project.utilityConnections.map(c => (
                <div key={c.id} className={`cc ${selected?.id === c.id ? "on" : ""}`} onClick={() => setSelected({ kind: "utilconn", id: c.id })}>
                  <span className="rdot" style={{ background: UTILITY_TYPES[c.utilityType]?.color }} />
                  <span className="rtag" style={{ fontSize: 10 }}>{nodes[c.fromId]?.tag} → {nodes[c.toId]?.tag}</span>
                  <span className="rname">{c.medium || UTILITY_TYPES[c.utilityType]?.label}</span>
                </div>
              ))}
            </>)}
          </div>
        </div>

        {/* ===== CANVAS ===== */}
        <div className="cwrap">
          {modeBanner && <div className="mbanner">{modeBanner}</div>}
          {planMode === "2d" && (
            <svg ref={svgRef} className="csvg" viewBox={`${cam.x} ${cam.y} ${cam.w} ${camH}`} preserveAspectRatio="none"
              onWheel={onWheel} onMouseDown={e => onCanvasDown(e, null)} onMouseMove={onCanvasMove} onMouseUp={onCanvasUp} onMouseLeave={onCanvasUp}
              style={{ cursor: mode === "select" && !placing ? "default" : "crosshair" }}>

              {/* grid */}
              <g stroke="var(--line)" strokeWidth={strokeW * 0.5} opacity="0.4">
                {Array.from({ length: Math.ceil(cam.w / gridStep) + 2 }, (_, i) => { const x = Math.floor(cam.x / gridStep) * gridStep + i * gridStep; return <line key={"v" + i} x1={x} y1={cam.y - 5} x2={x} y2={cam.y + camH + 5} />; })}
                {Array.from({ length: Math.ceil(camH / gridStep) + 2 }, (_, i) => { const y = Math.floor(cam.y / gridStep) * gridStep + i * gridStep; return <line key={"h" + i} x1={cam.x - 5} y1={y} x2={cam.x + cam.w + 5} y2={y} />; })}
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

              {/* material flows */}
              <g fill="none">
                {project.materialFlows.map(f => {
                  const a = nodes[f.fromId], b = nodes[f.toId];
                  if (!a || !b || a.x === undefined || b.x === undefined) return null;
                  const mt = MATERIAL_TYPES[f.materialType] || MATERIAL_TYPES.custom;
                  const on = selected?.id === f.id;
                  const midY = (a.y + b.y) / 2;
                  const path = `M ${a.x} ${a.y} C ${a.x} ${midY}, ${b.x} ${midY}, ${b.x} ${b.y}`;
                  return (<g key={f.id} onMouseDown={e => { e.stopPropagation(); setSelected({ kind: "flow", id: f.id }); }} style={{ cursor: "pointer" }}>
                    <path d={path} stroke="transparent" strokeWidth={strokeW * 14} />
                    <path d={path} stroke={mt.color} strokeWidth={strokeW * (on ? 4 : 2.5)} strokeDasharray={mt.dash || "none"} opacity={on ? 1 : 0.65} />
                    <circle cx={b.x} cy={b.y} r={strokeW * 5} fill={mt.color} opacity={on ? 1 : 0.65} />
                    {f.label && <text x={(a.x + b.x) / 2} y={midY - strokeW * 6} fontSize={strokeW * 28} textAnchor="middle" fill={mt.color} fontFamily="var(--mono)" opacity={on ? 1 : 0.5}>{f.label}</text>}
                  </g>);
                })}
              </g>

              {/* utility connections */}
              <g fill="none">
                {project.utilityConnections.map(c => {
                  const a = nodes[c.fromId], b = nodes[c.toId];
                  if (!a || !b || a.x === undefined || b.x === undefined) return null;
                  const col = UTILITY_TYPES[c.utilityType]?.color || "#888";
                  const on = selected?.id === c.id;
                  return (<g key={c.id} onMouseDown={e => { e.stopPropagation(); setSelected({ kind: "utilconn", id: c.id }); }} style={{ cursor: "pointer" }}>
                    <path d={`M ${a.x} ${a.y} L ${a.x} ${b.y} L ${b.x} ${b.y}`} stroke="transparent" strokeWidth={strokeW * 12} />
                    <path d={`M ${a.x} ${a.y} L ${a.x} ${b.y} L ${b.x} ${b.y}`} stroke={col} strokeWidth={strokeW * (on ? 3 : 1.5)} strokeDasharray={`${strokeW * 6} ${strokeW * 4}`} opacity={on ? 1 : 0.5} />
                  </g>);
                })}
              </g>

              {/* utility nodes */}
              {project.utilities.map(u => {
                const col = UTILITY_TYPES[u.type]?.color || "#888";
                const on = selected?.id === u.id || connectFrom === u.id;
                return (<g key={u.id} onMouseDown={e => { e.stopPropagation(); onCanvasDown(e, { kind: "utility", id: u.id }); }} style={{ cursor: isConnecting ? "pointer" : "move" }}>
                  <circle cx={u.x} cy={u.y} r={0.9} fill="var(--panel2)" stroke={col} strokeWidth={strokeW * (on ? 4 : 2)} />
                  <text x={u.x} y={u.y + 0.15} fontSize={0.45} textAnchor="middle" fill={col} fontFamily="var(--mono)" fontWeight="700">{UTILITY_TYPES[u.type]?.short}</text>
                  <text x={u.x} y={u.y - 1.2} fontSize={0.45} textAnchor="middle" fill="var(--text)" fontFamily="var(--mono)">{u.tag}</text>
                </g>);
              })}

              {/* workstations */}
              {project.equipment.filter(e => e.placed).map(e => {
                const on = selected?.id === e.id || connectFrom === e.id;
                const ws = WS_TYPES[e.wsType] || WS_TYPES.custom;
                const fillCol = on ? `${ws.color}25` : `${ws.color}15`;
                const hw = e.w / 2, hd = e.d / 2;
                return (<g key={e.id}>
                  {/* hit area + body */}
                  <g onMouseDown={ev => { ev.stopPropagation(); onCanvasDown(ev, { kind: "equipment", id: e.id }); }} style={{ cursor: isConnecting ? "pointer" : "move" }}>
                    <rect x={e.x - hw} y={e.y - hd} width={e.w} height={e.d} rx={strokeW * 2} fill={fillCol} stroke={on ? "var(--accent)" : ws.color} strokeWidth={strokeW * (on ? 3 : 1.4)} />
                    {wsIcon(e.wsType, e.x, e.y, e.w, e.d, strokeW, on)}
                    <text x={e.x} y={e.y - hd - strokeW * 8} fontSize={strokeW * 30} textAnchor="middle" fill={on ? "var(--accent)" : "var(--text)"} fontFamily="var(--mono)" fontWeight="700">{e.tag}</text>
                    <text x={e.x} y={e.y + hd + strokeW * 18} fontSize={strokeW * 22} textAnchor="middle" fill={ws.color} fontFamily="var(--mono)" opacity={0.7}>{ws.short}</text>
                  </g>
                  {/* resize handles */}
                  {on && !isConnecting && ["tl", "tr", "bl", "br"].map(corner => {
                    const hx = corner.includes("l") ? e.x - hw : e.x + hw;
                    const hy = corner.includes("t") ? e.y - hd : e.y + hd;
                    const cur = corner === "tl" || corner === "br" ? "nwse-resize" : "nesw-resize";
                    return (<rect key={corner} x={hx - handleSize / 2} y={hy - handleSize / 2} width={handleSize} height={handleSize}
                      fill="var(--accent)" stroke="var(--bg)" strokeWidth={strokeW} rx={strokeW}
                      style={{ cursor: cur }} onMouseDown={ev => onCornerDown(ev, e.id, corner)} />);
                  })}
                </g>);
              })}
            </svg>
          )}
          {planMode === "3d" && <div ref={threeRef} className="c3d" />}
          {planMode === "2d" && !project.floorplan && project.equipment.length === 0 && (
            <div className="empty"><div className="big">NO FLOOR PLAN LOADED</div><div>Import DXF or load the demo.</div>
              <button className="btn" onClick={() => { setProject(demoProject()); setTimeout(() => fitView(), 0); }}>LOAD DEMO</button></div>
          )}
          <div className="sbar">
            <div>X <b>{cursor.x.toFixed(2)}</b> Y <b>{cursor.y.toFixed(2)}</b></div>
            <div>WORKSTATIONS <b>{project.equipment.length}</b></div><div>FLOWS <b>{project.materialFlows.length}</b></div>
            <div style={{ marginLeft: "auto", borderRight: "none" }}>{project.floorplan ? project.floorplan.source : "—"}</div>
          </div>
          <div className="toasts">{toasts.map(t => <div key={t.id} className={`toast ${t.kind}`}>{t.msg}</div>)}</div>
        </div>

        {/* ===== INSPECTOR ===== */}
        <div className="insp">
          <div className="stabs"><button className="stab on" style={{ cursor: "default" }}>{selEquip ? WS_TYPES[selEquip.wsType]?.label || "Workstation" : selUtil ? "Utility" : selFlow ? "Material flow" : selUtilConn ? "Utility conn." : "Inspector"}</button></div>
          <div className="sbody">
            {!selected && <><div className="sec">Nothing selected</div><p className="hint">Select a workstation to edit type-specific properties, BOM, and material connections. Grab corners to resize. Use ⇄ MATERIAL to route ingredient flow between stations.</p></>}

            {selEquip && <WsInspector eq={selEquip} patch={p => patchEquip(selEquip.id, p)} project={project} nodes={nodes}
              onSelectFlow={id => setSelected({ kind: "flow", id })} onRemove={removeSelected} />}

            {selUtil && (<>
              <div className="sec">Utility node</div>
              <div className="fg2">
                <Fld label="Tag"><input value={selUtil.tag} onChange={e => patchUtil(selUtil.id, { tag: e.target.value })} /></Fld>
                <Fld label="Type"><select value={selUtil.type} onChange={e => patchUtil(selUtil.id, { type: e.target.value })}>{Object.entries(UTILITY_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></Fld>
              </div>
              <Fld label="Name"><input value={selUtil.name} onChange={e => patchUtil(selUtil.id, { name: e.target.value })} /></Fld>
              <Fld label="Capacity"><input value={selUtil.capacity || ""} onChange={e => patchUtil(selUtil.id, { capacity: e.target.value })} /></Fld>
              <button className="btn dng" style={{ width: "100%", marginTop: 10 }} onClick={removeSelected}>DELETE</button>
            </>)}

            {selFlow && (<>
              <div className="sec">Material flow</div>
              <p className="hint" style={{ fontFamily: "var(--mono)" }}>{nodes[selFlow.fromId]?.tag} → {nodes[selFlow.toId]?.tag}</p>
              <Fld label="Material"><select value={selFlow.materialType} onChange={e => patchFlow(selFlow.id, { materialType: e.target.value })}>{Object.entries(MATERIAL_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></Fld>
              <Fld label="Label"><input value={selFlow.label} placeholder="e.g. Sugar to mixer" onChange={e => patchFlow(selFlow.id, { label: e.target.value })} /></Fld>
              <Fld label="Flow rate"><input value={selFlow.flowRate} placeholder="e.g. 500 kg/hr" onChange={e => patchFlow(selFlow.id, { flowRate: e.target.value })} /></Fld>
              <Fld label="Notes"><textarea rows={2} value={selFlow.notes} onChange={e => patchFlow(selFlow.id, { notes: e.target.value })} /></Fld>
              <button className="btn dng" style={{ width: "100%" }} onClick={removeSelected}>DELETE FLOW</button>
            </>)}

            {selUtilConn && (<>
              <div className="sec">Utility connection</div>
              <p className="hint" style={{ fontFamily: "var(--mono)" }}>{nodes[selUtilConn.fromId]?.tag} → {nodes[selUtilConn.toId]?.tag}</p>
              <Fld label="Service"><select value={selUtilConn.utilityType} onChange={e => patchUtilConn(selUtilConn.id, { utilityType: e.target.value })}>{Object.entries(UTILITY_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}</select></Fld>
              <Fld label="Medium"><input value={selUtilConn.medium} onChange={e => patchUtilConn(selUtilConn.id, { medium: e.target.value })} /></Fld>
              <Fld label="Size"><input value={selUtilConn.size} onChange={e => patchUtilConn(selUtilConn.id, { size: e.target.value })} /></Fld>
              <button className="btn dng" style={{ width: "100%" }} onClick={removeSelected}>DELETE</button>
            </>)}
          </div>
        </div>
      </div>

      {/* ===== GITHUB MODAL ===== */}
      {ghOpen && (
        <div className="mback" onMouseDown={e => { if (e.target === e.currentTarget) setGhOpen(false); }}>
          <div className="mdl">
            <h3>GitHub Repository</h3>
            <p className="sub">Commit project snapshots. PAT needs <b>Contents: read &amp; write</b>.</p>
            <div className="fg2">
              <Fld label="Owner"><input value={project.github.owner} onChange={e => setProject(p => ({ ...p, github: { ...p.github, owner: e.target.value.trim() } }))} /></Fld>
              <Fld label="Repo"><input value={project.github.repo} onChange={e => setProject(p => ({ ...p, github: { ...p.github, repo: e.target.value.trim() } }))} /></Fld>
            </div>
            <div className="fg2">
              <Fld label="Branch"><input value={project.github.branch} onChange={e => setProject(p => ({ ...p, github: { ...p.github, branch: e.target.value.trim() } }))} /></Fld>
              <Fld label="Path"><input value={project.github.path} onChange={e => setProject(p => ({ ...p, github: { ...p.github, path: e.target.value.trim() } }))} /></Fld>
            </div>
            <Fld label="Token"><input type="password" placeholder="github_pat_…" value={project.github.token} onChange={e => setProject(p => ({ ...p, github: { ...p.github, token: e.target.value.trim() } }))} /></Fld>
            <div style={{ display: "flex", gap: 7, marginTop: 10 }}>
              <button className="btn on" disabled={ghBusy || !project.github.owner || !project.github.repo || !project.github.token} onClick={ghPush}>{ghBusy ? "…" : "PUSH SNAPSHOT"}</button>
              <div style={{ flex: 1 }} /><button className="btn gh" onClick={() => setGhOpen(false)}>CLOSE</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ==================== WORKSTATION INSPECTOR ======================== */

function WsInspector({ eq, patch, project, nodes, onSelectFlow, onRemove }) {
  const [tab, setTab] = useState("props");
  const ws = WS_TYPES[eq.wsType] || WS_TYPES.custom;
  const setBom = bom => patch({ bom });
  const patchCustom = (key, val) => patch({ custom: { ...eq.custom, [key]: val } });
  const flows = project.materialFlows.filter(f => f.fromId === eq.id || f.toId === eq.id);
  const utilConns = project.utilityConnections.filter(c => c.fromId === eq.id || c.toId === eq.id);

  return (<>
    <div style={{ display: "flex", gap: 3, marginBottom: 8, flexWrap: "wrap" }}>
      {["props", "bom", "flows"].map(t => (
        <button key={t} className={`btn sm ${tab === t ? "on" : "gh"}`} onClick={() => setTab(t)}>
          {t === "props" ? "PROPERTIES" : t === "bom" ? `BOM (${eq.bom.length})` : `FLOWS (${flows.length})`}
        </button>
      ))}
    </div>

    {tab === "props" && (<>
      <div className="fg2">
        <Fld label="Tag"><input value={eq.tag} onChange={e => patch({ tag: e.target.value })} /></Fld>
        <Fld label="Workstation type">
          <select value={eq.wsType} onChange={e => {
            const nws = WS_TYPES[e.target.value];
            patch({ wsType: e.target.value, ...(nws ? {} : {}) });
          }}>
            {Object.entries(WS_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
          </select>
        </Fld>
      </div>
      <Fld label="Name"><input value={eq.name} onChange={e => patch({ name: e.target.value })} /></Fld>
      <div className="fg2">
        <Fld label="Manufacturer"><input value={eq.manufacturer} onChange={e => patch({ manufacturer: e.target.value })} /></Fld>
        <Fld label="Model"><input value={eq.model} onChange={e => patch({ model: e.target.value })} /></Fld>
      </div>
      <div className="fg2">
        <Fld label="Power (kW)"><input type="number" value={eq.power} onChange={e => patch({ power: e.target.value === "" ? "" : parseFloat(e.target.value) })} /></Fld>
        <Fld label="Area / room"><input value={eq.area} onChange={e => patch({ area: e.target.value })} /></Fld>
      </div>

      <div className="sec">{ws.label} properties</div>
      {ws.props.map(p => (
        <Fld key={p.key} label={p.label}>
          {p.type === "select" ? (
            <select value={eq.custom[p.key] || p.opts?.[0] || ""} onChange={e => patchCustom(p.key, e.target.value)}>
              {(p.opts || []).map(o => <option key={o} value={o}>{o}</option>)}
            </select>
          ) : (
            <input type={p.type === "number" ? "number" : "text"} value={eq.custom[p.key] || ""} onChange={e => patchCustom(p.key, e.target.value)} />
          )}
        </Fld>
      ))}

      <div className="sec">Size (m) — drag corners to resize</div>
      <div className="fg3">
        <Fld label="W"><input type="number" step="0.1" value={eq.w} onChange={e => patch({ w: parseFloat(e.target.value) || 0.3 })} /></Fld>
        <Fld label="D"><input type="number" step="0.1" value={eq.d} onChange={e => patch({ d: parseFloat(e.target.value) || 0.3 })} /></Fld>
        <Fld label="H"><input type="number" step="0.1" value={eq.h} onChange={e => patch({ h: parseFloat(e.target.value) || 0.3 })} /></Fld>
      </div>
      <div className="fg2">
        <Fld label="X"><input type="number" step="0.1" value={round2(eq.x)} onChange={e => patch({ x: parseFloat(e.target.value) || 0 })} /></Fld>
        <Fld label="Y"><input type="number" step="0.1" value={round2(eq.y)} onChange={e => patch({ y: parseFloat(e.target.value) || 0 })} /></Fld>
      </div>
      <button className="btn dng" style={{ width: "100%", marginTop: 8 }} onClick={onRemove}>DELETE WORKSTATION</button>
    </>)}

    {tab === "bom" && (<>
      <div className="sec">Bill of materials — {eq.tag}</div>
      <table className="bt">
        <thead><tr><th>Part no.</th><th>Desc</th><th style={{ width: 36 }}>Qty</th><th style={{ width: 20 }} /></tr></thead>
        <tbody>{eq.bom.map(r => (
          <tr key={r.id}>
            <td><input style={{ fontFamily: "var(--mono)" }} value={r.pn} onChange={e => setBom(eq.bom.map(x => x.id === r.id ? { ...x, pn: e.target.value } : x))} /></td>
            <td><input value={r.desc} onChange={e => setBom(eq.bom.map(x => x.id === r.id ? { ...x, desc: e.target.value } : x))} /></td>
            <td><input type="number" value={r.qty} onChange={e => setBom(eq.bom.map(x => x.id === r.id ? { ...x, qty: parseFloat(e.target.value) || 0 } : x))} /></td>
            <td><button className="btn sm gh dng" style={{ padding: "1px 4px" }} onClick={() => setBom(eq.bom.filter(x => x.id !== r.id))}>✕</button></td>
          </tr>
        ))}</tbody>
      </table>
      {!eq.bom.length && <p className="hint">No parts recorded.</p>}
      <button className="btn sm" style={{ marginTop: 6 }} onClick={() => setBom([...eq.bom, { id: uid(), pn: "", desc: "", qty: 1 }])}>+ ADD PART</button>
      {eq.bom.length > 0 && <button className="btn sm gh" style={{ marginTop: 6, marginLeft: 6 }}
        onClick={() => downloadFile(`${eq.tag}-BOM.csv`, "part_no,description,qty\n" + eq.bom.map(r => `"${r.pn}","${r.desc}",${r.qty}`).join("\n"), "text/csv")}>EXPORT CSV</button>}
    </>)}

    {tab === "flows" && (<>
      <div className="sec">Material in/out</div>
      {flows.map(f => {
        const other = nodes[f.fromId === eq.id ? f.toId : f.fromId];
        const dir = f.fromId === eq.id ? "→" : "←";
        return (<div key={f.id} className="cc" onClick={() => onSelectFlow(f.id)}>
          <span className="rdot" style={{ background: MATERIAL_TYPES[f.materialType]?.color }} />
          <span className="rtag" style={{ fontSize: 10 }}>{dir} {other?.tag}</span>
          <span className="rname">{f.label || MATERIAL_TYPES[f.materialType]?.label}{f.flowRate ? ` · ${f.flowRate}` : ""}</span>
        </div>);
      })}
      {!flows.length && <p className="hint">No material connections. Use ⇄ MATERIAL to route flow.</p>}
      <div className="sec">Utility connections</div>
      {utilConns.map(c => {
        const other = nodes[c.fromId === eq.id ? c.toId : c.fromId];
        return (<div key={c.id} className="cc">
          <span className="rdot" style={{ background: UTILITY_TYPES[c.utilityType]?.color }} />
          <span className="rtag" style={{ fontSize: 10 }}>{other?.tag}</span>
          <span className="rname">{c.medium || UTILITY_TYPES[c.utilityType]?.label}</span>
        </div>);
      })}
      {!utilConns.length && <p className="hint">Use ⚡ UTILITY to connect services.</p>}
    </>)}
  </>);
}
