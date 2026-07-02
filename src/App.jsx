import React, { useState, useMemo, useRef, useCallback } from "react";
import {
  Upload, X, FileText, Star, StarOff, AlertTriangle, Copy, Search,
  Download, Minus, Layers, EyeOff, Eye, Trash2, Save, FolderOpen,
  ChevronLeft, ChevronRight, ChevronDown, Pin, AlertOctagon,
  MessageSquare, Printer, ClipboardPaste, Target,
} from "lucide-react";

// ---------- Parsing ----------
function parseProfileText(text) {
  const lines = text.split(/\r?\n/);
  const raw = new Map();
  let buffer = "";
  let startLine = 0;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (buffer) {
      buffer += line.trim();
    } else {
      startLine = i + 1;
      buffer = line;
    }
    if (buffer.trimEnd().endsWith("\\")) {
      buffer = buffer.trimEnd().slice(0, -1);
      continue;
    }
    const full = buffer;
    buffer = "";
    const trimmed = full.trim();
    if (!trimmed || trimmed.startsWith("#") || trimmed.startsWith("*")) continue;
    const eqIdx = trimmed.indexOf("=");
    if (eqIdx === -1) continue;
    const key = trimmed.slice(0, eqIdx).trim();
    const value = trimmed.slice(eqIdx + 1).trim();
    if (!key) continue;
    if (!raw.has(key)) raw.set(key, []);
    raw.get(key).push({ value, line: startLine });
  }
  const result = new Map();
  for (const [key, occurrences] of raw) {
    result.set(key, {
      value: occurrences[occurrences.length - 1].value,
      occurrences,
      isDuplicate: occurrences.length > 1,
    });
  }
  return result;
}

function isDefaultProfile(filename) {
  return /default/i.test(filename);
}

function extractSid(params) {
  const candidates = ["SAPSYSTEMNAME", "sapsystemname", "SAPSYSTEM"];
  for (const c of candidates) {
    if (params.has(c)) return params.get(c).value;
  }
  for (const [k, v] of params) {
    if (k.toUpperCase() === "SAPSYSTEMNAME") return v.value;
  }
  return null;
}

function categoryOf(name) {
  const idx = name.indexOf("/");
  if (idx <= 0) return "other";
  return name.slice(0, idx);
}

function isNumericLike(v) {
  const t = v.trim();
  return t !== "" && !isNaN(Number(t));
}

function rowDomId(name) {
  return "row-" + name.replace(/[^a-zA-Z0-9]/g, "-");
}

function uid() {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function downloadBlob(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------- Formatted Excel-compatible report (SpreadsheetML, no external library) ----------
function xmlEscape(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
function ssCell(value, styleId) {
  const isNum = typeof value === "number";
  const type = isNum ? "Number" : "String";
  const v = isNum ? value : xmlEscape(value);
  return `<Cell${styleId ? ` ss:StyleID="${styleId}"` : ""}><Data ss:Type="${type}">${v}</Data></Cell>`;
}
function ssRow(cells) {
  return `<Row>${cells}</Row>`;
}
function ssWorksheet(name, colWidths, rowsXml) {
  const cols = (colWidths || []).map((w) => `<Column ss:Width="${w}"/>`).join("");
  return `<Worksheet ss:Name="${xmlEscape(name.slice(0, 31))}"><Table>${cols}${rowsXml}</Table>` +
    `<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/>` +
    `<SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane><Selected/></WorksheetOptions></Worksheet>`;
}

const STYLES_XML = `
<Style ss:ID="sTitle"><Font ss:Bold="1" ss:Size="14" ss:Color="#111827"/></Style>
<Style ss:ID="sMetaLabel"><Font ss:Bold="1" ss:Color="#374151"/></Style>
<Style ss:ID="sMeta"><Font ss:Color="#374151"/></Style>
<Style ss:ID="sHeader"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1E3A5F" ss:Pattern="Solid"/><Alignment ss:Vertical="Center" ss:WrapText="1"/></Style>
<Style ss:ID="sHeaderBaseline"><Font ss:Bold="1" ss:Color="#1E3A5F"/><Interior ss:Color="#FCD34D" ss:Pattern="Solid"/><Alignment ss:Vertical="Center" ss:WrapText="1"/></Style>
<Style ss:ID="sMatch"><Interior ss:Color="#D1FAE5" ss:Pattern="Solid"/><Font ss:Color="#065F46"/></Style>
<Style ss:ID="sDiff"><Interior ss:Color="#FEF3C7" ss:Pattern="Solid"/><Font ss:Color="#92400E"/></Style>
<Style ss:ID="sMissing"><Interior ss:Color="#FEE2E2" ss:Pattern="Solid"/><Font ss:Color="#991B1B"/></Style>
<Style ss:ID="sExcluded"><Interior ss:Color="#F3F4F6" ss:Pattern="Solid"/><Font ss:Color="#6B7280" ss:Italic="1"/></Style>
<Style ss:ID="sBaselineRef"><Interior ss:Color="#EFF6FF" ss:Pattern="Solid"/><Font ss:Color="#1E3A5F" ss:Bold="1"/></Style>
<Style ss:ID="sDuplicate"><Interior ss:Color="#F3E8FF" ss:Pattern="Solid"/><Font ss:Color="#6B21A8"/></Style>
<Style ss:ID="sTypeMismatch"><Interior ss:Color="#CCFBF1" ss:Pattern="Solid"/><Font ss:Color="#0F766E"/></Style>
<Style ss:ID="sAction"><Interior ss:Color="#DBEAFE" ss:Pattern="Solid"/><Font ss:Color="#1E3A8A"/><Alignment ss:WrapText="1" ss:Vertical="Top"/></Style>
<Style ss:ID="sNormal"><Font ss:Color="#111827"/></Style>
<Style ss:ID="sLabel"><Font ss:Bold="1" ss:Color="#111827"/><Interior ss:Color="#F3F4F6" ss:Pattern="Solid"/></Style>
`;

function buildReportWorkbook({ profiles, rows, counts, excludedParams, baselineId }) {
  const baselineProfile = profiles.find((p) => p.id === baselineId);
  const sheets = [];

  const summaryRows = [
    ssRow(ssCell("SAP Profile Comparison Report", "sTitle")),
    ssRow(""),
    ssRow(ssCell("Generated", "sMetaLabel") + ssCell(new Date().toLocaleString(), "sMeta")),
    ssRow(ssCell("Profiles compared", "sMetaLabel") + ssCell(profiles.map((p) => p.name).join(", "), "sMeta")),
    ssRow(ssCell("Baseline profile", "sMetaLabel") + ssCell(baselineProfile ? baselineProfile.name : "None specified", "sMeta")),
    ssRow(""),
    ssRow(ssCell("Metric", "sHeader") + ssCell("Count", "sHeader")),
    ssRow(ssCell("Total parameters (union across profiles)", "sLabel") + ssCell(counts.total, "sNormal")),
    ssRow(ssCell("Matching across all profiles", "sLabel") + ssCell(counts.matching, "sMatch")),
    ssRow(ssCell("Different values", "sLabel") + ssCell(counts.different, "sDiff")),
    ssRow(ssCell("Missing in one or more profiles", "sLabel") + ssCell(counts.missing, "sMissing")),
    ssRow(ssCell("Parameters with duplicate entries", "sLabel") + ssCell(counts.duplicates, "sDuplicate")),
    ssRow(ssCell("Possible type mismatches (numeric vs text)", "sLabel") + ssCell(counts.typeMismatches, "sTypeMismatch")),
    ssRow(ssCell("Excluded from difference analysis", "sLabel") + ssCell(counts.excluded, "sExcluded")),
  ].join("");
  sheets.push(ssWorksheet("Summary", [320, 220], summaryRows));

  const fullHeader = ssRow(
    ssCell("Parameter", "sHeader") + ssCell("Status", "sHeader") + ssCell("Excluded", "sHeader") +
    ssCell("Has Duplicates", "sHeader") + ssCell("Type Mismatch", "sHeader") +
    profiles.map((p) => ssCell(p.name + (p.id === baselineId ? " (BASELINE)" : ""), p.id === baselineId ? "sHeaderBaseline" : "sHeader")).join("")
  );
  const statusStyleMap = { matching: "sMatch", different: "sDiff", missing: "sMissing" };
  const fullDataRows = rows.map((r) => {
    const excluded = excludedParams.has(r.name);
    const rowStyle = excluded ? "sExcluded" : statusStyleMap[r.status];
    const cells = [
      ssCell(r.name, rowStyle),
      ssCell(r.status, rowStyle),
      ssCell(excluded ? "Yes" : "No", rowStyle),
      ssCell(r.hasDuplicate ? "Yes" : "No", r.hasDuplicate && !excluded ? "sDuplicate" : rowStyle),
      ssCell(r.typeMismatch ? "Yes" : "No", r.typeMismatch && !excluded ? "sTypeMismatch" : rowStyle),
    ];
    profiles.forEach((p) => {
      const cell = r.cellByProfile[p.id];
      const val = cell ? cell.value || "(empty)" : "(not set)";
      let style = rowStyle;
      if (!excluded && p.id === baselineId) style = "sBaselineRef";
      else if (!excluded && cell && cell.isDuplicate) style = "sDuplicate";
      cells.push(ssCell(val, style));
    });
    return ssRow(cells.join(""));
  });
  sheets.push(ssWorksheet("Full Comparison", [220, 90, 70, 100, 100, ...profiles.map(() => 150)], fullHeader + fullDataRows.join("")));

  const diffRows = [];
  let diffHeaderCells;
  if (baselineProfile) {
    const others = profiles.filter((p) => p.id !== baselineId);
    diffHeaderCells = ssCell("Parameter", "sHeader") + others.map((p) => ssCell(p.name + " (current)", "sHeader")).join("") +
      ssCell(`Baseline (${baselineProfile.name})`, "sHeaderBaseline") + ssCell("Recommended Action", "sHeader");
    rows.forEach((r) => {
      if (excludedParams.has(r.name)) return;
      const baseCell = r.cellByProfile[baselineId];
      const affected = others.filter((p) => {
        const st = r.baselineStatus[p.id];
        return st === "diff" || st === "missing-in-profile" || st === "missing-in-baseline";
      });
      if (affected.length === 0) return;
      const cells = [ssCell(r.name, "sLabel")];
      others.forEach((p) => {
        const cell = r.cellByProfile[p.id];
        const st = r.baselineStatus[p.id];
        cells.push(ssCell(cell ? cell.value : "(not set)", st === "diff" ? "sDiff" : st === "missing-in-profile" ? "sMissing" : "sNormal"));
      });
      cells.push(ssCell(baseCell ? baseCell.value : "(not set)", "sBaselineRef"));
      const actions = affected.map((p) => {
        const st = r.baselineStatus[p.id];
        const cell = r.cellByProfile[p.id];
        if (st === "missing-in-profile") return `Add ${r.name} = ${baseCell.value} to ${p.name}`;
        if (st === "missing-in-baseline") return `Review: ${r.name} is set to ${cell.value} in ${p.name} but absent from baseline`;
        return `Update ${p.name}: change ${r.name} from ${cell.value} to ${baseCell.value}`;
      });
      cells.push(ssCell(actions.join(" | "), "sAction"));
      diffRows.push(ssRow(cells.join("")));
    });
  } else {
    diffHeaderCells = ssCell("Parameter", "sHeader") + profiles.map((p) => ssCell(p.name, "sHeader")).join("") + ssCell("Note", "sHeader");
    rows.forEach((r) => {
      if (excludedParams.has(r.name) || r.status !== "different") return;
      const cells = [ssCell(r.name, "sLabel")];
      profiles.forEach((p) => {
        const cell = r.cellByProfile[p.id];
        cells.push(ssCell(cell ? cell.value : "(not set)", "sDiff"));
      });
      cells.push(ssCell("Values differ across profiles - no baseline was specified, no action recommended", "sAction"));
      diffRows.push(ssRow(cells.join("")));
    });
  }
  sheets.push(ssWorksheet("Differences", [220, ...profiles.map(() => 140), 260], ssRow(diffHeaderCells) + diffRows.join("")));

  const missingHeader = ssRow(ssCell("Parameter", "sHeader") + profiles.map((p) => ssCell(p.name, "sHeader")).join(""));
  const missingRows = rows
    .filter((r) => r.status === "missing" && !excludedParams.has(r.name))
    .map((r) => {
      const cells = [ssCell(r.name, "sLabel")];
      profiles.forEach((p) => {
        const cell = r.cellByProfile[p.id];
        cells.push(ssCell(cell ? `Present (${cell.value})` : "MISSING", cell ? "sMatch" : "sMissing"));
      });
      return ssRow(cells.join(""));
    });
  sheets.push(ssWorksheet("Missing Parameters", [220, ...profiles.map(() => 150)], missingHeader + missingRows.join("")));

  const dupHeader = ssRow(["Profile", "Parameter", "Occurrence #", "Value", "Line Number"].map((h) => ssCell(h, "sHeader")).join(""));
  const dupRows = [];
  profiles.forEach((p) => {
    for (const [key, data] of p.params) {
      if (data.isDuplicate) {
        data.occurrences.forEach((occ, idx) => {
          dupRows.push(ssRow(
            ssCell(p.name, "sDuplicate") + ssCell(key, "sDuplicate") + ssCell(idx + 1, "sDuplicate") +
            ssCell(occ.value, "sDuplicate") + ssCell(occ.line, "sDuplicate")
          ));
        });
      }
    }
  });
  sheets.push(ssWorksheet("Duplicates", [150, 220, 90, 200, 90], dupHeader + dupRows.join("")));

  const exclHeader = ssRow(ssCell("Parameter", "sHeader") + ssCell("Reason", "sHeader"));
  const exclRows = Array.from(excludedParams.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([name, reason]) => ssRow(ssCell(name, "sExcluded") + ssCell(reason || "", "sExcluded")));
  sheets.push(ssWorksheet("Excluded Parameters", [280, 320], exclHeader + exclRows.join("")));

  return `<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n` +
    `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" ` +
    `xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" ` +
    `xmlns:html="http://www.w3.org/TR/REC-html40">\n<Styles>${STYLES_XML}</Styles>\n${sheets.join("\n")}\n</Workbook>`;
}

// ---------- Printable / PDF report (uses browser print-to-PDF, no external library) ----------
function buildPrintableReport({ profiles, rows, counts, excludedParams, baselineId }) {
  const baselineProfile = profiles.find((p) => p.id === baselineId);
  const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  const summaryCards = [
    ["Parameters", counts.total, "#1f2530"],
    ["Matching", counts.matching, "#15803d"],
    ["Different", counts.different, "#b45309"],
    ["Missing", counts.missing, "#dc2626"],
    ["Duplicates", counts.duplicates, "#7e22ce"],
    ["Type mismatches", counts.typeMismatches, "#0d9488"],
    ["Excluded", counts.excluded, "#6b7280"],
  ].map(([label, val, color]) => `
    <div style="border:1px solid #e2e5eb;border-radius:8px;padding:8px 12px;min-width:82px;">
      <div style="font-size:18px;font-weight:700;color:${color};">${val}</div>
      <div style="font-size:10px;color:#6b7280;margin-top:2px;">${esc(label)}</div>
    </div>`).join("");

  let diffSection;
  if (baselineProfile) {
    const others = profiles.filter((p) => p.id !== baselineId);
    const diffRowsHtml = [];
    rows.forEach((r) => {
      if (excludedParams.has(r.name)) return;
      const baseCell = r.cellByProfile[baselineId];
      const affected = others.filter((p) => {
        const st = r.baselineStatus[p.id];
        return st === "diff" || st === "missing-in-profile" || st === "missing-in-baseline";
      });
      if (!affected.length) return;
      const actions = affected.map((p) => {
        const st = r.baselineStatus[p.id];
        const cell = r.cellByProfile[p.id];
        if (st === "missing-in-profile") return `Add ${r.name} = ${baseCell.value} to ${p.name}`;
        if (st === "missing-in-baseline") return `Review: ${r.name} is set to ${cell.value} in ${p.name} but absent from baseline`;
        return `Update ${p.name}: change ${r.name} from ${cell.value} to ${baseCell.value}`;
      });
      diffRowsHtml.push(`<tr>
        <td style="padding:5px 8px;font-family:monospace;font-size:10.5px;border-bottom:1px solid #eef0f4;">${esc(r.name)}</td>
        <td style="padding:5px 8px;font-family:monospace;font-size:10.5px;border-bottom:1px solid #eef0f4;background:#eff6ff;">${esc(baseCell ? baseCell.value : "(not set)")}</td>
        <td style="padding:5px 8px;font-size:10.5px;border-bottom:1px solid #eef0f4;background:#dbeafe;">${esc(actions.join(" | "))}</td>
      </tr>`);
    });
    diffSection = `
      <h2 style="font-size:13px;margin:20px 0 8px;">Differences vs baseline (${esc(baselineProfile.name)})</h2>
      <table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#1e3a5f;color:#fff;">
        <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Parameter</th>
        <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Baseline value</th>
        <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Recommended action</th>
      </tr></thead><tbody>${diffRowsHtml.join("") || `<tr><td colspan="3" style="padding:8px;color:#6b7280;font-size:10.5px;">No differences from baseline.</td></tr>`}</tbody></table>`;
  } else {
    const diffRowsHtml = rows.filter((r) => r.status === "different" && !excludedParams.has(r.name)).map((r) => {
      const vals = profiles.map((p) => {
        const c = r.cellByProfile[p.id];
        return `${esc(p.name)}: ${esc(c ? c.value : "(not set)")}`;
      }).join(" | ");
      return `<tr><td style="padding:5px 8px;font-family:monospace;font-size:10.5px;border-bottom:1px solid #eef0f4;">${esc(r.name)}</td><td style="padding:5px 8px;font-size:10.5px;border-bottom:1px solid #eef0f4;background:#fef3c7;">${vals}</td></tr>`;
    });
    diffSection = `
      <h2 style="font-size:13px;margin:20px 0 8px;">Parameters with different values</h2>
      <p style="font-size:10.5px;color:#6b7280;margin:0 0 8px;">No baseline was specified, so no actions are recommended — differences are listed for review.</p>
      <table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#1e3a5f;color:#fff;">
        <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Parameter</th>
        <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Values by profile</th>
      </tr></thead><tbody>${diffRowsHtml.join("") || `<tr><td colspan="2" style="padding:8px;color:#6b7280;font-size:10.5px;">No differences found.</td></tr>`}</tbody></table>`;
  }

  const missingRowsHtml = rows.filter((r) => r.status === "missing" && !excludedParams.has(r.name)).map((r) => {
    const vals = profiles.map((p) => {
      const c = r.cellByProfile[p.id];
      return `${esc(p.name)}: ${c ? "present (" + esc(c.value) + ")" : "MISSING"}`;
    }).join(" | ");
    return `<tr><td style="padding:5px 8px;font-family:monospace;font-size:10.5px;border-bottom:1px solid #eef0f4;">${esc(r.name)}</td><td style="padding:5px 8px;font-size:10.5px;border-bottom:1px solid #eef0f4;">${vals}</td></tr>`;
  });

  const dupRowsHtml = [];
  profiles.forEach((p) => {
    for (const [key, data] of p.params) {
      if (data.isDuplicate) {
        dupRowsHtml.push(`<tr><td style="padding:5px 8px;font-size:10.5px;border-bottom:1px solid #eef0f4;">${esc(p.name)}</td><td style="padding:5px 8px;font-family:monospace;font-size:10.5px;border-bottom:1px solid #eef0f4;">${esc(key)}</td><td style="padding:5px 8px;font-size:10.5px;border-bottom:1px solid #eef0f4;">${data.occurrences.length}</td></tr>`);
      }
    }
  });

  const exclRowsHtml = Array.from(excludedParams.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([name, reason]) =>
    `<tr><td style="padding:5px 8px;font-family:monospace;font-size:10.5px;border-bottom:1px solid #eef0f4;">${esc(name)}</td><td style="padding:5px 8px;font-size:10.5px;border-bottom:1px solid #eef0f4;color:#6b7280;">${esc(reason || "—")}</td></tr>`
  );

  return `<!DOCTYPE html><html><head><meta charset="utf-8"/><title>SAP Profile Analyzer Report</title>
<style>
  @page { margin: 16mm 14mm; }
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; box-sizing: border-box; }
  body { font-family: Arial, Helvetica, sans-serif; color: #1f2530; margin:0; padding: 0 0 30px; }
  h1 { font-size: 17px; margin: 0 0 4px; }
  h2 { break-after: avoid; }
  tr { break-inside: avoid; }
  thead { display: table-header-group; }
</style></head>
<body>
  <h1>SAP Profile Analyzer — Comparison Report</h1>
  <div style="font-size:10.5px;color:#6b7280;margin-bottom:12px;">
    Generated ${esc(new Date().toLocaleString())} &nbsp;·&nbsp; Profiles: ${esc(profiles.map((p) => p.name).join(", "))}
    ${baselineProfile ? ` &nbsp;·&nbsp; Baseline: <strong>${esc(baselineProfile.name)}</strong>` : " · No baseline specified"}
  </div>
  <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;">${summaryCards}</div>
  ${diffSection}
  <h2 style="font-size:13px;margin:20px 0 8px;">Missing parameters</h2>
  <table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#1e3a5f;color:#fff;"><th style="padding:5px 8px;text-align:left;font-size:10.5px;">Parameter</th><th style="padding:5px 8px;text-align:left;font-size:10.5px;">Presence by profile</th></tr></thead>
  <tbody>${missingRowsHtml.join("") || `<tr><td colspan="2" style="padding:8px;color:#6b7280;font-size:10.5px;">No missing parameters.</td></tr>`}</tbody></table>

  <h2 style="font-size:13px;margin:20px 0 8px;">Duplicate parameter entries</h2>
  <table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#1e3a5f;color:#fff;"><th style="padding:5px 8px;text-align:left;font-size:10.5px;">Profile</th><th style="padding:5px 8px;text-align:left;font-size:10.5px;">Parameter</th><th style="padding:5px 8px;text-align:left;font-size:10.5px;">Occurrences</th></tr></thead>
  <tbody>${dupRowsHtml.join("") || `<tr><td colspan="3" style="padding:8px;color:#6b7280;font-size:10.5px;">No duplicates found.</td></tr>`}</tbody></table>

  <h2 style="font-size:13px;margin:20px 0 8px;">Excluded from analysis</h2>
  <table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#1e3a5f;color:#fff;"><th style="padding:5px 8px;text-align:left;font-size:10.5px;">Parameter</th><th style="padding:5px 8px;text-align:left;font-size:10.5px;">Reason</th></tr></thead>
  <tbody>${exclRowsHtml.join("") || `<tr><td colspan="2" style="padding:8px;color:#6b7280;font-size:10.5px;">No parameters excluded.</td></tr>`}</tbody></table>

  <div style="margin-top:18px;font-size:9.5px;color:#9aa1b0;">For the full parameter-by-parameter table, see the Excel export.</div>
</body></html>`;
}

const VIEW_MODES = [
  { id: "all", label: "All" },
  { id: "matching", label: "Matching" },
  { id: "different", label: "Different" },
  { id: "missing", label: "Missing" },
  { id: "duplicates", label: "Duplicates" },
  { id: "typemismatch", label: "Type Mismatch" },
];

function LegendItem({ color, label }) {
  return (
    <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
      <span style={{ width: 8, height: 8, borderRadius: "50%", background: color, display: "inline-block", flexShrink: 0 }}></span>
      {label}
    </span>
  );
}

export default function SAPProfileComparator() {
  const [profiles, setProfiles] = useState([]);
  const [baselineId, setBaselineId] = useState(null);
  const [viewMode, setViewMode] = useState("all");
  const [search, setSearch] = useState("");
  const [excludedParams, setExcludedParams] = useState(() => new Map());
  const [showExcluded, setShowExcluded] = useState(true);
  const [showExclusionsPanel, setShowExclusionsPanel] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [parseWarnings, setParseWarnings] = useState([]);
  const [pasteOpen, setPasteOpen] = useState(false);
  const [pasteName, setPasteName] = useState("");
  const [pasteText, setPasteText] = useState("");
  const [groupByCategory, setGroupByCategory] = useState(true);
  const [collapsedGroups, setCollapsedGroups] = useState(() => new Set());
  const [jumpIndex, setJumpIndex] = useState(0);
  const fileInputRef = useRef(null);
  const sessionInputRef = useRef(null);

  const handleFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList);
    const warnings = [];
    const newProfiles = [];
    for (const file of files) {
      try {
        const text = await file.text();
        const params = parseProfileText(text);
        if (params.size === 0) {
          warnings.push(`${file.name}: no key = value parameters were found. Check the file format.`);
        }
        const sid = isDefaultProfile(file.name) ? extractSid(params) : null;
        newProfiles.push({ id: uid(), name: file.name, params, size: file.size, sid });
      } catch (e) {
        warnings.push(`${file.name}: could not be read (${e.message}).`);
      }
    }
    setProfiles((prev) => [...prev, ...newProfiles]);
    if (warnings.length) setParseWarnings((prev) => [...prev, ...warnings]);
  }, []);

  const onDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files?.length) handleFiles(e.dataTransfer.files);
  };

  const addPastedProfile = () => {
    if (!pasteText.trim()) return;
    const params = parseProfileText(pasteText);
    const name = pasteName.trim() || `Pasted profile ${profiles.length + 1}`;
    const sid = isDefaultProfile(name) ? extractSid(params) : null;
    setProfiles((prev) => [...prev, { id: uid(), name, params, size: pasteText.length, sid }]);
    setPasteName("");
    setPasteText("");
    setPasteOpen(false);
  };

  const removeProfile = (id) => {
    setProfiles((prev) => prev.filter((p) => p.id !== id));
    if (baselineId === id) setBaselineId(null);
  };

  const clearAll = () => {
    setProfiles([]);
    setBaselineId(null);
    setExcludedParams(new Map());
    setParseWarnings([]);
  };

  const toggleExcluded = (name) => {
    setExcludedParams((prev) => {
      const next = new Map(prev);
      if (next.has(name)) next.delete(name);
      else next.set(name, "");
      return next;
    });
  };

  const setExclusionReason = (name, reason) => {
    setExcludedParams((prev) => {
      if (!prev.has(name)) return prev;
      const next = new Map(prev);
      next.set(name, reason);
      return next;
    });
  };

  const moveProfile = (id, direction) => {
    setProfiles((prev) => {
      const idx = prev.findIndex((p) => p.id === id);
      const newIdx = idx + direction;
      if (newIdx < 0 || newIdx >= prev.length) return prev;
      const next = [...prev];
      [next[idx], next[newIdx]] = [next[newIdx], next[idx]];
      return next;
    });
  };

  const pinProfile = (id) => {
    setProfiles((prev) => {
      const idx = prev.findIndex((p) => p.id === id);
      if (idx <= 0) return prev;
      const next = [...prev];
      const [item] = next.splice(idx, 1);
      next.unshift(item);
      return next;
    });
  };

  const toggleGroup = (cat) => {
    setCollapsedGroups((prev) => {
      const next = new Set(prev);
      if (next.has(cat)) next.delete(cat);
      else next.add(cat);
      return next;
    });
  };

  const allParamNames = useMemo(() => {
    const set = new Set();
    profiles.forEach((p) => {
      for (const k of p.params.keys()) set.add(k);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [profiles]);

  const rows = useMemo(() => {
    return allParamNames.map((name) => {
      const cellByProfile = {};
      profiles.forEach((p) => {
        const entry = p.params.get(name);
        cellByProfile[p.id] = entry
          ? { value: entry.value, isDuplicate: entry.isDuplicate, occurrences: entry.occurrences }
          : null;
      });
      const presentProfiles = profiles.filter((p) => cellByProfile[p.id]);
      let status;
      if (presentProfiles.length < profiles.length) {
        status = "missing";
      } else {
        const uniqueVals = new Set(presentProfiles.map((p) => cellByProfile[p.id].value));
        status = uniqueVals.size <= 1 ? "matching" : "different";
      }
      let baselineStatus = null;
      if (baselineId) {
        const baseCell = cellByProfile[baselineId];
        baselineStatus = {};
        profiles.forEach((p) => {
          if (p.id === baselineId) {
            baselineStatus[p.id] = "baseline";
            return;
          }
          const cell = cellByProfile[p.id];
          if (!baseCell && !cell) baselineStatus[p.id] = "both-missing";
          else if (!baseCell) baselineStatus[p.id] = "missing-in-baseline";
          else if (!cell) baselineStatus[p.id] = "missing-in-profile";
          else baselineStatus[p.id] = cell.value === baseCell.value ? "match" : "diff";
        });
      }
      const hasDuplicate = presentProfiles.some((p) => cellByProfile[p.id].isDuplicate);
      const presentVals = presentProfiles.map((p) => cellByProfile[p.id].value);
      const numFlags = presentVals.map(isNumericLike);
      const typeMismatch = presentVals.length > 1 && numFlags.some(Boolean) && numFlags.some((f) => !f);
      return { name, cellByProfile, status, baselineStatus, hasDuplicate, typeMismatch };
    });
  }, [allParamNames, profiles, baselineId]);

  const filteredRows = useMemo(() => {
    return rows.filter((r) => {
      if (search && !r.name.toLowerCase().includes(search.toLowerCase())) return false;
      if (!showExcluded && excludedParams.has(r.name)) return false;
      if (viewMode === "matching") return r.status === "matching";
      if (viewMode === "different") return r.status === "different";
      if (viewMode === "missing") return r.status === "missing";
      if (viewMode === "duplicates") return r.hasDuplicate;
      if (viewMode === "typemismatch") return r.typeMismatch;
      if (viewMode === "baseline" && baselineId) {
        return Object.entries(r.baselineStatus || {}).some(
          ([pid, st]) => pid !== baselineId && (st === "diff" || st === "missing-in-profile" || st === "missing-in-baseline")
        );
      }
      return true;
    });
  }, [rows, search, showExcluded, excludedParams, viewMode, baselineId]);

  const groups = useMemo(() => {
    if (!groupByCategory) return [["", filteredRows]];
    const map = new Map();
    filteredRows.forEach((r) => {
      const cat = categoryOf(r.name);
      if (!map.has(cat)) map.set(cat, []);
      map.get(cat).push(r);
    });
    return Array.from(map.entries()).sort((a, b) => {
      if (a[0] === "other") return 1;
      if (b[0] === "other") return -1;
      return a[0].localeCompare(b[0]);
    });
  }, [filteredRows, groupByCategory]);

  const counts = useMemo(() => {
    const nonExcluded = rows.filter((r) => !excludedParams.has(r.name));
    return {
      total: rows.length,
      matching: nonExcluded.filter((r) => r.status === "matching").length,
      different: nonExcluded.filter((r) => r.status === "different").length,
      missing: nonExcluded.filter((r) => r.status === "missing").length,
      duplicates: rows.filter((r) => r.hasDuplicate).length,
      typeMismatches: nonExcluded.filter((r) => r.typeMismatch).length,
      excluded: excludedParams.size,
    };
  }, [rows, excludedParams]);

  const diffTargets = useMemo(
    () => filteredRows.filter((r) => (r.status === "different" || r.status === "missing" || r.typeMismatch) && !excludedParams.has(r.name)),
    [filteredRows, excludedParams]
  );

  const jumpToNextDifference = () => {
    if (!diffTargets.length) return;
    const target = diffTargets[jumpIndex % diffTargets.length];
    setJumpIndex((i) => (i + 1) % diffTargets.length);
    const cat = categoryOf(target.name);
    setCollapsedGroups((prev) => {
      if (!groupByCategory || !prev.has(cat)) return prev;
      const next = new Set(prev);
      next.delete(cat);
      return next;
    });
    requestAnimationFrame(() => {
      setTimeout(() => {
        const el = document.getElementById(rowDomId(target.name));
        if (el) {
          el.scrollIntoView({ behavior: "smooth", block: "center" });
          el.style.boxShadow = "inset 0 0 0 2px #2563eb";
          setTimeout(() => { el.style.boxShadow = ""; }, 1400);
        }
      }, 60);
    });
  };

  const exportReport = () => {
    if (!profiles.length) return;
    const xml = buildReportWorkbook({ profiles, rows, counts, excludedParams, baselineId });
    downloadBlob(`SAP_Profile_Comparison_${new Date().toISOString().slice(0, 10)}.xls`, xml, "application/vnd.ms-excel");
  };

  const exportPDF = () => {
    if (!profiles.length) return;
    const html = buildPrintableReport({ profiles, rows, counts, excludedParams, baselineId });
    const iframe = document.createElement("iframe");
    iframe.style.position = "fixed";
    iframe.style.right = "0";
    iframe.style.bottom = "0";
    iframe.style.width = "0";
    iframe.style.height = "0";
    iframe.style.border = "0";
    document.body.appendChild(iframe);
    const doc = iframe.contentDocument || iframe.contentWindow.document;
    doc.open();
    doc.write(html);
    doc.close();
    setTimeout(() => {
      iframe.contentWindow.focus();
      iframe.contentWindow.print();
      setTimeout(() => document.body.removeChild(iframe), 1000);
    }, 300);
  };

  const exportSession = () => {
    if (!profiles.length) return;
    const data = {
      version: 2,
      exportedAt: new Date().toISOString(),
      baselineId,
      excludedParams: Array.from(excludedParams.entries()),
      profiles: profiles.map((p) => ({
        id: p.id,
        name: p.name,
        size: p.size,
        sid: p.sid,
        params: Array.from(p.params.entries()),
      })),
    };
    downloadBlob(`SAP_Comparison_Session_${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(data, null, 2), "application/json");
  };

  const importSession = async (file) => {
    try {
      const text = await file.text();
      const data = JSON.parse(text);
      if (!data.profiles) throw new Error("File does not look like a saved comparison session.");
      const restored = data.profiles.map((p) => ({
        id: p.id || uid(),
        name: p.name,
        size: p.size,
        sid: p.sid || null,
        params: new Map(p.params),
      }));
      const rawExcl = Array.isArray(data.excludedParams) ? data.excludedParams : [];
      const exclMap = new Map(rawExcl.map((item) => (Array.isArray(item) ? item : [item, ""])));
      setProfiles(restored);
      setBaselineId(data.baselineId || null);
      setExcludedParams(exclMap);
      setParseWarnings([]);
    } catch (e) {
      setParseWarnings((prev) => [...prev, `Could not load session file: ${e.message}`]);
    }
  };

  const statusStyle = {
    matching: { text: "#15803d", bg: "rgba(21,128,61,0.08)" },
    different: { text: "#b45309", bg: "rgba(180,83,9,0.09)" },
    missing: { text: "#dc2626", bg: "rgba(220,38,38,0.08)" },
  };

  const baselineCellStyle = (st) => {
    switch (st) {
      case "baseline": return { text: "#1e3a5f", bg: "rgba(37,99,235,0.07)" };
      case "match": return { text: "#15803d", bg: "rgba(21,128,61,0.08)" };
      case "diff": return { text: "#b45309", bg: "rgba(180,83,9,0.1)" };
      case "missing-in-profile": return { text: "#dc2626", bg: "rgba(220,38,38,0.1)" };
      case "missing-in-baseline": return { text: "#7e22ce", bg: "rgba(126,34,206,0.08)" };
      case "both-missing": return { text: "#9ca3af", bg: "transparent" };
      default: return { text: "#9ca3af", bg: "transparent" };
    }
  };

  const renderRow = (r) => {
    const excluded = excludedParams.has(r.name);
    const reason = excludedParams.get(r.name);
    const gutterColor = excluded ? "#d7dce6" : { matching: "#22c55e", different: "#f59e0b", missing: "#ef4444" }[r.status];
    return (
      <tr key={r.name} id={rowDomId(r.name)} style={{ borderTop: "1px solid #eef0f4", opacity: excluded ? 0.5 : 1 }}>
        <td style={{ background: gutterColor, width: 4 }}></td>
        <td style={{ padding: "8px 12px" }}>
          <input type="checkbox" checked={excluded} onChange={() => toggleExcluded(r.name)} title="Exclude from difference analysis" />
        </td>
        <td className="mono" style={{ padding: "8px 12px", fontSize: 12.5, fontWeight: 500 }}>
          {r.name}
          {r.hasDuplicate && (
            <span title="One or more profiles have duplicate entries for this parameter" style={{ marginLeft: 6, color: "#7e22ce", display: "inline-flex", verticalAlign: "middle" }}>
              <Copy size={11} />
            </span>
          )}
          {r.typeMismatch && (
            <span title="Values differ in type (numeric vs text) across profiles — possible misconfiguration" style={{ marginLeft: 6, color: "#0d9488", display: "inline-flex", verticalAlign: "middle" }}>
              <AlertOctagon size={11} />
            </span>
          )}
          {excluded && reason && (
            <span title={reason} style={{ marginLeft: 6, color: "#9aa1b0", display: "inline-flex", verticalAlign: "middle" }}>
              <MessageSquare size={11} />
            </span>
          )}
        </td>
        {profiles.map((p) => {
          const cell = r.cellByProfile[p.id];
          let style = { text: "#1f2530", bg: "transparent" };
          if (baselineId) {
            style = baselineCellStyle(r.baselineStatus[p.id]);
          } else if (cell) {
            style = statusStyle[r.status] || style;
          } else {
            style = statusStyle.missing;
          }
          const title = cell?.occurrences?.length > 1
            ? `${cell.occurrences.length} occurrences:\n` + cell.occurrences.map((o, i) => `#${i + 1} (line ${o.line}): ${o.value}`).join("\n")
            : undefined;
          return (
            <td
              key={p.id}
              title={title}
              className="mono"
              style={{ padding: "8px 12px", fontSize: 12, color: style.text, background: style.bg, borderLeft: "1px solid #eef0f4" }}
            >
              {cell ? (
                <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  {cell.value || <em style={{ color: "#9aa1b0" }}>(empty)</em>}
                  {cell.isDuplicate && <AlertTriangle size={11} color="#7e22ce" />}
                </span>
              ) : (
                <Minus size={12} color="#c3c9d3" />
              )}
            </td>
          );
        })}
      </tr>
    );
  };

  return (
    <div
      style={{
        fontFamily: "'Inter', -apple-system, sans-serif",
        background: "#f6f7fb",
        color: "#1f2530",
        minHeight: "100%",
        padding: "28px",
        boxSizing: "border-box",
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 10px; height: 10px; }
        ::-webkit-scrollbar-track { background: #eef0f4; }
        ::-webkit-scrollbar-thumb { background: #cbd2de; border-radius: 6px; }
        .mono { font-family: 'JetBrains Mono', monospace; }
        button { font-family: inherit; cursor: pointer; }
        input, select, textarea { font-family: inherit; }
        table { border-collapse: collapse; width: 100%; }
        th, td { text-align: left; }
        .diff-gutter { width: 4px; }
      `}</style>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 34, height: 34, borderRadius: 8, background: "#eef2ff", border: "1px solid #dbe2f5", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Layers size={18} color="#2563eb" />
          </div>
          <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0, letterSpacing: "-0.01em" }}>SAP Profile Analyzer</h1>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button
            onClick={exportSession}
            disabled={!profiles.length}
            title="Save the parsed comparison (not the original files) so you can reload it later"
            style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "7px 12px", fontSize: 12.5, color: profiles.length ? "#374151" : "#b7bcc7" }}
          >
            <Save size={13} /> Save comparison
          </button>
          <button
            onClick={() => sessionInputRef.current?.click()}
            title="Load a previously saved comparison session (.json)"
            style={{ display: "flex", alignItems: "center", gap: 6, background: "#fff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "7px 12px", fontSize: 12.5, color: "#374151" }}
          >
            <FolderOpen size={13} /> Load comparison
          </button>
          <input
            ref={sessionInputRef}
            type="file"
            accept=".json"
            hidden
            onChange={(e) => { if (e.target.files?.[0]) importSession(e.target.files[0]); e.target.value = ""; }}
          />
        </div>
      </div>
      <p style={{ color: "#6b7280", fontSize: 13.5, margin: "0 0 22px 46px" }}>
        Upload any SAP profile files to compare them side by side, surface differences and duplicates, export a report for the Basis team, and save comparisons to revisit later.
      </p>

      {/* Upload zone */}
      <div
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
        onClick={() => fileInputRef.current?.click()}
        style={{
          border: `1.5px dashed ${dragOver ? "#2563eb" : "#d7dce6"}`,
          background: dragOver ? "rgba(37,99,235,0.05)" : "#ffffff",
          borderRadius: 10,
          padding: "26px 20px",
          textAlign: "center",
          cursor: "pointer",
          transition: "all 0.15s ease",
        }}
      >
        <input
          ref={fileInputRef}
          type="file"
          multiple
          hidden
          onChange={(e) => { if (e.target.files?.length) handleFiles(e.target.files); e.target.value = ""; }}
        />
        <Upload size={22} color="#2563eb" style={{ marginBottom: 8 }} />
        <div style={{ fontSize: 14, fontWeight: 500 }}>Drop SAP profile files here, or click to browse</div>
        <div style={{ fontSize: 12.5, color: "#9aa1b0", marginTop: 4 }}>
          Accepts any text-based profile (DEFAULT.PFL, instance profiles, etc.) — parsed as key = value pairs. DEFAULT profiles get their SID auto-detected.
        </div>
      </div>

      <div style={{ textAlign: "center", margin: "10px 0 18px" }}>
        <button
          onClick={(e) => { e.stopPropagation(); setPasteOpen((v) => !v); }}
          style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "none", border: "none", fontSize: 12.5, color: "#2563eb", fontWeight: 500 }}
        >
          <ClipboardPaste size={13} /> {pasteOpen ? "Cancel pasting text" : "or paste profile text instead"}
        </button>
        {pasteOpen && (
          <div style={{ background: "#ffffff", border: "1px solid #e2e5eb", borderRadius: 10, padding: 14, marginTop: 10, textAlign: "left" }}>
            <input
              value={pasteName}
              onChange={(e) => setPasteName(e.target.value)}
              placeholder="Profile name (e.g. DEFAULT.PFL, PRD_DVEBMGS00)"
              style={{ width: "100%", background: "#f9fafb", border: "1px solid #e2e5eb", borderRadius: 6, padding: "7px 10px", fontSize: 12.5, marginBottom: 8, color: "#1f2530" }}
            />
            <textarea
              value={pasteText}
              onChange={(e) => setPasteText(e.target.value)}
              placeholder={"login/min_password_lng = 8\nrdisp/wp_no_dia = 20"}
              rows={6}
              className="mono"
              style={{ width: "100%", background: "#f9fafb", border: "1px solid #e2e5eb", borderRadius: 6, padding: "8px 10px", fontSize: 12, color: "#1f2530", resize: "vertical" }}
            />
            <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
              <button
                onClick={addPastedProfile}
                style={{ background: "#2563eb", border: "none", borderRadius: 6, padding: "6px 12px", fontSize: 12.5, fontWeight: 600, color: "#fff" }}
              >
                Add profile
              </button>
              <button
                onClick={() => { setPasteOpen(false); setPasteName(""); setPasteText(""); }}
                style={{ background: "none", border: "1px solid #e2e5eb", borderRadius: 6, padding: "6px 12px", fontSize: 12.5, color: "#6b7280" }}
              >
                Cancel
              </button>
            </div>
          </div>
        )}
      </div>

      {parseWarnings.length > 0 && (
        <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderLeft: "4px solid #f59e0b", borderRadius: 8, padding: "12px 14px", marginBottom: 18 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 600, color: "#92400e", marginBottom: 6 }}>
            <AlertTriangle size={15} /> Some files need attention
          </div>
          {parseWarnings.map((w, i) => (
            <div key={i} style={{ fontSize: 12, color: "#92400e", marginLeft: 21 }}>{w}</div>
          ))}
        </div>
      )}

      {/* Profile chips */}
      {profiles.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 22 }}>
          {profiles.map((p) => {
            const dupCount = Array.from(p.params.values()).filter((v) => v.isDuplicate).length;
            const isBaseline = p.id === baselineId;
            return (
              <div
                key={p.id}
                style={{
                  display: "flex", alignItems: "center", gap: 8,
                  background: isBaseline ? "rgba(37,99,235,0.06)" : "#ffffff",
                  border: `1px solid ${isBaseline ? "#2563eb" : "#e2e5eb"}`,
                  borderRadius: 8, padding: "7px 10px", fontSize: 12.5,
                }}
              >
                <FileText size={13} color="#8993a4" />
                <span className="mono" style={{ fontWeight: 500 }}>
                  {p.name}{p.sid ? <span style={{ color: "#2563eb" }}> (SID: {p.sid})</span> : null}
                </span>
                <span style={{ color: "#9aa1b0" }}>· {p.params.size} params</span>
                {dupCount > 0 && (
                  <span style={{ display: "flex", alignItems: "center", gap: 3, color: "#7e22ce" }}>
                    <Copy size={11} /> {dupCount}
                  </span>
                )}
                <button
                  onClick={() => setBaselineId(isBaseline ? null : p.id)}
                  title={isBaseline ? "Unset as baseline" : "Set as baseline"}
                  style={{ background: "none", border: "none", padding: 2, display: "flex" }}
                >
                  {isBaseline ? <Star size={14} color="#2563eb" fill="#2563eb" /> : <StarOff size={14} color="#9aa1b0" />}
                </button>
                <button onClick={() => removeProfile(p.id)} style={{ background: "none", border: "none", padding: 2, display: "flex" }}>
                  <X size={14} color="#9aa1b0" />
                </button>
              </div>
            );
          })}
          <button
            onClick={clearAll}
            style={{ display: "flex", alignItems: "center", gap: 5, background: "#fff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "7px 10px", fontSize: 12.5, color: "#6b7280" }}
          >
            <Trash2 size={12} /> Clear all
          </button>
        </div>
      )}

      {profiles.length === 0 ? (
        <div style={{ textAlign: "center", padding: "48px 20px" }}>
          <Layers size={28} color="#c7cedb" style={{ marginBottom: 10 }} />
          <div style={{ fontSize: 14, fontWeight: 500, color: "#6b7280", marginBottom: 4 }}>No profiles loaded yet</div>
          <div style={{ fontSize: 12.5, color: "#9aa1b0" }}>Upload at least two profiles to see a comparison, paste profile text, or load a previously saved comparison.</div>
        </div>
      ) : (
        <>
          {/* Summary stats */}
          <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
            {[
              { label: "Parameters", value: counts.total, color: "#1f2530" },
              { label: "Matching", value: counts.matching, color: "#15803d" },
              { label: "Different", value: counts.different, color: "#b45309" },
              { label: "Missing", value: counts.missing, color: "#dc2626" },
              { label: "Duplicates", value: counts.duplicates, color: "#7e22ce" },
              { label: "Type mismatches", value: counts.typeMismatches, color: "#0d9488" },
              { label: "Excluded", value: counts.excluded, color: "#6b7280" },
            ].map((s) => (
              <div key={s.label} style={{ background: "#ffffff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "10px 16px", minWidth: 88 }}>
                <div className="mono" style={{ fontSize: 19, fontWeight: 600, color: s.color }}>{s.value}</div>
                <div style={{ fontSize: 11, color: "#9aa1b0", marginTop: 2 }}>{s.label}</div>
              </div>
            ))}
          </div>

          {/* Match / different / missing overview bar */}
          {(counts.matching + counts.different + counts.missing) > 0 && (
            <div style={{ marginBottom: 18 }}>
              <div style={{ display: "flex", height: 14, borderRadius: 7, overflow: "hidden", border: "1px solid #e2e5eb" }}>
                {counts.matching > 0 && (
                  <div style={{ width: `${(counts.matching / (counts.matching + counts.different + counts.missing)) * 100}%`, background: "#15803d" }} title={`Matching: ${counts.matching}`}></div>
                )}
                {counts.different > 0 && (
                  <div style={{ width: `${(counts.different / (counts.matching + counts.different + counts.missing)) * 100}%`, background: "#b45309" }} title={`Different: ${counts.different}`}></div>
                )}
                {counts.missing > 0 && (
                  <div style={{ width: `${(counts.missing / (counts.matching + counts.different + counts.missing)) * 100}%`, background: "#dc2626" }} title={`Missing: ${counts.missing}`}></div>
                )}
              </div>
              <div style={{ display: "flex", gap: 16, marginTop: 6, fontSize: 11.5, color: "#6b7280", flexWrap: "wrap" }}>
                {(() => {
                  const denom = counts.matching + counts.different + counts.missing;
                  const pct = (n) => Math.round((n / denom) * 100);
                  return (
                    <>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <span style={{ width: 8, height: 8, borderRadius: 2, background: "#15803d", display: "inline-block" }}></span>
                        Matching {counts.matching} ({pct(counts.matching)}%)
                      </span>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <span style={{ width: 8, height: 8, borderRadius: 2, background: "#b45309", display: "inline-block" }}></span>
                        Different {counts.different} ({pct(counts.different)}%)
                      </span>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <span style={{ width: 8, height: 8, borderRadius: 2, background: "#dc2626", display: "inline-block" }}></span>
                        Missing {counts.missing} ({pct(counts.missing)}%)
                      </span>
                    </>
                  );
                })()}
              </div>
            </div>
          )}

          {/* Controls */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
            <div style={{ display: "flex", flexWrap: "wrap", background: "#ffffff", border: "1px solid #e2e5eb", borderRadius: 8, padding: 3 }}>
              {VIEW_MODES.map((m) => (
                <button
                  key={m.id}
                  onClick={() => setViewMode(m.id)}
                  style={{
                    background: viewMode === m.id ? "#eef2ff" : "transparent",
                    color: viewMode === m.id ? "#2563eb" : "#6b7280",
                    border: "none", borderRadius: 6, padding: "6px 12px", fontSize: 12.5, fontWeight: 500,
                  }}
                >
                  {m.label}
                </button>
              ))}
              {baselineId && (
                <button
                  onClick={() => setViewMode("baseline")}
                  style={{
                    background: viewMode === "baseline" ? "#eef2ff" : "transparent",
                    color: viewMode === "baseline" ? "#2563eb" : "#6b7280",
                    border: "none", borderRadius: 6, padding: "6px 12px", fontSize: 12.5, fontWeight: 500,
                    display: "flex", alignItems: "center", gap: 4,
                  }}
                >
                  <Star size={11} /> vs Baseline
                </button>
              )}
            </div>

            <div style={{ position: "relative" }}>
              <Search size={13} color="#9aa1b0" style={{ position: "absolute", left: 10, top: 9 }} />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Filter parameters…"
                style={{
                  background: "#ffffff", border: "1px solid #e2e5eb", borderRadius: 8,
                  padding: "7px 10px 7px 30px", fontSize: 12.5, color: "#1f2530", width: 190, outline: "none",
                }}
              />
            </div>

            <button
              onClick={() => setShowExcluded((v) => !v)}
              style={{ display: "flex", alignItems: "center", gap: 6, background: "#ffffff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "7px 10px", fontSize: 12.5, color: "#6b7280" }}
            >
              {showExcluded ? <Eye size={13} /> : <EyeOff size={13} />}
              {showExcluded ? "Showing excluded" : "Hiding excluded"}
            </button>

            <button
              onClick={() => setShowExclusionsPanel((v) => !v)}
              style={{ display: "flex", alignItems: "center", gap: 6, background: showExclusionsPanel ? "#eef2ff" : "#ffffff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "7px 10px", fontSize: 12.5, color: showExclusionsPanel ? "#2563eb" : "#6b7280" }}
            >
              <MessageSquare size={13} /> Exclusions{excludedParams.size > 0 ? ` (${excludedParams.size})` : ""}
            </button>

            <div style={{ flex: 1 }} />

            <button
              onClick={exportReport}
              style={{ display: "flex", alignItems: "center", gap: 6, background: "#2563eb", border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, color: "#ffffff" }}
            >
              <Download size={13} /> Excel report
            </button>
            <button
              onClick={exportPDF}
              style={{ display: "flex", alignItems: "center", gap: 6, background: "#ffffff", border: "1px solid #2563eb", borderRadius: 8, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, color: "#2563eb" }}
            >
              <Printer size={13} /> PDF report
            </button>
          </div>

          {/* Exclusions panel */}
          {showExclusionsPanel && (
            <div style={{ border: "1px solid #e2e5eb", borderRadius: 10, background: "#ffffff", padding: 14, marginBottom: 14 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, marginBottom: 8, color: "#374151" }}>Excluded parameters ({excludedParams.size})</div>
              {excludedParams.size === 0 ? (
                <div style={{ fontSize: 12, color: "#9aa1b0" }}>No parameters excluded yet. Check the "Excl." box next to a parameter row to exclude it, then note why here.</div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 220, overflowY: "auto" }}>
                  {Array.from(excludedParams.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([name, reason]) => (
                    <div key={name} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span className="mono" style={{ fontSize: 12, minWidth: 220, color: "#374151" }}>{name}</span>
                      <input
                        value={reason}
                        onChange={(e) => setExclusionReason(name, e.target.value)}
                        placeholder="Reason (optional) — e.g. expected to differ per system"
                        style={{ flex: 1, background: "#f9fafb", border: "1px solid #e2e5eb", borderRadius: 6, padding: "5px 8px", fontSize: 12, color: "#1f2530" }}
                      />
                      <button onClick={() => toggleExcluded(name)} style={{ background: "none", border: "none", display: "flex" }} title="Remove exclusion">
                        <X size={13} color="#9aa1b0" />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {baselineId && (
            <div style={{ fontSize: 12, color: "#2563eb", marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}>
              <Star size={12} fill="#2563eb" /> Baseline: <span className="mono">{profiles.find((p) => p.id === baselineId)?.name}</span> — other profiles are evaluated against it, and the report includes recommended actions.
            </div>
          )}

          {/* Legend + table toolbar */}
          <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 10 }}>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 14, fontSize: 11.5, color: "#6b7280" }}>
              <LegendItem color="#22c55e" label="Matching" />
              <LegendItem color="#f59e0b" label="Different" />
              <LegendItem color="#ef4444" label="Missing" />
              <LegendItem color="#7e22ce" label="Duplicate entry" />
              <LegendItem color="#0d9488" label="Type mismatch" />
              <LegendItem color="#2563eb" label="Baseline column" />
              <LegendItem color="#9ca3af" label="Excluded" />
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button
                onClick={() => setGroupByCategory((v) => !v)}
                style={{ display: "flex", alignItems: "center", gap: 5, background: "#ffffff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "6px 10px", fontSize: 12, color: "#6b7280" }}
              >
                <Layers size={12} /> {groupByCategory ? "Grouped by category" : "Flat list"}
              </button>
              {groupByCategory && (
                <>
                  <button onClick={() => setCollapsedGroups(new Set())} style={{ background: "#ffffff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "6px 10px", fontSize: 12, color: "#6b7280" }}>
                    Expand all
                  </button>
                  <button onClick={() => setCollapsedGroups(new Set(groups.map(([cat]) => cat)))} style={{ background: "#ffffff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "6px 10px", fontSize: 12, color: "#6b7280" }}>
                    Collapse all
                  </button>
                </>
              )}
              <button
                onClick={jumpToNextDifference}
                disabled={!diffTargets.length}
                style={{ display: "flex", alignItems: "center", gap: 5, background: "#ffffff", border: "1px solid #e2e5eb", borderRadius: 8, padding: "6px 10px", fontSize: 12, color: diffTargets.length ? "#2563eb" : "#c3c9d3" }}
              >
                <Target size={12} /> Jump to next difference{diffTargets.length ? ` (${diffTargets.length})` : ""}
              </button>
            </div>
          </div>

          {/* Table */}
          <div style={{ border: "1px solid #e2e5eb", borderRadius: 10, overflow: "hidden", background: "#ffffff" }}>
            <div style={{ overflowX: "auto", maxHeight: 560, overflowY: "auto" }}>
              <table>
                <thead>
                  <tr style={{ background: "#f5f6f9", position: "sticky", top: 0, zIndex: 2 }}>
                    <th className="diff-gutter"></th>
                    <th style={{ padding: "10px 12px", fontSize: 11.5, color: "#6b7280", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.03em", minWidth: 60 }}>Excl.</th>
                    <th style={{ padding: "10px 12px", fontSize: 11.5, color: "#6b7280", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.03em", minWidth: 220 }}>Parameter</th>
                    {profiles.map((p) => (
                      <th key={p.id} style={{ padding: "8px 12px", fontSize: 11.5, color: p.id === baselineId ? "#2563eb" : "#6b7280", fontWeight: 600, minWidth: 160, borderLeft: "1px solid #e9ebf0" }}>
                        <div className="mono" style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                          <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                            {p.id === baselineId && <Star size={11} fill="#2563eb" />}
                            {p.name}
                          </span>
                          {p.sid && <span style={{ fontSize: 10, fontWeight: 600, color: "#2563eb" }}>SID: {p.sid}</span>}
                          <span style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 2 }}>
                            <button onClick={() => moveProfile(p.id, -1)} title="Move column left" style={{ background: "none", border: "none", padding: 0, display: "flex" }}>
                              <ChevronLeft size={12} color="#b7bcc7" />
                            </button>
                            <button onClick={() => pinProfile(p.id)} title="Pin to front" style={{ background: "none", border: "none", padding: 0, display: "flex" }}>
                              <Pin size={11} color="#b7bcc7" />
                            </button>
                            <button onClick={() => moveProfile(p.id, 1)} title="Move column right" style={{ background: "none", border: "none", padding: 0, display: "flex" }}>
                              <ChevronRight size={12} color="#b7bcc7" />
                            </button>
                          </span>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {groups.map(([cat, catRows]) => (
                    <React.Fragment key={cat || "flat"}>
                      {groupByCategory && (
                        <tr onClick={() => toggleGroup(cat)} style={{ background: "#f9fafb", cursor: "pointer", borderTop: "1px solid #e2e5eb" }}>
                          <td colSpan={profiles.length + 3} style={{ padding: "7px 12px", fontSize: 12, fontWeight: 600, color: "#374151" }}>
                            <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              <ChevronDown size={13} style={{ transform: collapsedGroups.has(cat) ? "rotate(-90deg)" : "none", transition: "transform 0.1s" }} />
                              {cat === "other" ? "Other parameters" : cat}
                              <span style={{ fontWeight: 400, color: "#9aa1b0" }}>· {catRows.length} parameter{catRows.length === 1 ? "" : "s"}</span>
                            </span>
                          </td>
                        </tr>
                      )}
                      {!collapsedGroups.has(cat) && catRows.map((r) => renderRow(r))}
                    </React.Fragment>
                  ))}
                  {filteredRows.length === 0 && (
                    <tr>
                      <td colSpan={profiles.length + 3} style={{ padding: 24, textAlign: "center", color: "#9aa1b0", fontSize: 13 }}>
                        No parameters match this view.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
          <div style={{ fontSize: 11.5, color: "#9aa1b0", marginTop: 8 }}>
            Showing {filteredRows.length} of {rows.length} parameters. Data stays in this browser tab only — export a report or save a comparison to keep it.
          </div>
        </>
      )}
    </div>
  );
}
