import React, { useState, useMemo, useRef, useCallback, useEffect } from "react";
import {
  Upload, X, FileText, Star, StarOff, AlertTriangle, Copy, Search,
  Download, Minus, Layers, EyeOff, Eye, Trash2, Save, FolderOpen,
  ChevronLeft, ChevronRight, ChevronDown, Pin, AlertOctagon,
  MessageSquare, Printer, ClipboardPaste, Target,
  Wrench, CheckCircle2, ClipboardCheck, ListTodo, Fingerprint, Sun, Moon, Contrast,
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

// Detects a profile's SAP SID from its content (SAPSYSTEMNAME) or, failing that, from the
// conventional "<SID>_<InstanceName><Nr>_<hostname>" instance profile filename pattern.
// SAP SIDs are always exactly 3 characters; anything wildly longer than that is rejected
// rather than trusted, since it later gets used to build a RegExp for SID normalization.
function detectSid(filename, params) {
  const fromParam = extractSid(params);
  if (fromParam && fromParam.length <= 12) return fromParam.toUpperCase();
  if (isDefaultProfile(filename)) return null;
  const base = filename.replace(/\.[^.]+$/, "");
  const match = base.match(/^([A-Za-z0-9]{2,8})_/);
  return match ? match[1].toUpperCase() : null;
}

function normalizeForSid(value, sid) {
  if (!sid || sid.length > 12) return value;
  const escaped = sid.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return value.replace(new RegExp(escaped, "gi"), "{SID}");
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

// SAP profiles are plain text and normally well under 1MB; anything drastically larger is
// almost certainly the wrong file selected by mistake, so it's rejected before parsing rather
// than risking the tab hanging on a huge input.
const MAX_PROFILE_FILE_SIZE = 5 * 1024 * 1024; // 5MB
const MAX_SESSION_FILE_SIZE = 25 * 1024 * 1024; // 25MB — a session bundles multiple profiles + proposals

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
    ssRow(ssCell("Matched after ignoring SID differences", "sLabel") + ssCell(counts.sidNormalized, "sBaselineRef")),
    ssRow(ssCell("Excluded from difference analysis", "sLabel") + ssCell(counts.excluded, "sExcluded")),
  ].join("");
  sheets.push(ssWorksheet("Summary", [320, 220], summaryRows));

  const fullHeader = ssRow(
    ssCell("Parameter", "sHeader") + ssCell("Status", "sHeader") + ssCell("Excluded", "sHeader") +
    ssCell("Has Duplicates", "sHeader") + ssCell("Type Mismatch", "sHeader") + ssCell("SID Normalized", "sHeader") +
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
      ssCell(r.sidNormalized ? "Yes" : "No", r.sidNormalized && !excluded ? "sBaselineRef" : rowStyle),
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
  sheets.push(ssWorksheet("Full Comparison", [220, 90, 70, 100, 100, 100, ...profiles.map(() => 150)], fullHeader + fullDataRows.join("")));

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
    ["SID-normalized", counts.sidNormalized, "#2563eb"],
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

// ---------- Remediation plan report (proposed fixes, per profile) ----------
function buildRemediationReport({ profiles, rows, proposals, excludedParams, baselineId }) {
  const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const baselineProfile = profiles.find((p) => p.id === baselineId);

  const actionable = rows.filter((r) => (r.status === "different" || r.status === "missing") && !excludedParams.has(r.name));
  const resolvedRows = actionable.filter((r) => proposals.get(r.name)?.resolved);
  const pendingRows = actionable.filter((r) => !proposals.get(r.name)?.resolved);

  let totalActions = 0;
  const byProfile = new Map(profiles.map((p) => [p.id, []]));
  const noActionRows = [];
  resolvedRows.forEach((r) => {
    const proposal = proposals.get(r.name);
    let hasAction = false;
    profiles.forEach((p) => {
      const entry = proposal.perProfile.get(p.id);
      if (!entry || entry.action === "keep") return;
      hasAction = true;
      totalActions++;
      const cell = r.cellByProfile[p.id];
      byProfile.get(p.id).push({
        name: r.name,
        current: cell ? cell.value : "(not set)",
        action: entry.action,
        value: entry.value,
        note: proposal.note,
      });
    });
    if (!hasAction) noActionRows.push({ name: r.name, note: proposal.note });
  });

  const summaryCards = [
    ["Parameters with differences", actionable.length, "#1f2530"],
    ["Decided", resolvedRows.length, "#15803d"],
    ["Still pending", pendingRows.length, "#dc2626"],
    ["Total actions to apply", totalActions, "#2563eb"],
  ].map(([label, val, color]) => `
    <div style="border:1px solid #e2e5eb;border-radius:8px;padding:8px 12px;min-width:100px;">
      <div style="font-size:18px;font-weight:700;color:${color};">${val}</div>
      <div style="font-size:10px;color:#6b7280;margin-top:2px;">${esc(label)}</div>
    </div>`).join("");

  const profileSections = profiles.map((p) => {
    const actions = byProfile.get(p.id);
    if (!actions.length) return "";
    const rowsHtml = actions.map((a) => {
      const actionLabel = a.action === "remove" ? "Remove parameter" : `Set to: ${esc(a.value || "(empty)")}`;
      const actionColor = a.action === "remove" ? "#fee2e2" : "#dbeafe";
      const actionText = a.action === "remove" ? "#991b1b" : "#1e3a8a";
      return `<tr>
        <td style="padding:5px 8px;font-family:monospace;font-size:10.5px;border-bottom:1px solid #eef0f4;">${esc(a.name)}</td>
        <td style="padding:5px 8px;font-family:monospace;font-size:10.5px;border-bottom:1px solid #eef0f4;color:#6b7280;">${esc(a.current)}</td>
        <td style="padding:5px 8px;font-size:10.5px;border-bottom:1px solid #eef0f4;background:${actionColor};color:${actionText};font-weight:500;">${actionLabel}</td>
        <td style="padding:5px 8px;font-size:10.5px;border-bottom:1px solid #eef0f4;color:#6b7280;">${esc(a.note || "—")}</td>
      </tr>`;
    }).join("");
    return `
      <h2 style="font-size:13px;margin:20px 0 8px;">${esc(p.name)}${p.sid ? " (SID: " + esc(p.sid) + ")" : ""}${p.id === baselineId ? " — baseline" : ""}</h2>
      <table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#1e3a5f;color:#fff;">
        <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Parameter</th>
        <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Current value</th>
        <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Proposed action</th>
        <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Note</th>
      </tr></thead><tbody>${rowsHtml}</tbody></table>`;

  }).join("");

  const pendingHtml = pendingRows.length
    ? `<ul style="columns:2;font-family:monospace;font-size:10.5px;color:#6b7280;margin:0;padding-left:16px;">${pendingRows.map((r) => `<li>${esc(r.name)}</li>`).join("")}</ul>`
    : `<p style="font-size:10.5px;color:#6b7280;">Nothing pending — every difference has a decision.</p>`;

  return `<!DOCTYPE html><html><head><meta charset="utf-8"/><title>SAP Remediation Plan</title>
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
  <h1>SAP Profile Analyzer — Remediation plan</h1>
  <div style="font-size:10.5px;color:#6b7280;margin-bottom:12px;">
    Generated ${esc(new Date().toLocaleString())} &nbsp;·&nbsp; Profiles: ${esc(profiles.map((p) => p.name).join(", "))}
    ${baselineProfile ? ` &nbsp;·&nbsp; Baseline: <strong>${esc(baselineProfile.name)}</strong>` : ""}
  </div>
  <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px;">${summaryCards}</div>
  ${resolvedRows.length === 0 ? `<p style="font-size:11px;color:#6b7280;">No fixes have been proposed yet — open a parameter row, click "Propose fix", pick an action for at least one profile (or leave everything as "Keep as is" if no change is needed), then check "Mark as decided".</p>` : profileSections}
  ${noActionRows.length ? `
  <h2 style="font-size:13px;margin:20px 0 8px;">Reviewed — no action needed</h2>
  <table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#1e3a5f;color:#fff;">
    <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Parameter</th>
    <th style="padding:5px 8px;text-align:left;font-size:10.5px;">Note</th>
  </tr></thead><tbody>${noActionRows.map((r) => `<tr><td style="padding:5px 8px;font-family:monospace;font-size:10.5px;border-bottom:1px solid #eef0f4;">${esc(r.name)}</td><td style="padding:5px 8px;font-size:10.5px;border-bottom:1px solid #eef0f4;color:#6b7280;">${esc(r.note || "—")}</td></tr>`).join("")}</tbody></table>` : ""}
  <h2 style="font-size:13px;margin:20px 0 8px;">Still pending a decision</h2>
  ${pendingHtml}
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

// ---------- Branding — edit these two lines to put your own name on the app ----------
const BRAND_NAME = "Shitij Bagga";
const BRAND_EMAIL = "shitij.bagga@hotmail.com";
const BRAND_LINKEDIN = "http://www.linkedin.com/in/shitijbagga";

// ---------- Themes ----------
// Light and Dark share the same status hues (green/amber/red); Colorblind-safe swaps them
// for an Okabe-Ito-inspired blue/orange/purple palette so matching/different/missing stay
// distinguishable for red-green color blindness. Report exports (Excel/PDF) are intentionally
// NOT themed — they stay a stable, printable design regardless of the on-screen theme.
const THEMES = {
  light: {
    bg: "#f6f7fb", panelBg: "#ffffff", panelBgAlt: "#f5f6f9",
    border: "#e2e5eb", borderLight: "#eef0f4",
    textPrimary: "#1f2530", textSecondary: "#374151", textMuted: "#6b7280", textFaint: "#9aa1b0", textFainter: "#c3c9d3",
    accent: "#2563eb", accentRgb: "37,99,235", accentBg: "#eef2ff", accentBorder: "#dbe2f5",
    scrollTrack: "#eef0f4", scrollThumb: "#cbd2de",
    matchText: "#15803d", matchGutter: "#22c55e", matchRgb: "21,128,61",
    diffText: "#b45309", diffGutter: "#f59e0b", diffRgb: "180,83,9",
    missingText: "#dc2626", missingGutter: "#ef4444", missingRgb: "220,38,38",
    duplicateText: "#7e22ce", typeMismatchText: "#0d9488", excludedText: "#6b7280",
    warningBg: "#fffbeb", warningBorder: "#fde68a", warningText: "#92400e",
  },
  dark: {
    bg: "#12161d", panelBg: "#1a1f29", panelBgAlt: "#20262f",
    border: "#2b3240", borderLight: "#252b36",
    textPrimary: "#e7ebf1", textSecondary: "#c3cad6", textMuted: "#98a2b3", textFaint: "#6b7688", textFainter: "#4b5563",
    accent: "#5b9df4", accentRgb: "91,157,244", accentBg: "rgba(91,157,244,0.12)", accentBorder: "rgba(91,157,244,0.35)",
    scrollTrack: "#20262f", scrollThumb: "#3a4250",
    matchText: "#4ade80", matchGutter: "#22c55e", matchRgb: "74,222,128",
    diffText: "#fbbf24", diffGutter: "#f59e0b", diffRgb: "251,191,36",
    missingText: "#f87171", missingGutter: "#ef4444", missingRgb: "248,113,113",
    duplicateText: "#c084fc", typeMismatchText: "#2dd4bf", excludedText: "#8993a4",
    warningBg: "rgba(245,158,11,0.1)", warningBorder: "rgba(245,158,11,0.35)", warningText: "#fbbf24",
  },
  colorblind: {
    bg: "#f6f7fb", panelBg: "#ffffff", panelBgAlt: "#f5f6f9",
    border: "#e2e5eb", borderLight: "#eef0f4",
    textPrimary: "#1f2530", textSecondary: "#374151", textMuted: "#6b7280", textFaint: "#9aa1b0", textFainter: "#c3c9d3",
    accent: "#0072B2", accentRgb: "0,114,178", accentBg: "#e6f0f7", accentBorder: "#b3d1e6",
    scrollTrack: "#eef0f4", scrollThumb: "#cbd2de",
    matchText: "#0072B2", matchGutter: "#0072B2", matchRgb: "0,114,178",
    diffText: "#b8730a", diffGutter: "#E69F00", diffRgb: "230,159,0",
    missingText: "#a5477a", missingGutter: "#CC79A7", missingRgb: "204,121,167",
    duplicateText: "#8a7600", typeMismatchText: "#009E73", excludedText: "#6b7280",
    warningBg: "#fff7e6", warningBorder: "#f0cf7a", warningText: "#8a5a00",
  },
};

function getInitialTheme() {
  try {
    const saved = window.localStorage.getItem("sap-analyzer-theme");
    if (saved && THEMES[saved]) return saved;
    if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) return "dark";
  } catch (e) {
    // localStorage/matchMedia unavailable — fall back to light
  }
  return "light";
}

export default function SAPProfileComparator() {
  const [profiles, setProfiles] = useState([]);
  const [baselineId, setBaselineId] = useState(null);
  const [viewMode, setViewMode] = useState("all");
  const [search, setSearch] = useState("");
  const [excludedParams, setExcludedParams] = useState(() => new Map());
  const [showExcluded, setShowExcluded] = useState(true);
  const [ignoreSidDiffs, setIgnoreSidDiffs] = useState(false);
  const [showExclusionsPanel, setShowExclusionsPanel] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [parseWarnings, setParseWarnings] = useState([]);
  const [pasteOpen, setPasteOpen] = useState(false);
  const [pasteName, setPasteName] = useState("");
  const [pasteText, setPasteText] = useState("");
  const [groupByCategory, setGroupByCategory] = useState(true);
  const [collapsedGroups, setCollapsedGroups] = useState(() => new Set());
  const [jumpIndex, setJumpIndex] = useState(0);
  const [pendingJumpIndex, setPendingJumpIndex] = useState(0);
  const [proposals, setProposals] = useState(() => new Map());
  const [openProposalRows, setOpenProposalRows] = useState(() => new Set());
  const [theme, setTheme] = useState(getInitialTheme);
  const fileInputRef = useRef(null);
  const sessionInputRef = useRef(null);

  useEffect(() => {
    try {
      window.localStorage.setItem("sap-analyzer-theme", theme);
    } catch (e) {
      // localStorage unavailable (e.g. private browsing) — theme just won't persist
    }
  }, [theme]);

  const t = THEMES[theme];

  const handleFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList);
    const warnings = [];
    const newProfiles = [];
    for (const file of files) {
      try {
        if (file.size > MAX_PROFILE_FILE_SIZE) {
          warnings.push(`${file.name}: file is ${(file.size / (1024 * 1024)).toFixed(1)}MB, which is larger than the ${MAX_PROFILE_FILE_SIZE / (1024 * 1024)}MB limit for a profile file. Skipped — check you selected the right file.`);
          continue;
        }
        const text = await file.text();
        const params = parseProfileText(text);
        if (params.size === 0) {
          warnings.push(`${file.name}: no key = value parameters were found. Check the file format.`);
        }
        const sid = detectSid(file.name, params);
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
    const sid = detectSid(name, params);
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
    setProposals(new Map());
    setOpenProposalRows(new Set());
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

  const toggleProposalRow = (name) => {
    setOpenProposalRows((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  };

  const emptyProposal = () => ({ resolved: false, note: "", perProfile: new Map(profiles.map((p) => [p.id, { action: "keep", value: "" }])) });

  const updateProposal = (name, updater) => {
    setProposals((prev) => {
      const next = new Map(prev);
      const existing = next.get(name) || emptyProposal();
      next.set(name, updater(existing));
      return next;
    });
  };

  const setProposalAction = (name, profileId, action) => updateProposal(name, (p) => {
    const perProfile = new Map(p.perProfile);
    const cur = perProfile.get(profileId) || { action: "keep", value: "" };
    perProfile.set(profileId, { action, value: action === "set" ? cur.value : "" });
    return { ...p, perProfile };
  });

  const setProposalValue = (name, profileId, value) => updateProposal(name, (p) => {
    const perProfile = new Map(p.perProfile);
    const cur = perProfile.get(profileId) || { action: "set", value: "" };
    perProfile.set(profileId, { action: "set", value });
    return { ...p, perProfile };
  });

  const applyBaselineToProposal = (name, profileId, baselineValue) => updateProposal(name, (p) => {
    const perProfile = new Map(p.perProfile);
    perProfile.set(profileId, { action: "set", value: baselineValue });
    return { ...p, perProfile };
  });

  const setProposalNote = (name, note) => updateProposal(name, (p) => ({ ...p, note }));

  const toggleProposalResolved = (name) => updateProposal(name, (p) => ({ ...p, resolved: !p.resolved }));

  const clearProposal = (name) => {
    setProposals((prev) => {
      if (!prev.has(name)) return prev;
      const next = new Map(prev);
      next.delete(name);
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
      let sidNormalized = false;
      if (presentProfiles.length < profiles.length) {
        status = "missing";
      } else {
        const uniqueVals = new Set(presentProfiles.map((p) => cellByProfile[p.id].value));
        if (uniqueVals.size <= 1) {
          status = "matching";
        } else if (ignoreSidDiffs) {
          const normVals = new Set(presentProfiles.map((p) => normalizeForSid(cellByProfile[p.id].value, p.sid)));
          if (normVals.size <= 1) {
            status = "matching";
            sidNormalized = true;
          } else {
            status = "different";
          }
        } else {
          status = "different";
        }
      }
      let baselineStatus = null;
      if (baselineId) {
        const baseCell = cellByProfile[baselineId];
        const baseProfile = profiles.find((p) => p.id === baselineId);
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
          else if (cell.value === baseCell.value) baselineStatus[p.id] = "match";
          else if (ignoreSidDiffs && normalizeForSid(cell.value, p.sid) === normalizeForSid(baseCell.value, baseProfile?.sid)) baselineStatus[p.id] = "match";
          else baselineStatus[p.id] = "diff";
        });
      }
      const hasDuplicate = presentProfiles.some((p) => cellByProfile[p.id].isDuplicate);
      const presentVals = presentProfiles.map((p) => cellByProfile[p.id].value);
      const numFlags = presentVals.map(isNumericLike);
      const typeMismatch = presentVals.length > 1 && numFlags.some(Boolean) && numFlags.some((f) => !f);
      return { name, cellByProfile, status, baselineStatus, hasDuplicate, typeMismatch, sidNormalized };
    });
  }, [allParamNames, profiles, baselineId, ignoreSidDiffs]);

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
    const actionable = nonExcluded.filter((r) => r.status === "different" || r.status === "missing");
    const decided = actionable.filter((r) => proposals.get(r.name)?.resolved).length;
    return {
      total: rows.length,
      matching: nonExcluded.filter((r) => r.status === "matching").length,
      different: nonExcluded.filter((r) => r.status === "different").length,
      missing: nonExcluded.filter((r) => r.status === "missing").length,
      duplicates: rows.filter((r) => r.hasDuplicate).length,
      typeMismatches: nonExcluded.filter((r) => r.typeMismatch).length,
      excluded: excludedParams.size,
      actionable: actionable.length,
      decided,
      pendingDecisions: actionable.length - decided,
      sidNormalized: nonExcluded.filter((r) => r.sidNormalized).length,
    };
  }, [rows, excludedParams, proposals]);

  const diffTargets = useMemo(
    () => filteredRows.filter((r) => (r.status === "different" || r.status === "missing" || r.typeMismatch) && !excludedParams.has(r.name)),
    [filteredRows, excludedParams]
  );

  const pendingTargets = useMemo(
    () => filteredRows.filter((r) => (r.status === "different" || r.status === "missing") && !excludedParams.has(r.name) && !proposals.get(r.name)?.resolved),
    [filteredRows, excludedParams, proposals]
  );

  const scrollToRow = (name, expandProposal) => {
    const cat = categoryOf(name);
    setCollapsedGroups((prev) => {
      if (!groupByCategory || !prev.has(cat)) return prev;
      const next = new Set(prev);
      next.delete(cat);
      return next;
    });
    if (expandProposal) setOpenProposalRows((prev) => (prev.has(name) ? prev : new Set(prev).add(name)));
    requestAnimationFrame(() => {
      setTimeout(() => {
        const el = document.getElementById(rowDomId(name));
        if (el) {
          el.scrollIntoView({ behavior: "smooth", block: "center" });
          el.style.boxShadow = `inset 0 0 0 2px ${t.accent}`;
          setTimeout(() => { el.style.boxShadow = ""; }, 1400);
        }
      }, 60);
    });
  };

  const jumpToNextDifference = () => {
    if (!diffTargets.length) return;
    const target = diffTargets[jumpIndex % diffTargets.length];
    setJumpIndex((i) => (i + 1) % diffTargets.length);
    scrollToRow(target.name, false);
  };

  const jumpToNextPending = () => {
    if (!pendingTargets.length) return;
    const target = pendingTargets[pendingJumpIndex % pendingTargets.length];
    setPendingJumpIndex((i) => (i + 1) % pendingTargets.length);
    scrollToRow(target.name, true);
  };

  const exportReport = () => {
    if (!profiles.length) return;
    const xml = buildReportWorkbook({ profiles, rows, counts, excludedParams, baselineId });
    downloadBlob(`SAP_Profile_Comparison_${new Date().toISOString().slice(0, 10)}.xls`, xml, "application/vnd.ms-excel");
  };

  const printHtml = (html) => {
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

  const exportPDF = () => {
    if (!profiles.length) return;
    printHtml(buildPrintableReport({ profiles, rows, counts, excludedParams, baselineId }));
  };

  const exportRemediationPDF = () => {
    if (!profiles.length) return;
    printHtml(buildRemediationReport({ profiles, rows, proposals, excludedParams, baselineId }));
  };

  const exportSession = async () => {
    if (!profiles.length) return;
    const data = {
      version: 3,
      exportedAt: new Date().toISOString(),
      baselineId,
      excludedParams: Array.from(excludedParams.entries()),
      proposals: Array.from(proposals.entries()).map(([name, p]) => [
        name,
        { resolved: p.resolved, note: p.note, perProfile: Array.from(p.perProfile.entries()) },
      ]),
      profiles: profiles.map((p) => ({
        id: p.id,
        name: p.name,
        size: p.size,
        sid: p.sid,
        params: Array.from(p.params.entries()),
      })),
    };
    const json = JSON.stringify(data, null, 2);
    const suggestedName = `SAP_Comparison_Session_${new Date().toISOString().slice(0, 10)}.json`;

    if (window.showSaveFilePicker) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName,
          types: [{ description: "SAP Profile Analyzer session", accept: { "application/json": [".json"] } }],
        });
        const writable = await handle.createWritable();
        await writable.write(json);
        await writable.close();
        return;
      } catch (e) {
        if (e.name === "AbortError") return; // user cancelled the picker — do nothing
        // any other error (e.g. picker unsupported in this context) falls through to the plain download below
      }
    }
    downloadBlob(suggestedName, json, "application/json");
  };

  const importSession = async (file) => {
    try {
      if (file.size > MAX_SESSION_FILE_SIZE) {
        throw new Error(`file is ${(file.size / (1024 * 1024)).toFixed(1)}MB, which is larger than the ${MAX_SESSION_FILE_SIZE / (1024 * 1024)}MB limit for a saved session.`);
      }
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
      const rawProposals = Array.isArray(data.proposals) ? data.proposals : [];
      const proposalsMap = new Map(rawProposals.map(([name, p]) => [
        name,
        { resolved: !!p.resolved, note: p.note || "", perProfile: new Map(p.perProfile || []) },
      ]));
      setProfiles(restored);
      setBaselineId(data.baselineId || null);
      setExcludedParams(exclMap);
      setProposals(proposalsMap);
      setOpenProposalRows(new Set());
      setParseWarnings([]);
    } catch (e) {
      setParseWarnings((prev) => [...prev, `Could not load session file: ${e.message}`]);
    }
  };

  const statusStyle = {
    matching: { text: t.matchText, bg: `rgba(${t.matchRgb},0.08)` },
    different: { text: t.diffText, bg: `rgba(${t.diffRgb},0.09)` },
    missing: { text: t.missingText, bg: `rgba(${t.missingRgb},0.08)` },
  };

  const baselineCellStyle = (st) => {
    switch (st) {
      case "baseline": return { text: t.accent, bg: `rgba(${t.accentRgb},0.07)` };
      case "match": return { text: t.matchText, bg: `rgba(${t.matchRgb},0.08)` };
      case "diff": return { text: t.diffText, bg: `rgba(${t.diffRgb},0.1)` };
      case "missing-in-profile": return { text: t.missingText, bg: `rgba(${t.missingRgb},0.1)` };
      case "missing-in-baseline": return { text: t.duplicateText, bg: `rgba(${t.missingRgb},0.08)` };
      case "both-missing": return { text: t.textFaint, bg: "transparent" };
      default: return { text: t.textFaint, bg: "transparent" };
    }
  };

  const renderRow = (r) => {
    const excluded = excludedParams.has(r.name);
    const reason = excludedParams.get(r.name);
    const gutterColor = excluded ? t.textFainter : { matching: t.matchGutter, different: t.diffGutter, missing: t.missingGutter }[r.status];
    const canPropose = !excluded && (r.status === "different" || r.status === "missing");
    const proposal = proposals.get(r.name);
    const isOpen = openProposalRows.has(r.name);
    return (
      <React.Fragment key={r.name}>
        <tr id={rowDomId(r.name)} style={{ borderTop: `1px solid ${t.borderLight}`, opacity: excluded ? 0.5 : 1 }}>
          <td style={{ background: gutterColor, width: 4 }}></td>
          <td style={{ padding: "8px 12px" }}>
            <input type="checkbox" checked={excluded} onChange={() => toggleExcluded(r.name)} title="Exclude from difference analysis" />
          </td>
          <td className="mono" style={{ padding: "8px 12px", fontSize: 12.5, fontWeight: 500, color: t.textPrimary }}>
            {r.name}
            {r.hasDuplicate && (
              <span title="One or more profiles have duplicate entries for this parameter" style={{ marginLeft: 6, color: t.duplicateText, display: "inline-flex", verticalAlign: "middle" }}>
                <Copy size={11} />
              </span>
            )}
            {r.typeMismatch && (
              <span title="Values differ in type (numeric vs text) across profiles — possible misconfiguration" style={{ marginLeft: 6, color: t.typeMismatchText, display: "inline-flex", verticalAlign: "middle" }}>
                <AlertOctagon size={11} />
              </span>
            )}
            {r.sidNormalized && (
              <span title="Values only differ by each profile's own SID (e.g. xs4.example.com vs ps4.example.com) — treated as matching because 'Ignore SID differences' is on" style={{ marginLeft: 6, fontSize: 9, fontWeight: 700, color: t.accent, background: t.accentBg, borderRadius: 4, padding: "1px 4px", verticalAlign: "middle" }}>
                SID
              </span>
            )}
            {excluded && reason && (
              <span title={reason} style={{ marginLeft: 6, color: t.textFaint, display: "inline-flex", verticalAlign: "middle" }}>
                <MessageSquare size={11} />
              </span>
            )}
            {proposal?.resolved && (
              <span title="A fix has been proposed for this parameter" style={{ marginLeft: 6, color: t.matchText, display: "inline-flex", verticalAlign: "middle" }}>
                <CheckCircle2 size={11} />
              </span>
            )}
            {canPropose && (
              <button
                onClick={() => toggleProposalRow(r.name)}
                title="Propose a fix for this parameter"
                style={{
                  marginLeft: 8, display: "inline-flex", alignItems: "center", gap: 4, verticalAlign: "middle",
                  background: isOpen ? t.accentBg : t.panelBg, border: `1px solid ${isOpen ? t.accent : t.border}`,
                  borderRadius: 6, padding: "2px 7px", fontSize: 10.5, fontWeight: 500, color: isOpen ? t.accent : t.textMuted,
                }}
              >
                <Wrench size={10} /> Propose fix
              </button>
            )}
          </td>
          {profiles.map((p) => {
            const cell = r.cellByProfile[p.id];
            let style = { text: t.textPrimary, bg: "transparent" };
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
                style={{ padding: "8px 12px", fontSize: 12, color: style.text, background: style.bg, borderLeft: `1px solid ${t.borderLight}` }}
              >
                {cell ? (
                  <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    {cell.value || <em style={{ color: t.textFaint }}>(empty)</em>}
                    {cell.isDuplicate && <AlertTriangle size={11} color={t.duplicateText} />}
                  </span>
                ) : (
                  <Minus size={12} color={t.textFainter} />
                )}
              </td>
            );
          })}
        </tr>
        {isOpen && canPropose && (
          <tr>
            <td colSpan={profiles.length + 3} style={{ padding: "12px 16px", background: t.panelBgAlt, borderTop: `1px solid ${t.border}`, borderBottom: `1px solid ${t.border}` }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: t.textSecondary }}>
                  Propose fix for <span className="mono">{r.name}</span>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {profiles.map((p) => {
                    const cell = r.cellByProfile[p.id];
                    const draft = proposal?.perProfile.get(p.id) || { action: "keep", value: "" };
                    const baseCell = baselineId ? r.cellByProfile[baselineId] : null;
                    return (
                      <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, flexWrap: "wrap" }}>
                        <span className="mono" style={{ minWidth: 150, color: t.textSecondary }}>
                          {p.name}{p.sid && isDefaultProfile(p.name) ? <span style={{ color: t.accent }}> (SID: {p.sid})</span> : null}
                        </span>
                        <span className="mono" style={{ minWidth: 130, color: t.textFaint }}>{cell ? cell.value : "(not set)"}</span>
                        <select
                          value={draft.action}
                          onChange={(e) => setProposalAction(r.name, p.id, e.target.value)}
                          style={{ background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 6, padding: "5px 8px", fontSize: 12, color: t.textPrimary }}
                        >
                          <option value="keep">Keep as is</option>
                          <option value="set">Set value…</option>
                          <option value="remove">Remove parameter</option>
                        </select>
                        {draft.action === "set" && (
                          <input
                            value={draft.value}
                            onChange={(e) => setProposalValue(r.name, p.id, e.target.value)}
                            placeholder="Proposed value"
                            className="mono"
                            style={{ background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 6, padding: "5px 8px", fontSize: 12, color: t.textPrimary, width: 160 }}
                          />
                        )}
                        {baseCell && p.id !== baselineId && (
                          <button
                            onClick={() => applyBaselineToProposal(r.name, p.id, baseCell.value)}
                            style={{ background: "none", border: `1px solid ${t.border}`, borderRadius: 6, padding: "4px 8px", fontSize: 11, color: t.accent }}
                          >
                            Match baseline
                          </button>
                        )}
                      </div>
                    );
                  })}
                </div>
                <input
                  value={proposal?.note || ""}
                  onChange={(e) => setProposalNote(r.name, e.target.value)}
                  placeholder="Notes for the Basis team (optional)"
                  style={{ background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 6, padding: "6px 10px", fontSize: 12, color: t.textPrimary }}
                />
                <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                  <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: t.textSecondary }}>
                    <input type="checkbox" checked={!!proposal?.resolved} onChange={() => toggleProposalResolved(r.name)} />
                    Mark as decided
                  </label>
                  <span style={{ fontSize: 11, color: t.textFaint }}>
                    Set an action above for the profiles that need to change. If nothing needs to change, checking this alone lists it as "reviewed — no action needed" in the report.
                  </span>
                  <button
                    onClick={() => clearProposal(r.name)}
                    style={{ display: "flex", alignItems: "center", gap: 5, background: "none", border: `1px solid ${t.border}`, borderRadius: 6, padding: "5px 10px", fontSize: 12, color: t.textMuted }}
                  >
                    <Trash2 size={11} /> Clear proposal
                  </button>
                  <button
                    onClick={() => toggleProposalRow(r.name)}
                    style={{ background: "none", border: "none", fontSize: 12, color: t.accent }}
                  >
                    Close
                  </button>
                </div>
              </div>
            </td>
          </tr>
        )}
      </React.Fragment>
    );
  };

  return (
    <div
      style={{
        fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
        background: t.bg,
        color: t.textPrimary,
        minHeight: "100%",
        padding: "28px",
        boxSizing: "border-box",
      }}
    >
      <style>{`
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 10px; height: 10px; }
        ::-webkit-scrollbar-track { background: ${t.scrollTrack}; }
        ::-webkit-scrollbar-thumb { background: ${t.scrollThumb}; border-radius: 6px; }
        .mono { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace; }
        button { font-family: inherit; cursor: pointer; }
        input, select, textarea { font-family: inherit; }
        table { border-collapse: collapse; width: 100%; }
        th, td { text-align: left; }
        .diff-gutter { width: 4px; }
      `}</style>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 34, height: 34, borderRadius: 8, background: t.accentBg, border: `1px solid ${t.accentBorder}`, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Layers size={18} color={t.accent} />
          </div>
          <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0, letterSpacing: "-0.01em", color: t.textPrimary }}>SAP Profile Analyzer</h1>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div role="group" aria-label="Theme" style={{ display: "flex", background: t.panelBgAlt, border: `1px solid ${t.border}`, borderRadius: 8, padding: 3, gap: 2 }}>
            <button
              onClick={() => setTheme("light")}
              title="Light theme"
              style={{
                display: "flex", alignItems: "center", justifyContent: "center", width: 28, height: 28,
                borderRadius: 6, border: "none", background: theme === "light" ? t.accentBg : "transparent",
                color: theme === "light" ? t.accent : t.textFaint,
              }}
            >
              <Sun size={14} />
            </button>
            <button
              onClick={() => setTheme("dark")}
              title="Dark theme"
              style={{
                display: "flex", alignItems: "center", justifyContent: "center", width: 28, height: 28,
                borderRadius: 6, border: "none", background: theme === "dark" ? t.accentBg : "transparent",
                color: theme === "dark" ? t.accent : t.textFaint,
              }}
            >
              <Moon size={14} />
            </button>
            <button
              onClick={() => setTheme("colorblind")}
              title="Colorblind-safe theme"
              style={{
                display: "flex", alignItems: "center", justifyContent: "center", width: 28, height: 28,
                borderRadius: 6, border: "none", background: theme === "colorblind" ? t.accentBg : "transparent",
                color: theme === "colorblind" ? t.accent : t.textFaint,
              }}
            >
              <Contrast size={14} />
            </button>
          </div>
          <button
            onClick={exportSession}
            disabled={!profiles.length}
            title="Save the parsed comparison (not the original files) so you can reload it later"
            style={{ display: "flex", alignItems: "center", gap: 6, background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "7px 12px", fontSize: 12.5, color: profiles.length ? t.textSecondary : t.textFainter }}
          >
            <Save size={13} /> Save comparison
          </button>
          <button
            onClick={() => sessionInputRef.current?.click()}
            title="Load a previously saved comparison session (.json)"
            style={{ display: "flex", alignItems: "center", gap: 6, background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "7px 12px", fontSize: 12.5, color: t.textSecondary }}
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
      <p style={{ color: t.textMuted, fontSize: 13.5, margin: "0 0 22px 46px" }}>
        Upload any SAP profile files to compare them side by side, surface differences and duplicates, export a report for the Basis team, and save comparisons to revisit later.{" "}
        <a href="/SAP_Profile_Analyzer_User_Guide.pdf" target="_blank" rel="noopener noreferrer" style={{ color: t.accent, whiteSpace: "nowrap" }}>
          New here? Read the user guide
        </a>
      </p>

      {/* Upload zone */}
      <div
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
        onClick={() => fileInputRef.current?.click()}
        style={{
          border: `1.5px dashed ${dragOver ? t.accent : t.border}`,
          background: dragOver ? `rgba(${t.accentRgb},0.05)` : t.panelBg,
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
        <Upload size={22} color={t.accent} style={{ marginBottom: 8 }} />
        <div style={{ fontSize: 14, fontWeight: 500, color: t.textPrimary }}>Drop SAP profile files here, or click to browse</div>
        <div style={{ fontSize: 12.5, color: t.textFaint, marginTop: 4 }}>
          Accepts any text-based profile (DEFAULT.PFL, instance profiles, etc.) — parsed as key = value pairs. DEFAULT profiles get their SID auto-detected.
        </div>
      </div>

      <div style={{ textAlign: "center", margin: "10px 0 18px" }}>
        <button
          onClick={(e) => { e.stopPropagation(); setPasteOpen((v) => !v); }}
          style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "none", border: "none", fontSize: 12.5, color: t.accent, fontWeight: 500 }}
        >
          <ClipboardPaste size={13} /> {pasteOpen ? "Cancel pasting text" : "or paste profile text instead"}
        </button>
        {pasteOpen && (
          <div style={{ background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 10, padding: 14, marginTop: 10, textAlign: "left" }}>
            <input
              value={pasteName}
              onChange={(e) => setPasteName(e.target.value)}
              placeholder="Profile name (e.g. DEFAULT.PFL, PRD_DVEBMGS00)"
              style={{ width: "100%", background: t.panelBgAlt, border: `1px solid ${t.border}`, borderRadius: 6, padding: "7px 10px", fontSize: 12.5, marginBottom: 8, color: t.textPrimary }}
            />
            <textarea
              value={pasteText}
              onChange={(e) => setPasteText(e.target.value)}
              placeholder={"login/min_password_lng = 8\nrdisp/wp_no_dia = 20"}
              rows={6}
              className="mono"
              style={{ width: "100%", background: t.panelBgAlt, border: `1px solid ${t.border}`, borderRadius: 6, padding: "8px 10px", fontSize: 12, color: t.textPrimary, resize: "vertical" }}
            />
            <div style={{ marginTop: 8, display: "flex", gap: 8 }}>
              <button
                onClick={addPastedProfile}
                style={{ background: t.accent, border: "none", borderRadius: 6, padding: "6px 12px", fontSize: 12.5, fontWeight: 600, color: "#ffffff" }}
              >
                Add profile
              </button>
              <button
                onClick={() => { setPasteOpen(false); setPasteName(""); setPasteText(""); }}
                style={{ background: "none", border: `1px solid ${t.border}`, borderRadius: 6, padding: "6px 12px", fontSize: 12.5, color: t.textMuted }}
              >
                Cancel
              </button>
            </div>
          </div>
        )}
      </div>

      {parseWarnings.length > 0 && (
        <div style={{ background: t.warningBg, border: `1px solid ${t.warningBorder}`, borderLeft: `4px solid ${t.diffGutter}`, borderRadius: 8, padding: "12px 14px", marginBottom: 18 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 600, color: t.warningText, marginBottom: 6 }}>
            <AlertTriangle size={15} /> Some files need attention
          </div>
          {parseWarnings.map((w, i) => (
            <div key={i} style={{ fontSize: 12, color: t.warningText, marginLeft: 21 }}>{w}</div>
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
                  background: isBaseline ? `rgba(${t.accentRgb},0.06)` : t.panelBg,
                  border: `1px solid ${isBaseline ? t.accent : t.border}`,
                  borderRadius: 8, padding: "7px 10px", fontSize: 12.5,
                }}
              >
                <FileText size={13} color={t.textFaint} />
                <span className="mono" style={{ fontWeight: 500, color: t.textPrimary }}>
                  {p.name}{p.sid && isDefaultProfile(p.name) ? <span style={{ color: t.accent }}> (SID: {p.sid})</span> : null}
                </span>
                <span style={{ color: t.textFaint }}>· {p.params.size} params</span>
                {dupCount > 0 && (
                  <span style={{ display: "flex", alignItems: "center", gap: 3, color: t.duplicateText }}>
                    <Copy size={11} /> {dupCount}
                  </span>
                )}
                <button
                  onClick={() => setBaselineId(isBaseline ? null : p.id)}
                  title={isBaseline ? "Unset as baseline" : "Set as baseline"}
                  style={{ background: "none", border: "none", padding: 2, display: "flex" }}
                >
                  {isBaseline ? <Star size={14} color={t.accent} fill={t.accent} /> : <StarOff size={14} color={t.textFaint} />}
                </button>
                <button onClick={() => removeProfile(p.id)} style={{ background: "none", border: "none", padding: 2, display: "flex" }}>
                  <X size={14} color={t.textFaint} />
                </button>
              </div>
            );
          })}
          <button
            onClick={clearAll}
            style={{ display: "flex", alignItems: "center", gap: 5, background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "7px 10px", fontSize: 12.5, color: t.textMuted }}
          >
            <Trash2 size={12} /> Clear all
          </button>
        </div>
      )}

      {profiles.length === 0 ? (
        <div style={{ textAlign: "center", padding: "48px 20px" }}>
          <Layers size={28} color={t.textFainter} style={{ marginBottom: 10 }} />
          <div style={{ fontSize: 14, fontWeight: 500, color: t.textMuted, marginBottom: 4 }}>No profiles loaded yet</div>
          <div style={{ fontSize: 12.5, color: t.textFaint }}>Upload at least two profiles to see a comparison, paste profile text, or load a previously saved comparison.</div>
        </div>
      ) : (
        <>
          {/* Summary stats */}
          <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
            {[
              { label: "Parameters", value: counts.total, color: t.textPrimary },
              { label: "Matching", value: counts.matching, color: t.matchText },
              { label: "Different", value: counts.different, color: t.diffText },
              { label: "Missing", value: counts.missing, color: t.missingText },
              { label: "Duplicates", value: counts.duplicates, color: t.duplicateText },
              { label: "Type mismatches", value: counts.typeMismatches, color: t.typeMismatchText },
              { label: "Excluded", value: counts.excluded, color: t.textMuted },
              { label: "Decided", value: `${counts.decided}/${counts.actionable}`, color: t.accent },
              ...(ignoreSidDiffs ? [{ label: "SID-normalized", value: counts.sidNormalized, color: t.accent }] : []),
            ].map((s) => (
              <div key={s.label} style={{ background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "10px 16px", minWidth: 88 }}>
                <div className="mono" style={{ fontSize: 19, fontWeight: 600, color: s.color }}>{s.value}</div>
                <div style={{ fontSize: 11, color: t.textFaint, marginTop: 2 }}>{s.label}</div>
              </div>
            ))}
          </div>

          {/* Match / different / missing overview bar */}
          {(counts.matching + counts.different + counts.missing) > 0 && (
            <div style={{ marginBottom: 18 }}>
              <div style={{ display: "flex", height: 14, borderRadius: 7, overflow: "hidden", border: `1px solid ${t.border}` }}>
                {counts.matching > 0 && (
                  <div style={{ width: `${(counts.matching / (counts.matching + counts.different + counts.missing)) * 100}%`, background: t.matchGutter }} title={`Matching: ${counts.matching}`}></div>
                )}
                {counts.different > 0 && (
                  <div style={{ width: `${(counts.different / (counts.matching + counts.different + counts.missing)) * 100}%`, background: t.diffGutter }} title={`Different: ${counts.different}`}></div>
                )}
                {counts.missing > 0 && (
                  <div style={{ width: `${(counts.missing / (counts.matching + counts.different + counts.missing)) * 100}%`, background: t.missingGutter }} title={`Missing: ${counts.missing}`}></div>
                )}
              </div>
              <div style={{ display: "flex", gap: 16, marginTop: 6, fontSize: 11.5, color: t.textMuted, flexWrap: "wrap" }}>
                {(() => {
                  const denom = counts.matching + counts.different + counts.missing;
                  const pct = (n) => Math.round((n / denom) * 100);
                  return (
                    <>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <span style={{ width: 8, height: 8, borderRadius: 2, background: t.matchGutter, display: "inline-block" }}></span>
                        Matching {counts.matching} ({pct(counts.matching)}%)
                      </span>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <span style={{ width: 8, height: 8, borderRadius: 2, background: t.diffGutter, display: "inline-block" }}></span>
                        Different {counts.different} ({pct(counts.different)}%)
                      </span>
                      <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <span style={{ width: 8, height: 8, borderRadius: 2, background: t.missingGutter, display: "inline-block" }}></span>
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
            <div style={{ display: "flex", flexWrap: "wrap", background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: 3 }}>
              {VIEW_MODES.map((m) => (
                <button
                  key={m.id}
                  onClick={() => setViewMode(m.id)}
                  style={{
                    background: viewMode === m.id ? t.accentBg : "transparent",
                    color: viewMode === m.id ? t.accent : t.textMuted,
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
                    background: viewMode === "baseline" ? t.accentBg : "transparent",
                    color: viewMode === "baseline" ? t.accent : t.textMuted,
                    border: "none", borderRadius: 6, padding: "6px 12px", fontSize: 12.5, fontWeight: 500,
                    display: "flex", alignItems: "center", gap: 4,
                  }}
                >
                  <Star size={11} /> vs Baseline
                </button>
              )}
            </div>

            <div style={{ position: "relative" }}>
              <Search size={13} color={t.textFaint} style={{ position: "absolute", left: 10, top: 9 }} />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Filter parameters…"
                style={{
                  background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8,
                  padding: "7px 10px 7px 30px", fontSize: 12.5, color: t.textPrimary, width: 190, outline: "none",
                }}
              />
            </div>

            <button
              onClick={() => setShowExcluded((v) => !v)}
              style={{ display: "flex", alignItems: "center", gap: 6, background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "7px 10px", fontSize: 12.5, color: t.textMuted }}
            >
              {showExcluded ? <Eye size={13} /> : <EyeOff size={13} />}
              {showExcluded ? "Showing excluded" : "Hiding excluded"}
            </button>

            <button
              onClick={() => setShowExclusionsPanel((v) => !v)}
              style={{ display: "flex", alignItems: "center", gap: 6, background: showExclusionsPanel ? t.accentBg : t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "7px 10px", fontSize: 12.5, color: showExclusionsPanel ? t.accent : t.textMuted }}
            >
              <MessageSquare size={13} /> Exclusions{excludedParams.size > 0 ? ` (${excludedParams.size})` : ""}
            </button>

            {profiles.some((p) => p.sid) && (
              <button
                onClick={() => setIgnoreSidDiffs((v) => !v)}
                title="Treat values that only differ because each profile's own SID appears in them (e.g. xs4.example.com vs ps4.example.com) as matching, not different"
                style={{ display: "flex", alignItems: "center", gap: 6, background: ignoreSidDiffs ? t.accentBg : t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "7px 10px", fontSize: 12.5, color: ignoreSidDiffs ? t.accent : t.textMuted }}
              >
                <Fingerprint size={13} /> {ignoreSidDiffs ? "Ignoring SID differences" : "Ignore SID differences"}
              </button>
            )}

            <div style={{ flex: 1 }} />

            <button
              onClick={exportReport}
              style={{ display: "flex", alignItems: "center", gap: 6, background: t.accent, border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, color: "#ffffff" }}
            >
              <Download size={13} /> Excel report
            </button>
            <button
              onClick={exportPDF}
              style={{ display: "flex", alignItems: "center", gap: 6, background: t.panelBg, border: `1px solid ${t.accent}`, borderRadius: 8, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, color: t.accent }}
            >
              <Printer size={13} /> PDF report
            </button>
            <button
              onClick={exportRemediationPDF}
              disabled={counts.decided === 0}
              title={counts.decided === 0 ? "Propose at least one fix and mark it as decided to enable this" : "Export the parameters you've decided how to fix, grouped by profile"}
              style={{ display: "flex", alignItems: "center", gap: 6, background: t.panelBg, border: `1px solid ${counts.decided ? t.matchText : t.border}`, borderRadius: 8, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, color: counts.decided ? t.matchText : t.textFainter }}
            >
              <ClipboardCheck size={13} /> Remediation plan
            </button>
          </div>

          {/* Exclusions panel */}
          {showExclusionsPanel && (
            <div style={{ border: `1px solid ${t.border}`, borderRadius: 10, background: t.panelBg, padding: 14, marginBottom: 14 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, marginBottom: 8, color: t.textSecondary }}>Excluded parameters ({excludedParams.size})</div>
              {excludedParams.size === 0 ? (
                <div style={{ fontSize: 12, color: t.textFaint }}>No parameters excluded yet. Check the "Excl." box next to a parameter row to exclude it, then note why here.</div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 220, overflowY: "auto" }}>
                  {Array.from(excludedParams.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([name, reason]) => (
                    <div key={name} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span className="mono" style={{ fontSize: 12, minWidth: 220, color: t.textSecondary }}>{name}</span>
                      <input
                        value={reason}
                        onChange={(e) => setExclusionReason(name, e.target.value)}
                        placeholder="Reason (optional) — e.g. expected to differ per system"
                        style={{ flex: 1, background: t.panelBgAlt, border: `1px solid ${t.border}`, borderRadius: 6, padding: "5px 8px", fontSize: 12, color: t.textPrimary }}
                      />
                      <button onClick={() => toggleExcluded(name)} style={{ background: "none", border: "none", display: "flex" }} title="Remove exclusion">
                        <X size={13} color={t.textFaint} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {baselineId && (
            <div style={{ fontSize: 12, color: t.accent, marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}>
              <Star size={12} fill={t.accent} /> Baseline: <span className="mono">{profiles.find((p) => p.id === baselineId)?.name}</span> — other profiles are evaluated against it, and the report includes recommended actions.
            </div>
          )}

          {/* Legend + table toolbar */}
          <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 10 }}>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 14, fontSize: 11.5, color: t.textMuted }}>
              <LegendItem color={t.matchGutter} label="Matching" />
              <LegendItem color={t.diffGutter} label="Different" />
              <LegendItem color={t.missingGutter} label="Missing" />
              <LegendItem color={t.duplicateText} label="Duplicate entry" />
              <LegendItem color={t.typeMismatchText} label="Type mismatch" />
              <LegendItem color={t.accent} label="Baseline column" />
              <LegendItem color={t.textFaint} label="Excluded" />
              <LegendItem color={t.matchText} label="Fix decided" />
              {ignoreSidDiffs && <LegendItem color={t.accent} label="Matched after ignoring SID" />}
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button
                onClick={() => setGroupByCategory((v) => !v)}
                style={{ display: "flex", alignItems: "center", gap: 5, background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "6px 10px", fontSize: 12, color: t.textMuted }}
              >
                <Layers size={12} /> {groupByCategory ? "Grouped by category" : "Flat list"}
              </button>
              {groupByCategory && (
                <>
                  <button onClick={() => setCollapsedGroups(new Set())} style={{ background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "6px 10px", fontSize: 12, color: t.textMuted }}>
                    Expand all
                  </button>
                  <button onClick={() => setCollapsedGroups(new Set(groups.map(([cat]) => cat)))} style={{ background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "6px 10px", fontSize: 12, color: t.textMuted }}>
                    Collapse all
                  </button>
                </>
              )}
              <button
                onClick={jumpToNextDifference}
                disabled={!diffTargets.length}
                style={{ display: "flex", alignItems: "center", gap: 5, background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "6px 10px", fontSize: 12, color: diffTargets.length ? t.accent : t.textFainter }}
              >
                <Target size={12} /> Jump to next difference{diffTargets.length ? ` (${diffTargets.length})` : ""}
              </button>
              <button
                onClick={jumpToNextPending}
                disabled={!pendingTargets.length}
                title="Jump to the next difference or missing parameter that doesn't have a decided proposal yet"
                style={{ display: "flex", alignItems: "center", gap: 5, background: t.panelBg, border: `1px solid ${t.border}`, borderRadius: 8, padding: "6px 10px", fontSize: 12, color: pendingTargets.length ? t.accent : t.textFainter }}
              >
                <ListTodo size={12} /> Jump to next pending decision{pendingTargets.length ? ` (${pendingTargets.length})` : ""}
              </button>
            </div>
          </div>

          {/* Table */}
          <div style={{ border: `1px solid ${t.border}`, borderRadius: 10, overflow: "hidden", background: t.panelBg }}>
            <div style={{ overflowX: "auto", maxHeight: 560, overflowY: "auto" }}>
              <table>
                <thead>
                  <tr style={{ background: t.panelBgAlt, position: "sticky", top: 0, zIndex: 2 }}>
                    <th className="diff-gutter"></th>
                    <th style={{ padding: "10px 12px", fontSize: 11.5, color: t.textMuted, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.03em", minWidth: 60 }}>Excl.</th>
                    <th style={{ padding: "10px 12px", fontSize: 11.5, color: t.textMuted, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.03em", minWidth: 220 }}>Parameter</th>
                    {profiles.map((p) => (
                      <th key={p.id} style={{ padding: "8px 12px", fontSize: 11.5, color: p.id === baselineId ? t.accent : t.textMuted, fontWeight: 600, minWidth: 160, borderLeft: `1px solid ${t.borderLight}` }}>
                        <div className="mono" style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                          <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                            {p.id === baselineId && <Star size={11} fill={t.accent} />}
                            {p.name}
                          </span>
                          {p.sid && isDefaultProfile(p.name) && <span style={{ fontSize: 10, fontWeight: 600, color: t.accent }}>SID: {p.sid}</span>}
                          <span style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 2 }}>
                            <button onClick={() => moveProfile(p.id, -1)} title="Move column left" style={{ background: "none", border: "none", padding: 0, display: "flex" }}>
                              <ChevronLeft size={12} color={t.textFainter} />
                            </button>
                            <button onClick={() => pinProfile(p.id)} title="Pin to front" style={{ background: "none", border: "none", padding: 0, display: "flex" }}>
                              <Pin size={11} color={t.textFainter} />
                            </button>
                            <button onClick={() => moveProfile(p.id, 1)} title="Move column right" style={{ background: "none", border: "none", padding: 0, display: "flex" }}>
                              <ChevronRight size={12} color={t.textFainter} />
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
                        <tr onClick={() => toggleGroup(cat)} style={{ background: t.panelBgAlt, cursor: "pointer", borderTop: `1px solid ${t.border}` }}>
                          <td colSpan={profiles.length + 3} style={{ padding: "7px 12px", fontSize: 12, fontWeight: 600, color: t.textSecondary }}>
                            <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              <ChevronDown size={13} style={{ transform: collapsedGroups.has(cat) ? "rotate(-90deg)" : "none", transition: "transform 0.1s" }} />
                              {cat === "other" ? "Other parameters" : cat}
                              <span style={{ fontWeight: 400, color: t.textFaint }}>· {catRows.length} parameter{catRows.length === 1 ? "" : "s"}</span>
                            </span>
                          </td>
                        </tr>
                      )}
                      {!collapsedGroups.has(cat) && catRows.map((r) => renderRow(r))}
                    </React.Fragment>
                  ))}
                  {filteredRows.length === 0 && (
                    <tr>
                      <td colSpan={profiles.length + 3} style={{ padding: 24, textAlign: "center", color: t.textFaint, fontSize: 13 }}>
                        No parameters match this view.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
          <div style={{ fontSize: 11.5, color: t.textFaint, marginTop: 8 }}>
            Showing {filteredRows.length} of {rows.length} parameters. Data stays in this browser tab only — export a report or save a comparison to keep it.
          </div>
        </>
      )}

      <div style={{ marginTop: 32, paddingTop: 16, borderTop: `1px solid ${t.border}`, textAlign: "center", fontSize: 11.5, color: t.textFaint }}>
        SAP Profile Analyzer · Built by {BRAND_NAME} ·{" "}
        <a href={`mailto:${BRAND_EMAIL}`} style={{ color: t.textFaint }}>{BRAND_EMAIL}</a>
        {" · "}
        <a href={BRAND_LINKEDIN} target="_blank" rel="noopener noreferrer" style={{ color: t.textFaint }}>{BRAND_LINKEDIN}</a>
      </div>
    </div>
  );
}
