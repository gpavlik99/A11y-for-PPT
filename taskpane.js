
/* global Office, PowerPoint */

/**
 * LCM PowerPoint A11y Checker — v4 (Desktop + Web)
 * Adds:
 * - Severity buckets + filters (Critical/Serious/Moderate/Minor)
 * - CSV export
 * - "Skipped in this environment" callouts in the Checks list
 * - Per-user resolved/intentional markers (RoamingSettings) from v3
 */

const CHECKS = [
  { id: "slideTitles", label: "Slides should have titles", fn: checkSlideTitles, defaultSeverity: "serious" },
  { id: "duplicateTitles", label: "Avoid duplicate slide titles (warn)", fn: checkDuplicateTitles, defaultSeverity: "minor", isWarning: true },
  { id: "emptySlides", label: "Slides should not be empty", fn: checkEmptySlides, defaultSeverity: "serious" },
  { id: "textSize", label: "Text should be readable (min size)", fn: checkTextSize, defaultSeverity: "moderate" },
  { id: "textFormatting", label: "Avoid excessive text styling", fn: checkTextFormatting, defaultSeverity: "minor" },
  { id: "manualListFormatting", label: "Use real lists (avoid manual bullets)", fn: checkManualListFormatting, defaultSeverity: "moderate" },
  { id: "altText", label: "Images/shapes should have alt text", fn: checkAltText, defaultSeverity: "serious" },
  { id: "overlappingElements", label: "Overlapping elements may impact reading order", fn: checkOverlappingElements, defaultSeverity: "minor", isWarning: true },
  { id: "vagueLinkText", label: "Links should be descriptive (best-effort)", fn: checkVagueLinks, defaultSeverity: "moderate" },
];

let isScanning = false;
let lastScan = null;
let resolvedIndex = new Set();
let uiState = {
  hideResolved: true,
  showWarnings: true,
  severity: { critical: true, serious: true, moderate: true, minor: true }
};

Office.onReady(async () => {
  wireUi();
  await loadResolvedIndex();
  await initSlideCount();
  renderScoringSummary(null);
});

function wireUi() {
  document.getElementById("check-accessibility").addEventListener("click", runScan);
  document.getElementById("export-json").addEventListener("click", exportJson);
  const exportCsvBtn = document.getElementById("export-csv");
  if (exportCsvBtn) exportCsvBtn.addEventListener("click", exportCsv);

  const clearBtn = document.getElementById("clear-resolved");
  if (clearBtn) clearBtn.addEventListener("click", clearResolved);

  const hideResolvedToggle = document.getElementById("toggle-hide-resolved");
  if (hideResolvedToggle) hideResolvedToggle.addEventListener("change", () => {
    uiState.hideResolved = hideResolvedToggle.checked;
    rerenderIssuesFromLastScan();
  });

  const showWarningsToggle = document.getElementById("toggle-show-warnings");
  if (showWarningsToggle) showWarningsToggle.addEventListener("change", () => {
    uiState.showWarnings = showWarningsToggle.checked;
    rerenderIssuesFromLastScan();
  });

  // Severity bucket toggles
  ["critical","serious","moderate","minor"].forEach(sev => {
    const el = document.getElementById(`sev-${sev}`);
    if (!el) return;
    el.addEventListener("change", () => {
      uiState.severity[sev] = el.checked;
      rerenderIssuesFromLastScan();
    });
  });

  document.querySelectorAll('input[name="scanMode"]').forEach(r => r.addEventListener("change", handleScanModeChange));

  ["range-from","range-to"].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.addEventListener("input", updateRangeDisplay); el.addEventListener("change", updateRangeDisplay); }
  });
  ["range-from-number","range-to-number"].forEach(id => {
    const el = document.getElementById(id);
    if (el) { el.addEventListener("input", syncRangeInputs); el.addEventListener("change", syncRangeInputs); }
  });

  handleScanModeChange();
}

function setSummary(text) { document.getElementById("summary-text").textContent = text; }
function setProgress(pct) {
  document.getElementById("progress-container").style.display = "block";
  document.getElementById("progress-fill").style.width = `${pct}%`;
}
function hideProgress() {
  document.getElementById("progress-container").style.display = "none";
  document.getElementById("progress-fill").style.width = "0%";
}

async function initSlideCount() {
  try { updateSlideCount(await getTotalSlideCount()); }
  catch { updateSlideCount(1); }
}

function handleScanModeChange() {
  const mode = getScanConfig().mode;
  document.getElementById("page-range-container").style.display = mode === "range" ? "block" : "none";
  updateRangeDisplay();
}

function updateSlideCount(totalSlides) {
  const ids = ['range-from','range-to','range-from-number','range-to-number'];
  ids.forEach(id => { const el = document.getElementById(id); if (el) el.max = totalSlides; });
  const toS = document.getElementById('range-to');
  const toN = document.getElementById('range-to-number');
  if (toS) toS.value = String(totalSlides);
  if (toN) toN.value = String(totalSlides);
  updateRangeDisplay();
}

function clampInt(val, min, max) {
  const n = parseInt(String(val||""),10);
  if (Number.isNaN(n)) return min;
  return Math.max(min, Math.min(max, n));
}

function syncRangeInputs() {
  const fromN = document.getElementById('range-from-number');
  const toN = document.getElementById('range-to-number');
  const fromS = document.getElementById('range-from');
  const toS = document.getElementById('range-to');

  let from = clampInt(fromN.value, 1, parseInt(fromN.max||"1",10));
  let to = clampInt(toN.value, 1, parseInt(toN.max||"1",10));
  if (from > to) [from,to] = [to,from];

  fromN.value = String(from); toN.value = String(to);
  fromS.value = String(from); toS.value = String(to);
  updateRangeDisplay();
}

function updateRangeDisplay() {
  const fromS = document.getElementById('range-from');
  const toS = document.getElementById('range-to');
  const fromN = document.getElementById('range-from-number');
  const toN = document.getElementById('range-to-number');

  let from = clampInt(fromS.value, 1, parseInt(fromS.max||"1",10));
  let to = clampInt(toS.value, 1, parseInt(toS.max||"1",10));
  if (from > to) [from,to] = [to,from];

  fromS.value = String(from); toS.value = String(to);
  fromN.value = String(from); toN.value = String(to);

  document.getElementById('range-from-display').textContent = String(from);
  document.getElementById('range-to-display').textContent = String(to);
  document.getElementById('range-total-display').textContent = String(Math.max(0, to-from+1));

  const max = parseInt(fromS.max||"1",10);
  const leftPct = ((from-1)/Math.max(1,max-1))*100;
  const rightPct = ((to-1)/Math.max(1,max-1))*100;
  const fill = document.getElementById("range-fill");
  fill.style.left = `${leftPct}%`;
  fill.style.width = `${Math.max(0,rightPct-leftPct)}%`;
}

function getScanConfig() {
  const mode = document.getElementById("scan-range").checked ? "range" : "all";
  return {
    mode,
    fromSlide: parseInt(document.getElementById("range-from").value||"1",10),
    toSlide: parseInt(document.getElementById("range-to").value||"1",10),
  };
}

/* ---------------------------
   Per-user state (RoamingSettings)
---------------------------- */

function getDocNamespace() {
  const url = (Office.context && Office.context.document && Office.context.document.url) ? Office.context.document.url : "unsaved";
  return `lcmPptA11y:${url}`;
}

function issueKey(issue) {
  const ns = getDocNamespace();
  const slide = issue.slideNum ?? 0;
  const check = issue.check ?? "unknown";
  const shape = issue.shapeId ?? "";
  const extra = issue.extraKey ?? "";
  return `${ns}|${check}|s${slide}|sh${shape}|${extra}`;
}

function getSettingsBag() {
  const rs = Office.context.roamingSettings;
  const bag = rs.get(getDocNamespace());
  if (bag && typeof bag === "object") return bag;
  return { resolved: {}, intentional: {}, lastScanCounts: null };
}

function setSettingsBag(bag) {
  const rs = Office.context.roamingSettings;
  rs.set(getDocNamespace(), bag);
  return new Promise((resolve) => rs.saveAsync(() => resolve()));
}



function getLastScanCounts() {
  try {
    const bag = getSettingsBag();
    return bag.lastScanCounts || null;
  } catch {
    return null;
  }
}

async function setLastScanCounts(counts) {
  const bag = getSettingsBag();
  bag.lastScanCounts = counts;
  await setSettingsBag(bag);
}

async function loadResolvedIndex() {
  try {
    const bag = getSettingsBag();
    const resolved = bag.resolved || {};
    const intentional = bag.intentional || {};
    resolvedIndex = new Set([...Object.keys(resolved), ...Object.keys(intentional)]);
  } catch {
    resolvedIndex = new Set();
  }
}

async function markResolved(issue, kind = "resolved") {
  const key = issueKey(issue);
  const bag = getSettingsBag();
  if (kind === "resolved") bag.resolved[key] = { at: new Date().toISOString() };
  if (kind === "intentional") bag.intentional[key] = { at: new Date().toISOString() };
  await setSettingsBag(bag);
  resolvedIndex.add(key);
  rerenderIssuesFromLastScan();
  const _counts = renderScoringSummary(lastScan);
  if (_counts) { await setLastScanCounts(_counts); }
}

async function clearResolved() {
  await setSettingsBag({ resolved: {}, intentional: {} });
  resolvedIndex = new Set();
  rerenderIssuesFromLastScan();
  const _counts = renderScoringSummary(lastScan);
  if (_counts) { await setLastScanCounts(_counts); }
}

/* ---------------------------
   Scan runner + "skipped" callouts
---------------------------- */

async function runScan() {
  if (isScanning) return;
  isScanning = true;

  document.getElementById("results-list").innerHTML = "";
  document.getElementById("issues-container").innerHTML = '<div class="muted">Scanning…</div>';
  document.getElementById("panel-badge").classList.add("hidden");

  setSummary("Scanning…");
  setProgress(1);

  const scanConfig = getScanConfig();
  const allIssues = [];
  const perCheck = [];

  for (let i=0;i<CHECKS.length;i++) {
    const c = CHECKS[i];
    renderCheckRow(c.id, c.label);
    setProgress(Math.round((i/CHECKS.length)*100));
    try {
      const res = await c.fn(scanConfig);

      // normalize
      const normalized = {
        name: c.id,
        success: !!res.success,
        message: res.message || "",
        skipped: !!res.skipped,
        details: Array.isArray(res.details) ? res.details : []
      };
      perCheck.push(normalized);

      if (normalized.skipped) {
        updateCheckRow(c.id, "skipped", "Skipped in this environment");
      } else {
        const isWarn = !!c.isWarning;
        updateCheckRow(c.id, (normalized.success || isWarn) ? "success" : "failed", normalized.message);
      }

      if (Array.isArray(normalized.details)) {
        allIssues.push(...normalized.details.map(d => ({
          check: c.id,
          severity: normalizeSeverity(d.severity || c.defaultSeverity || "moderate"),
          ...d
        })));
      }
    } catch (e) {
      perCheck.push({name:c.id, success:false, message:e.message||String(e), skipped:false, details:[]});
      updateCheckRow(c.id, "failed", e.message||String(e));
    }
  }

  setProgress(100); hideProgress();

  lastScan = { time:new Date().toISOString(), scanConfig, perCheck, issues: allIssues };

  const _counts = renderScoringSummary(lastScan);
  if (_counts) { await setLastScanCounts(_counts); }

  rerenderIssuesFromLastScan();

  const failed = perCheck.filter(r=>!r.success && !r.skipped).length;
  setSummary(failed===0 ? "✅ Scan complete" : `⚠️ Scan complete (${failed} check(s) flagged)`);

  const badge = document.getElementById("panel-badge");
  badge.textContent = failed===0 ? "Complete" : "Needs review";
  badge.className = failed===0 ? "pill success" : "pill failed";
  badge.classList.remove("hidden");

  isScanning = false;
}

function normalizeSeverity(sev) {
  const s = String(sev || "").toLowerCase().trim();
  if (["critical","serious","moderate","minor"].includes(s)) return s;
  // map legacy "warning" to minor
  if (s === "warning") return "minor";
  return "moderate";
}

function rerenderIssuesFromLastScan() {
  if (!lastScan) {
    document.getElementById("issues-container").innerHTML = '<div class="muted">Run a scan to see issues.</div>';
    return;
  }
  renderIssues(lastScan.issues || []);
}

function renderCheckRow(id, label) {
  const li = document.createElement("li");
  li.className = "li";
  li.innerHTML = `
    <span class="check unchecked" id="${id}-dot"></span>
    <div class="text">
      <strong>${escapeHtml(label)}</strong><br/>
      <span class="muted" id="${id}-msg">Running…</span>
    </div>`;
  document.getElementById("results-list").appendChild(li);
}

function updateCheckRow(id, state, msg) {
  const dot = document.getElementById(`${id}-dot`);
  if (state === "skipped") {
    dot.className = "check unchecked";
    const m = document.getElementById(`${id}-msg`);
    m.innerHTML = `${escapeHtml(msg || "")} <span class="pill skipped" style="margin-left:6px;">Skipped</span>`;
    return;
  }
  dot.className = `check ${state}`;
  document.getElementById(`${id}-msg`).textContent = msg || "";
}

function renderIssues(issues) {
  const container = document.getElementById("issues-container");
  const hideResolved = !!uiState.hideResolved;
  const showWarnings = !!uiState.showWarnings;

  const visible = issues.filter(issue => {
    const sev = normalizeSeverity(issue.severity);
    if (!uiState.severity[sev]) return false;

    if (!showWarnings && (issue.isWarning === true)) return false;

    if (!hideResolved) return true;
    return !resolvedIndex.has(issueKey(issue));
  });

  if (!visible.length) {
    container.innerHTML = '<div class="muted">No issues to show (based on your filters).</div>';
    return;
  }
  container.innerHTML = "";

  const bySlide = new Map();
  for (const issue of visible) {
    const s = issue.slideNum || 0;
    if (!bySlide.has(s)) bySlide.set(s, []);
    bySlide.get(s).push(issue);
  }

  const slides = Array.from(bySlide.keys()).sort((a,b)=>a-b);
  for (const slideNum of slides) {
    const header = document.createElement("div");
    header.className = "issue-item";
    header.style.borderLeftColor = "#555";
    header.innerHTML = `<div class="issue-text">Slide ${slideNum}</div>`;
    container.appendChild(header);

    for (const issue of bySlide.get(slideNum)) {
      const sev = normalizeSeverity(issue.severity);
      const div = document.createElement("div");
      div.className = "issue-item";
      div.innerHTML = `
        <div class="issue-text">
          <span class="badge ${sev}">${sev.toUpperCase()}</span>
          <span style="margin-left:6px;">${escapeHtml(issue.title || issue.check)}</span>
        </div>
        <div class="issue-meta">${escapeHtml(issue.description || "")}</div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:6px;">
          <button class="btn small" data-slide="${slideNum}">Go to slide</button>
          <button class="btn small" data-action="resolve">Mark resolved</button>
          ${issue.check === "duplicateTitles" ? `<button class="btn small" data-action="intentional">Mark intentional</button>` : ""}
        </div>
      `;
      div.querySelector("button[data-slide]").addEventListener("click", () => selectSlide(slideNum));
      div.querySelector("button[data-action='resolve']").addEventListener("click", () => markResolved(issue, "resolved"));
      const intentionalBtn = div.querySelector("button[data-action='intentional']");
      if (intentionalBtn) intentionalBtn.addEventListener("click", () => markResolved(issue, "intentional"));
      container.appendChild(div);
    }
  }
}




/* ---------------------------
   Scoring Summary
---------------------------- */

function renderScoringSummary(scan) {
  const openTotalEl = document.getElementById("score-open-total");
  const openFilteredEl = document.getElementById("score-open-filtered-total");
  const resolvedTotalEl = document.getElementById("score-resolved-total");
  const openCriticalEl = document.getElementById("score-open-critical");
  const openSeriousEl = document.getElementById("score-open-serious");
  const openModerateEl = document.getElementById("score-open-moderate");
  const openMinorEl = document.getElementById("score-open-minor");
  const footEl = document.getElementById("score-foot");
  const deltaEl = document.getElementById("score-delta");
  const gateEl = document.getElementById("score-gate");
  const badgeEl = document.getElementById("scoring-badge");

  if (!openTotalEl || !resolvedTotalEl || !openCriticalEl) return null;

  if (!scan || !Array.isArray(scan.issues)) {
    [openTotalEl, openFilteredEl, resolvedTotalEl, openCriticalEl, openSeriousEl, openModerateEl, openMinorEl]
      .forEach(el => { if (el) el.textContent = "—"; });
    if (footEl) footEl.textContent = "Run a scan to see counts.";
    if (deltaEl) deltaEl.textContent = "";
    if (gateEl) gateEl.textContent = "";
    if (badgeEl) badgeEl.classList.add("hidden");
    return null;
  }

  const counts = {
    open: { total: 0, critical: 0, serious: 0, moderate: 0, minor: 0 },
    resolved: { total: 0 },
    filteredOpen: 0
  };

  for (const issue of scan.issues) {
    const sev = normalizeSeverity(issue.severity);
    const key = issueKey(issue);
    const isResolved = resolvedIndex.has(key);

    if (isResolved) {
      counts.resolved.total += 1;
      continue;
    }

    counts.open.total += 1;
    if (counts.open[sev] !== undefined) counts.open[sev] += 1;

    // Apply current UI filters (what the issues list shows)
    const passesSeverity = uiState.severity[sev] !== false;
    const passesWarnings = uiState.showWarnings || !issue.isWarning;
    const passesResolved = !uiState.hideResolved || !isResolved;

    if (passesSeverity && passesWarnings && passesResolved) {
      counts.filteredOpen += 1;
    }
  }

  openTotalEl.textContent = String(counts.open.total);
  if (openFilteredEl) openFilteredEl.textContent = String(counts.filteredOpen);
  resolvedTotalEl.textContent = String(counts.resolved.total);
  openCriticalEl.textContent = String(counts.open.critical);
  openSeriousEl.textContent = String(counts.open.serious);
  openModerateEl.textContent = String(counts.open.moderate);
  openMinorEl.textContent = String(counts.open.minor);

  if (footEl) footEl.textContent = "Counts reflect your per-user resolved state and current filters.";

  // PASS/FAIL gate (simple, explicit)
  const pass = (counts.open.critical === 0 && counts.open.serious === 0);
  if (gateEl) {
    gateEl.textContent = pass
      ? "Gate: PASS (0 Critical, 0 Serious)"
      : `Gate: FAIL (${counts.open.critical} Critical, ${counts.open.serious} Serious)`;
  }

  if (badgeEl) {
    badgeEl.classList.remove("hidden");
    if (pass) {
      badgeEl.textContent = "On track";
      badgeEl.className = "pill success";
    } else {
      badgeEl.textContent = "Needs attention";
      badgeEl.className = "pill failed";
    }
  }

  // Trend delta vs previous scan (per-user)
  const prev = getLastScanCounts();
  if (deltaEl) {
    if (!prev || !prev.open) {
      deltaEl.textContent = "Delta vs last scan: —";
    } else {
      const dTotal = (counts.open.total - (prev.open.total || 0));
      const dC = (counts.open.critical - (prev.open.critical || 0));
      const dS = (counts.open.serious - (prev.open.serious || 0));
      const dM = (counts.open.moderate - (prev.open.moderate || 0));
      const dMi = (counts.open.minor - (prev.open.minor || 0));

      const fmt = (n) => (n === 0 ? "0" : (n > 0 ? `+${n}` : `${n}`));
      deltaEl.textContent = `Delta vs last scan (open): Total ${fmt(dTotal)} | Critical ${fmt(dC)} | Serious ${fmt(dS)} | Moderate ${fmt(dM)} | Minor ${fmt(dMi)}`;
    }
  }

  return counts;
}

  const counts = {
    open: { total: 0, critical: 0, serious: 0, moderate: 0, minor: 0 },
    resolved: { total: 0 }
  };

  for (const issue of scan.issues) {
    const sev = normalizeSeverity(issue.severity);
    const key = issueKey(issue);
    const isResolved = resolvedIndex.has(key);

    if (isResolved) {
      counts.resolved.total += 1;
      continue;
    }

    counts.open.total += 1;
    if (counts.open[sev] !== undefined) counts.open[sev] += 1;
  }

  openTotalEl.textContent = String(counts.open.total);
  resolvedTotalEl.textContent = String(counts.resolved.total);
  openCriticalEl.textContent = String(counts.open.critical);
  openSeriousEl.textContent = String(counts.open.serious);
  openModerateEl.textContent = String(counts.open.moderate);
  openMinorEl.textContent = String(counts.open.minor);

  footEl.textContent = "Counts reflect your per-user resolved state for this deck.";

  if (badgeEl) {
    badgeEl.classList.remove("hidden");
    if (counts.open.critical === 0 && counts.open.serious === 0) {
      badgeEl.textContent = "On track";
      badgeEl.className = "pill success";
    } else {
      badgeEl.textContent = "Needs attention";
      badgeEl.className = "pill failed";
    }
  }
}

/* ---------------------------
   Exporters
---------------------------- */

function exportJson() {
  if (!lastScan) { alert("Run a scan first."); return; }
  const blob = new Blob([JSON.stringify(lastScan, null, 2)], { type:"application/json" });
  downloadBlob(blob, `lcm-ppt-a11y-scan-${stamp()}.json`);
}

function exportCsv() {
  if (!lastScan) { alert("Run a scan first."); return; }
  const rows = [];

  // Header
  rows.push([

    "scan_time",
    "check",
    "severity",
    "slide_num",
    "shape_id",
    "title",
    "description",
    "resolved",
    "intentional"
  ]);

  // SUMMARY ROW (counts + deltas)
  const currentCounts = renderScoringSummary(lastScan) || getLastScanCounts();
  const prevCounts = getLastScanCounts();
  if (currentCounts) {
    const fmt = (n) => (n === 0 ? '0' : (n > 0 ? `+${n}` : `${n}`));
    const dTotal = prevCounts?.open ? (currentCounts.open.total - (prevCounts.open.total || 0)) : null;
    const dC = prevCounts?.open ? (currentCounts.open.critical - (prevCounts.open.critical || 0)) : null;
    const dS = prevCounts?.open ? (currentCounts.open.serious - (prevCounts.open.serious || 0)) : null;
    const dM = prevCounts?.open ? (currentCounts.open.moderate - (prevCounts.open.moderate || 0)) : null;
    const dMi = prevCounts?.open ? (currentCounts.open.minor - (prevCounts.open.minor || 0)) : null;
    const pass = (currentCounts.open.critical === 0 && currentCounts.open.serious === 0);
    const summaryDesc = `Open Total=${currentCounts.open.total}; Resolved=${currentCounts.resolved.total}; Critical=${currentCounts.open.critical}; Serious=${currentCounts.open.serious}; Moderate=${currentCounts.open.moderate}; Minor=${currentCounts.open.minor}; Open After Filters=${currentCounts.filteredOpen}; Gate=${pass ? 'PASS' : 'FAIL'}; Delta(Open)=${dTotal===null?'—':fmt(dTotal)} (C ${dC===null?'—':fmt(dC)}, S ${dS===null?'—':fmt(dS)}, M ${dM===null?'—':fmt(dM)}, Mi ${dMi===null?'—':fmt(dMi)})`;
    rows.push([lastScan.time, "SUMMARY", "", "", "", "Scoring Summary", summaryDesc, "", ""]);
  }

  const bag = getSettingsBag();
  const resolved = bag.resolved || {};
  const intentional = bag.intentional || {};

  for (const issue of (lastScan.issues || [])) {
    const sev = normalizeSeverity(issue.severity);
    const key = issueKey(issue);
    rows.push([
      lastScan.time,
      issue.check || "",
      sev,
      issue.slideNum || "",
      issue.shapeId || "",
      issue.title || "",
      issue.description || "",
      resolved[key] ? "true" : "false",
      intentional[key] ? "true" : "false"
    ]);
  }

  const csvText = rows.map(r => r.map(csvEscape).join(",")).join("\n");
  const blob = new Blob([csvText], { type:"text/csv;charset=utf-8" });
  downloadBlob(blob, `lcm-ppt-a11y-scan-${stamp()}.csv`);
}

function csvEscape(val) {
  const s = String(val ?? "");
  if (/[",\n\r]/.test(s)) return `"${s.replaceAll('"','""')}"`;
  return s;
}

function stamp() {
  return new Date().toISOString().replace(/[:.]/g,"-");
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function escapeHtml(str) {
  return String(str ?? "").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;")
    .replaceAll('"',"&quot;").replaceAll("'","&#39;");
}

/* ---------------------------
   PowerPoint helpers
---------------------------- */

async function getTotalSlideCount() {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    return slides.items.length || 1;
  });
}

async function getSlidesInRange(scanConfig) {
  const total = await getTotalSlideCount();
  if (scanConfig.mode !== "range") return { from:1, to:total, total };
  let from = Math.max(1, scanConfig.fromSlide||1);
  let to = Math.min(total, scanConfig.toSlide||total);
  if (from>to) [from,to]=[to,from];
  return { from, to, total };
}

async function selectSlide(slideNum) {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();
    const idx = Math.max(0, Math.min(slides.items.length-1, slideNum-1));
    const slide = slides.items[idx];
    if (slide && slide.select) slide.select(); // best-effort
    await context.sync();
  });
}

function boundsOverlap(a,b){ return !(a.right<=b.left||a.left>=b.right||a.bottom<=b.top||a.top>=b.bottom); }

/* ---------------------------
   Checks (best-effort)
---------------------------- */

async function checkSlideTitles(scanConfig) {
  const range = await getSlidesInRange(scanConfig);
  return PowerPoint.run(async (context) => {
    const slides=context.presentation.slides; slides.load("items"); await context.sync();
    const issues=[];
    for (let i=range.from-1;i<=range.to-1;i++){
      const shapes = slides.items[i].shapes;
      shapes.load("items/textFrame/textRange/text"); await context.sync();
      let title=null;
      for (const sh of shapes.items){
        const t = sh.textFrame?.textRange?.text ? sh.textFrame.textRange.text.trim() : "";
        if (t){ title=t; break; }
      }
      if (!title) issues.push({slideNum:i+1,title:"Missing slide title",description:"Add a title so screen readers can identify the slide.", severity:"serious"});
    }
    return { success:issues.length===0, message:issues.length?`Found ${issues.length} slide(s) missing a clear title.`:`All slides in range (${range.from}-${range.to}) appear to have titles.`, details:issues };
  });
}

async function checkDuplicateTitles(scanConfig){
  const range = await getSlidesInRange(scanConfig);
  return PowerPoint.run(async (context)=>{
    const slides=context.presentation.slides; slides.load("items"); await context.sync();
    const titles=[];
    for (let i=range.from-1;i<=range.to-1;i++){
      const shapes=slides.items[i].shapes;
      shapes.load("items/textFrame/textRange/text"); await context.sync();
      let t0="";
      for (const sh of shapes.items){ const t=sh.textFrame?.textRange?.text?sh.textFrame.textRange.text.trim():""; if (t){t0=t;break;} }
      titles.push({slideNum:i+1,title:t0});
    }
    const map=new Map();
    for (const t of titles){
      const k=t.title.toLowerCase().trim(); if(!k) continue;
      if(!map.has(k)) map.set(k,[]);
      map.get(k).push(t.slideNum);
    }
    const dupes=[];
    for (const [k,nums] of map.entries()){
      if(nums.length>1) dupes.push({
        slideNum:nums[0],
        title:"Duplicate slide title",
        description:`The title "${k}" is used on slides: ${nums.join(", ")}. This may be OK (continued sections), but consider making titles unique.`,
        extraKey:`title:${k}`,
        severity:"minor"
      });
    }
    return { success:true, message:dupes.length?`Found ${dupes.length} duplicated title group(s) (warning).`:"No duplicate titles detected in the selected range.", details:dupes };
  });
}

async function checkEmptySlides(scanConfig){
  const range = await getSlidesInRange(scanConfig);
  return PowerPoint.run(async (context)=>{
    const slides=context.presentation.slides; slides.load("items"); await context.sync();
    const issues=[]; let checked=0;
    for (let i=range.from-1;i<=range.to-1;i++){
      checked++;
      const shapes=slides.items[i].shapes;
      shapes.load("items/type,items/textFrame/textRange/text"); await context.sync();
      if(!shapes.items.length){ issues.push({slideNum:i+1,title:"Empty slide",description:"No elements found on the slide.", severity:"serious"}); continue; }
      const anyText = shapes.items.some(sh => (sh.textFrame?.textRange?.text||"").trim().length>0);
      const anyNonText = shapes.items.some(sh => String(sh.type||"").toLowerCase() && !String(sh.type||"").toLowerCase().includes("text"));
      if(!anyText && !anyNonText) issues.push({slideNum:i+1,title:"Empty slide",description:"Only empty text placeholders were found.", severity:"serious"});
    }
    return { success:issues.length===0, message:issues.length?`Found ${issues.length} empty slide(s).`:`Scanned ${checked} slide(s) - none appear empty.`, details:issues };
  });
}

async function checkTextSize(scanConfig){
  const range = await getSlidesInRange(scanConfig);
  const MIN_PT=12;
  return PowerPoint.run(async (context)=>{
    const slides=context.presentation.slides; slides.load("items"); await context.sync();
    const issues=[]; let textBlocks=0; let capabilityMissing=false;
    for (let i=range.from-1;i<=range.to-1;i++){
      const shapes=slides.items[i].shapes;
      shapes.load("items/id,items/textFrame/textRange/text,items/textFrame/textRange/font/size"); await context.sync();
      for (const sh of shapes.items){
        const tr=sh.textFrame?.textRange; const text=(tr?.text||"").trim(); if(!text) continue;
        textBlocks++;
        const size=(typeof tr?.font?.size==="number")?tr.font.size:null;
        if(size===null){ capabilityMissing=true; continue; }
        if(size>0 && size<MIN_PT){
          issues.push({slideNum:i+1,shapeId:sh.id,title:"Small text size",description:`Text appears to be ${Math.round(size)}pt. Consider increasing to at least ${MIN_PT}pt.`, severity:"moderate"});
        }
      }
    }
    if (capabilityMissing && issues.length===0) {
      return { success:true, skipped:true, message:"", details:[] };
    }
    const msg = issues.length?`Found ${issues.length} text element(s) below ${MIN_PT}pt.`:(textBlocks?`Scanned ${textBlocks} text element(s) - no small text detected.`:"No text elements found in the selected range.");
    return { success:issues.length===0, message:msg, details:issues };
  });
}

async function checkTextFormatting(scanConfig){
  const range = await getSlidesInRange(scanConfig);
  return PowerPoint.run(async (context)=>{
    const slides=context.presentation.slides; slides.load("items"); await context.sync();
    const issues=[]; let blocks=0; let capabilityMissing=false;
    for (let i=range.from-1;i<=range.to-1;i++){
      const shapes=slides.items[i].shapes;
      shapes.load("items/id,items/textFrame/textRange/text,items/textFrame/textRange/font/bold,items/textFrame/textRange/font/italic,items/textFrame/textRange/font/underline"); await context.sync();
      for (const sh of shapes.items){
        const tr=sh.textFrame?.textRange; const text=(tr?.text||"").trim(); if(!text) continue;
        blocks++;
        const b=tr?.font?.bold; const it=tr?.font?.italic; const u=tr?.font?.underline;
        if (b===undefined || it===undefined || u===undefined) { capabilityMissing=true; continue; }
        if (b===true && it===true && u===true){
          issues.push({slideNum:i+1,shapeId:sh.id,title:"Excessive text styling",description:"This text appears bold + italic + underlined. Reduce styling and rely on structure instead.", severity:"minor"});
        }
      }
    }
    if (capabilityMissing && issues.length===0) {
      return { success:true, skipped:true, message:"", details:[] };
    }
    return { success:issues.length===0, message:issues.length?`Found ${issues.length} block(s) with excessive styling.`:(blocks?`Scanned ${blocks} text block(s) - no excessive styling detected.`:"No text elements found in the selected range."), details:issues };
  });
}

async function checkManualListFormatting(scanConfig){
  const range = await getSlidesInRange(scanConfig);
  const bulletRe=/^\s*(?:[-–—*•]|\d+[.)]|\w[.)])\s+/;
  return PowerPoint.run(async (context)=>{
    const slides=context.presentation.slides; slides.load("items"); await context.sync();
    const issues=[];
    for (let i=range.from-1;i<=range.to-1;i++){
      const shapes=slides.items[i].shapes;
      shapes.load("items/id,items/textFrame/textRange/text"); await context.sync();
      for (const sh of shapes.items){
        const text=sh.textFrame?.textRange?.text||""; if(!text.trim()) continue;
        const lines=text.split(/\r?\n/);
        let manual=0; for (const ln of lines){ if(bulletRe.test(ln)) manual++; }
        if(manual>=5) issues.push({slideNum:i+1,shapeId:sh.id,title:"Possible manual list formatting",description:`This block has ${manual} line(s) that look like manually typed bullets/numbering. Use PowerPoint’s list formatting for screen reader structure.`, severity:"moderate"});
      }
    }
    return { success:issues.length===0, message:issues.length?`Found ${issues.length} text block(s) with likely manual list formatting.`:"No large manual lists detected (best-effort).", details:issues };
  });
}

async function checkAltText(scanConfig){
  const range = await getSlidesInRange(scanConfig);
  return PowerPoint.run(async (context)=>{
    const slides=context.presentation.slides; slides.load("items"); await context.sync();
    const issues=[]; let visuals=0; let capabilityMissing=false;
    for (let i=range.from-1;i<=range.to-1;i++){
      const shapes=slides.items[i].shapes;
      shapes.load("items/id,items/type,items/altTextTitle,items/altTextDescription"); await context.sync();
      for (const sh of shapes.items){
        const type=String(sh.type||"").toLowerCase();
        const isVisual=type.includes("picture")||type.includes("graphic")||type.includes("image")||type.includes("media");
        if(isVisual) visuals++;
        if (sh.altTextTitle === undefined && sh.altTextDescription === undefined) { capabilityMissing=true; continue; }
        const alt=((sh.altTextDescription||sh.altTextTitle||"")+"").trim();
        if(isVisual && !alt) issues.push({slideNum:i+1,shapeId:sh.id,title:"Missing alt text",description:"A visual element is missing alt text. Add a short description that conveys meaning.", severity:"serious"});
      }
    }
    if (capabilityMissing && issues.length===0) {
      return { success:true, skipped:true, message:"", details:[] };
    }
    return { success:issues.length===0, message:issues.length?`Found ${issues.length} visual element(s) missing alt text.`:(visuals? "No missing alt text detected for visual elements (best-effort).":"No visual elements found (best-effort)."), details:issues };
  });
}

async function checkOverlappingElements(scanConfig){
  const range = await getSlidesInRange(scanConfig);
  return PowerPoint.run(async (context)=>{
    const slides=context.presentation.slides; slides.load("items"); await context.sync();
    const issues=[]; let capabilityMissing=false;
    for (let i=range.from-1;i<=range.to-1;i++){
      const shapes=slides.items[i].shapes;
      shapes.load("items/id,items/left,items/top,items/width,items/height"); await context.sync();
      const b = shapes.items.map(sh=>{
        const left=typeof sh.left==="number"?sh.left:null;
        const top=typeof sh.top==="number"?sh.top:null;
        const w=typeof sh.width==="number"?sh.width:null;
        const h=typeof sh.height==="number"?sh.height:null;
        if([left,top,w,h].some(v=>v===null)) { capabilityMissing=true; return null; }
        return {id:sh.id,left,top,width:w,height:h,right:left+w,bottom:top+h};
      }).filter(Boolean);
      let found=false;
      for (let a=0;a<b.length && !found;a++){
        for (let c=a+1;c<b.length;c++){
          if(boundsOverlap(b[a],b[c])){ found=true; break; }
        }
      }
      if(found) issues.push({slideNum:i+1,title:"Overlapping elements detected",description:"Elements overlap on this slide. This can cause confusing reading order. Review Selection Pane / Reading Order.", severity:"minor"});
    }
    if (capabilityMissing && issues.length===0) {
      return { success:true, skipped:true, message:"", details:[] };
    }
    return { success:issues.length===0, message:issues.length?`Found overlapping elements on ${issues.length} slide(s) (best-effort).`:"No overlapping elements detected (best-effort).", details:issues };
  });
}

async function checkVagueLinks(scanConfig){
  const range = await getSlidesInRange(scanConfig);
  const vague=new Set(["click here","here","learn more","more","this","link"]);
  return PowerPoint.run(async (context)=>{
    const slides=context.presentation.slides; slides.load("items"); await context.sync();
    const issues=[];
    for (let i=range.from-1;i<=range.to-1;i++){
      const shapes=slides.items[i].shapes;
      shapes.load("items/id,items/textFrame/textRange/text"); await context.sync();
      for (const sh of shapes.items){
        const text=sh.textFrame?.textRange?.text||""; if(!text) continue;
        for (const line of text.split(/\r?\n/).map(s=>s.trim()).filter(Boolean)){
          if(vague.has(line.toLowerCase())){
            issues.push({slideNum:i+1,shapeId:sh.id,title:"Potential vague link text",description:`Found "${line}". If this is a link, make it descriptive (e.g., "Download the report (PDF)").`, severity:"moderate"});
          }
        }
      }
    }
    return { success:issues.length===0, message:issues.length?`Found ${issues.length} potential vague link issue(s).`:"No obvious vague link text found (best-effort).", details:issues };
  });
}
