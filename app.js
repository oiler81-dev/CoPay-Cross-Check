// Copay Cross-Check (Seen-only)
// Eligibility (pVerify): Patient Name, DOB, Spec.Copay (In-Net), Location
// Patient Balance (AdvancedMD): Patient Name (First Last), Patient Birth Date, Patient Payments, Appointment User, Appointment Status

const els = {
  fileA: document.getElementById("fileA"),
  fileB: document.getElementById("fileB"),
  runBtn: document.getElementById("runBtn"),
  exportBtn: document.getElementById("exportBtn"),
  status: document.getElementById("status"),

  kpiPatients: document.getElementById("kpiPatients"),
  kpiCollectedCount: document.getElementById("kpiCollectedCount"),
  kpiPartialCount: document.getElementById("kpiPartialCount"),
  kpiNotCollectedCount: document.getElementById("kpiNotCollectedCount"),
  kpiExpectedTotal: document.getElementById("kpiExpectedTotal"),
  kpiCollectedTotal: document.getElementById("kpiCollectedTotal"),
  kpiNotCollectedTotal: document.getElementById("kpiNotCollectedTotal"),
  kpiAvgConfidence: document.getElementById("kpiAvgConfidence"),

  statusFilter: document.getElementById("statusFilter"),
  locationFilter: document.getElementById("locationFilter"),
  searchBox: document.getElementById("searchBox"),

  resultsBody: document.getElementById("resultsBody"),
  usersBody: document.getElementById("usersBody"),
};

let LAST_RESULTS = [];
let LAST_USER_ROWS = [];
let LAST_META = {};

els.runBtn.addEventListener("click", run);
els.exportBtn.addEventListener("click", exportResults);

function setStatus(msg){ els.status.textContent = msg || ""; }

function money(n){
  const v = Number.isFinite(n) ? n : 0;
  return "$" + v.toFixed(2);
}

function normalizeSpaces(s){
  return (s ?? "").toString().trim().replace(/\s+/g, " ");
}

function stripNonLetters(s){
  return normalizeSpaces(s).toLowerCase().replace(/[^a-z\s]/g, "").trim();
}

function parseNameParts(full){
  const raw = normalizeSpaces(full);
  if (!raw) return { first:"", last:"", full:"" };

  if (raw.includes(",")){
    const [last, rest] = raw.split(",", 2);
    const tokens = stripNonLetters(rest).split(" ").filter(Boolean);
    return {
      first: tokens[0] || "",
      last: stripNonLetters(last),
      full: stripNonLetters(tokens.join(" ") + " " + last),
    };
  }

  const tokens = stripNonLetters(raw).split(" ").filter(Boolean);
  if (tokens.length === 1) return { first: tokens[0], last:"", full: tokens[0] };

  return {
    first: tokens[0] || "",
    last: tokens[tokens.length - 1] || "",
    full: tokens.join(" "),
  };
}

function parseMoney(v){
  if (v === null || v === undefined) return 0;
  const s = v.toString().replace(/[^0-9.-]+/g, "");
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}

function toISODateLoose(v){
  if (v === null || v === undefined || v === "") return "";

  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)){
    return v.toISOString().slice(0,10);
  }

  // Excel serial number
  if (typeof v === "number"){
    const d = XLSX.SSF.parse_date_code(v);
    if (d && d.y && d.m && d.d){
      const mm = String(d.m).padStart(2, "0");
      const dd = String(d.d).padStart(2, "0");
      return `${d.y}-${mm}-${dd}`;
    }
  }

  const s = v.toString().trim();
  const d2 = new Date(s);
  if (!isNaN(d2)) return d2.toISOString().slice(0,10);

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m){
    const mm = String(m[1]).padStart(2, "0");
    const dd = String(m[2]).padStart(2, "0");
    let yy = m[3];
    if (yy.length === 2) yy = "20" + yy;
    return `${yy}-${mm}-${dd}`;
  }

  return "";
}

function levenshtein(a,b){
  a = a || ""; b = b || "";
  if (a === b) return 0;
  const alen = a.length, blen = b.length;
  if (!alen) return blen;
  if (!blen) return alen;

  const v0 = new Array(blen+1);
  const v1 = new Array(blen+1);
  for (let i=0;i<=blen;i++) v0[i]=i;

  for (let i=0;i<alen;i++){
    v1[0]=i+1;
    for (let j=0;j<blen;j++){
      const cost = a[i] === b[j] ? 0 : 1;
      v1[j+1] = Math.min(v1[j]+1, v0[j+1]+1, v0[j]+cost);
    }
    for (let j=0;j<=blen;j++) v0[j]=v1[j];
  }
  return v1[blen];
}

function similarityScore(a,b){
  a = stripNonLetters(a); b = stripNonLetters(b);
  if (!a || !b) return 0;
  const dist = levenshtein(a,b);
  const maxLen = Math.max(a.length,b.length) || 1;
  const sim = 1 - dist/maxLen;
  return Math.max(0, Math.min(1, sim));
}

function matchConfidence(eligName, balName){
  const e = parseNameParts(eligName);
  const b = parseNameParts(balName);
  const firstSim = similarityScore(e.first, b.first);
  const lastSim  = similarityScore(e.last,  b.last);
  const fullSim  = similarityScore(e.full,  b.full);

  const score = (lastSim*0.55) + (firstSim*0.30) + (fullSim*0.15);
  return Math.round(score * 100);
}

function bestNameMatchRow(targetName, rows, nameField){
  let best = null;
  let bestScore = -1;
  let notes = "";

  for (const r of rows){
    const cand = r[nameField] ?? "";
    const conf = matchConfidence(targetName, cand);
    if (conf > bestScore){
      bestScore = conf;
      best = r;
    }
  }

  if (!best) return { row:null, confidence:0, notes:"No candidate rows" };

  if (bestScore >= 92) notes = "Strong match";
  else if (bestScore >= 82) notes = "Good match";
  else if (bestScore >= 70) notes = "Weak match";
  else notes = "Very weak match";

  return { row: best, confidence: bestScore, notes };
}

/** Seen-only logic:
 *  We treat these as "seen" unless they look like a cancel/no-show.
 *  Adjust keywords here if AdvancedMD uses different wording.
 */
function isSeenStatus(statusRaw){
  const s = normalizeSpaces(statusRaw).toLowerCase();

  if (!s) return false;

  // hard excludes first
  const bad = ["cancel", "canceled", "no show", "noshow", "resched", "reschedule", "left without", "lwbs"];
  if (bad.some(k => s.includes(k))) return false;

  // seen-ish includes
  const good = ["seen", "complete", "completed", "checked out", "check out", "arrived", "in progress", "roomed"];
  if (good.some(k => s.includes(k))) return true;

  // if it's unknown text, be conservative (don’t include)
  return false;
}

async function run(){
  const fileA = els.fileA.files[0];
  const fileB = els.fileB.files[0];
  if (!fileA || !fileB){
    alert("Upload both reports first.");
    return;
  }

  setStatus("Reading files...");
  els.exportBtn.disabled = true;
  LAST_RESULTS = [];
  LAST_USER_ROWS = [];
  LAST_META = {};

  try{
    const [eligRows, balRows] = await Promise.all([readAny(fileA), readAny(fileB)]);

    const elig = normalizeEligibility(eligRows);
    const balAll = normalizeBalance(balRows);

    // SEEN ONLY
    const balSeen = balAll.filter(r => r.isSeen);

    setStatus(
      `Loaded: Eligibility ${elig.length.toLocaleString()} copay rows, ` +
      `AMD ${balAll.length.toLocaleString()} rows (${balSeen.length.toLocaleString()} seen). Matching...`
    );

    const results = crossCheckSeenOnly(elig, balSeen);
    LAST_RESULTS = results;

    buildLocationOptions(results);
    renderAll();

    els.exportBtn.disabled = results.length === 0;
    setStatus(`Done. Results include only pVerify patients who were SEEN in AMD: ${results.length.toLocaleString()}.`);
  } catch (err){
    console.error(err);
    setStatus("Error. Check console for details.");
    alert("Something broke while reading/matching these files. Open DevTools Console for details.");
  }
}

function readAny(file){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try{
        const data = e.target.result;
        const wb = XLSX.read(data, { type:"array" });
        const ws = wb.Sheets[wb.SheetNames[0]];

        const aoa = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
        const headerIdx = detectHeaderRowIndex(aoa);
        const rows = aoaToObjects(aoa, headerIdx);
        resolve(rows);
      } catch(ex){
        reject(ex);
      }
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function detectHeaderRowIndex(aoa){
  const keys = [
    "patient name",
    "dob",
    "spec.copay",
    "location",
    "patient name (first last)",
    "patient birth date",
    "patient payments",
    "appointment user",
    "appointment status",
  ];

  let bestIdx = 0;
  let bestScore = -1;

  for (let i=0;i<Math.min(aoa.length, 35);i++){
    const row = aoa[i] || [];
    const joined = row.map(x => (x ?? "").toString().toLowerCase()).join(" | ");
    let score = 0;
    for (const k of keys) if (joined.includes(k)) score++;
    if (score > bestScore){
      bestScore = score;
      bestIdx = i;
    }
  }
  return bestIdx;
}

function aoaToObjects(aoa, headerIdx){
  const headerRow = (aoa[headerIdx] || []).map(h => normalizeSpaces(h));
  const rows = [];

  for (let r = headerIdx + 1; r < aoa.length; r++){
    const row = aoa[r] || [];
    if (row.every(v => (v ?? "").toString().trim() === "")) continue;

    const obj = {};
    for (let c = 0; c < headerRow.length; c++){
      const key = headerRow[c];
      if (!key) continue;
      obj[key] = row[c];
    }
    rows.push(obj);
  }
  return rows;
}

function normalizeEligibility(rows){
  const out = [];
  for (const r of rows){
    const name = r["Patient Name"];
    const dob  = r["DOB"];
    const copay = r["Spec.Copay (In-Net)"];
    const location = r["Location"];

    if (!normalizeSpaces(name)) continue;

    const expected = parseMoney(copay);
    if (!(expected > 0)) continue;

    out.push({
      patientName: normalizeSpaces(name),
      dobISO: toISODateLoose(dob),
      expected,
      location: normalizeSpaces(location) || "Unknown",
      raw: r
    });
  }
  return out;
}

function normalizeBalance(rows){
  const out = [];
  for (const r of rows){
    const name = r["Patient Name (First Last)"];
    const dob  = r["Patient Birth Date"];
    const payments = r["Patient Payments"];
    const user = r["Appointment User"];
    const apptStatus = r["Appointment Status"]; // NEW

    if (!normalizeSpaces(name)) continue;

    out.push({
      patientName: normalizeSpaces(name),
      dobISO: toISODateLoose(dob),
      collected: parseMoney(payments),
      user: normalizeSpaces(user) || "Unknown",
      appointmentStatus: normalizeSpaces(apptStatus) || "",
      isSeen: isSeenStatus(apptStatus),
      raw: r
    });
  }
  return out;
}

function crossCheckSeenOnly(elig, balSeen){
  // Index seen balance rows by DOB
  const byDob = new Map();
  for (const b of balSeen){
    if (!b.dobISO) continue;
    if (!byDob.has(b.dobISO)) byDob.set(b.dobISO, []);
    byDob.get(b.dobISO).push(b);
  }

  const results = [];

  for (const e of elig){
    const candidates = byDob.get(e.dobISO) || [];

    // IMPORTANT: If no seen AMD record exists for this DOB, discard (do not include).
    if (candidates.length === 0) continue;

    // Pick best name match among seen candidates
    const pick = bestNameMatchRow(e.patientName, candidates, "patientName");
    const matchedRow = pick.row;

    // If we have candidates but best match is still extremely weak, you can either discard or keep.
    // I’m keeping it, because staff can sort by confidence and investigate.
    const confidence = pick.confidence;
    const notes = pick.notes;

    const collected = matchedRow ? (matchedRow.collected || 0) : 0;
    const user = matchedRow ? (matchedRow.user || "Unknown") : "Unknown";
    const apptStatus = matchedRow ? (matchedRow.appointmentStatus || "") : "";

    let status = "Not Collected";
    if (collected >= e.expected && e.expected > 0) status = "Collected";
    else if (collected > 0 && collected < e.expected) status = "Partial";

    results.push({
      patient: e.patientName,
      dob: e.dobISO,
      location: e.location,
      expected: e.expected,
      collected,
      status,
      collectedBy: status === "Not Collected" ? "N/A" : user,
      confidence,
      matchNotes: `${notes}${apptStatus ? ` | Appt: ${apptStatus}` : ""}`,
      balancePatientName: matchedRow ? matchedRow.patientName : "",
    });
  }

  LAST_META = {
    generatedAt: new Date().toISOString(),
    eligRowsUsed: elig.length,
    balSeenRowsUsed: balSeen.length
  };

  return results;
}

function buildLocationOptions(results){
  const locs = new Set(results.map(r => r.location).filter(Boolean));
  const arr = Array.from(locs).sort((a,b) => a.localeCompare(b));
  els.locationFilter.innerHTML =
    `<option value="All">All locations</option>` +
    arr.map(l => `<option value="${escapeHtml(l)}">${escapeHtml(l)}</option>`).join("");
}

function renderAll(){
  renderKPIs(LAST_RESULTS);
  renderPatientTable(LAST_RESULTS);
  renderUserBreakdown(LAST_RESULTS);

  els.statusFilter.onchange = () => applyFilters();
  els.locationFilter.onchange = () => applyFilters();
  els.searchBox.oninput = () => applyFilters();
}

function applyFilters(){
  const status = els.statusFilter.value;
  const location = els.locationFilter.value;
  const q = normalizeSpaces(els.searchBox.value).toLowerCase();

  const filtered = LAST_RESULTS.filter(r => {
    if (status !== "All" && r.status !== status) return false;
    if (location !== "All" && r.location !== location) return false;
    if (q){
      const hay = `${r.patient} ${r.dob} ${r.location} ${r.collectedBy} ${r.status} ${r.matchNotes} ${r.balancePatientName}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  renderKPIs(filtered, LAST_RESULTS);
  renderPatientTable(filtered);
  renderUserBreakdown(filtered);
}

function renderKPIs(results, global=null){
  const patients = results.length;
  const collectedCount = results.filter(r => r.status === "Collected").length;
  const partialCount = results.filter(r => r.status === "Partial").length;
  const notCount = results.filter(r => r.status === "Not Collected").length;

  const expectedTotal = results.reduce((s,r) => s + (r.expected || 0), 0);
  const collectedTotal = results.reduce((s,r) => s + (r.collected || 0), 0);
  const notCollectedTotal = Math.max(0, expectedTotal - collectedTotal);

  const avgConfidence = results.length
    ? Math.round(results.reduce((s,r) => s + (r.confidence || 0), 0) / results.length)
    : 0;

  els.kpiPatients.textContent = patients.toLocaleString();
  els.kpiCollectedCount.textContent = collectedCount.toLocaleString();
  els.kpiPartialCount.textContent = partialCount.toLocaleString();
  els.kpiNotCollectedCount.textContent = notCount.toLocaleString();

  els.kpiExpectedTotal.textContent = money(expectedTotal);
  els.kpiCollectedTotal.textContent = money(collectedTotal);
  els.kpiNotCollectedTotal.textContent = money(notCollectedTotal);

  if (patients === 0 && global && global.length){
    const globalAvg = Math.round(global.reduce((s,r) => s + (r.confidence || 0), 0) / global.length);
    els.kpiAvgConfidence.textContent = `${globalAvg}`;
  } else {
    els.kpiAvgConfidence.textContent = `${avgConfidence}`;
  }
}

function renderPatientTable(results){
  els.resultsBody.innerHTML = "";

  const rowsHtml = results.map(r => {
    const rowClass =
      r.status === "Collected" ? "row-collected" :
      r.status === "Partial" ? "row-partial" :
      "row-not";

    const badgeClass =
      r.status === "Collected" ? "badge collected" :
      r.status === "Partial" ? "badge partial" :
      "badge not";

    return `
      <tr class="${rowClass}">
        <td>${escapeHtml(r.patient)}</td>
        <td>${escapeHtml(r.dob)}</td>
        <td>${escapeHtml(r.location)}</td>
        <td>${money(r.expected)}</td>
        <td>${money(r.collected)}</td>
        <td><span class="${badgeClass}">${escapeHtml(r.status)}</span></td>
        <td>${escapeHtml(r.collectedBy)}</td>
        <td>${escapeHtml(String(r.confidence))}</td>
        <td>${escapeHtml(r.matchNotes)}${r.balancePatientName ? ` (Bal: ${escapeHtml(r.balancePatientName)})` : ""}</td>
      </tr>
    `;
  }).join("");

  els.resultsBody.innerHTML = rowsHtml || `<tr><td colspan="9">No results.</td></tr>`;
}

function renderUserBreakdown(results){
  const map = new Map();

  for (const r of results){
    const user = normalizeSpaces(r.collectedBy) || "Unknown";

    if (!map.has(user)){
      map.set(user, {
        user,
        collectedCount: 0,
        partialCount: 0,
        notCount: 0,
        expectedTotal: 0,
        collectedTotal: 0,
        confidenceSum: 0,
        n: 0,
      });
    }

    const u = map.get(user);
    if (r.status === "Collected") u.collectedCount++;
    else if (r.status === "Partial") u.partialCount++;
    else u.notCount++;

    u.expectedTotal += (r.expected || 0);
    u.collectedTotal += (r.collected || 0);
    u.confidenceSum += (r.confidence || 0);
    u.n++;
  }

  const rows = Array.from(map.values())
    .map(u => ({
      ...u,
      notCollectedTotal: Math.max(0, u.expectedTotal - u.collectedTotal),
      avgConfidence: u.n ? Math.round(u.confidenceSum / u.n) : 0
    }))
    .sort((a,b) => (b.collectedTotal - a.collectedTotal));

  LAST_USER_ROWS = rows;

  els.usersBody.innerHTML = rows.map(u => `
    <tr>
      <td>${escapeHtml(u.user)}</td>
      <td>${u.collectedCount.toLocaleString()}</td>
      <td>${u.partialCount.toLocaleString()}</td>
      <td>${u.notCount.toLocaleString()}</td>
      <td>${money(u.expectedTotal)}</td>
      <td>${money(u.collectedTotal)}</td>
      <td>${money(u.notCollectedTotal)}</td>
      <td>${escapeHtml(String(u.avgConfidence))}</td>
    </tr>
  `).join("") || `<tr><td colspan="8">No users.</td></tr>`;
}

function exportResults(){
  if (!LAST_RESULTS.length) return;

  const wb = XLSX.utils.book_new();

  const resultsForSheet = LAST_RESULTS.map(r => ({
    Patient: r.patient,
    DOB: r.dob,
    Location: r.location,
    Expected: r.expected,
    Collected: r.collected,
    Status: r.status,
    "Collected By": r.collectedBy,
    "Match Confidence": r.confidence,
    "Match Notes": r.matchNotes,
    "Balance Patient Name": r.balancePatientName,
  }));
  const ws1 = XLSX.utils.json_to_sheet(resultsForSheet);
  XLSX.utils.book_append_sheet(wb, ws1, "Patient Results");

  const usersForSheet = LAST_USER_ROWS.map(u => ({
    User: u.user,
    "Collected (count)": u.collectedCount,
    "Partial (count)": u.partialCount,
    "Not Collected (count)": u.notCount,
    "Expected Total": u.expectedTotal,
    "Collected Total": u.collectedTotal,
    "Not Collected Total": Math.max(0, u.expectedTotal - u.collectedTotal),
    "Avg Confidence": u.avgConfidence,
  }));
  const ws2 = XLSX.utils.json_to_sheet(usersForSheet);
  XLSX.utils.book_append_sheet(wb, ws2, "Per-User Summary");

  const meta = [
    { Key: "GeneratedAt", Value: LAST_META.generatedAt || "" },
    { Key: "EligibilityRowsUsed", Value: LAST_META.eligRowsUsed || 0 },
    { Key: "BalanceSeenRowsUsed", Value: LAST_META.balSeenRowsUsed || 0 },
  ];
  const ws3 = XLSX.utils.json_to_sheet(meta);
  XLSX.utils.book_append_sheet(wb, ws3, "Meta");

  const stamp = new Date().toISOString().slice(0,19).replace(/[:T]/g,"-");
  XLSX.writeFile(wb, `copay-crosscheck-${stamp}.xlsx`);
}

function escapeHtml(s){
  return (s ?? "").toString()
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}
