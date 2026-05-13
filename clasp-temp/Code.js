// ================================================================
// EQUIPMENT INVENTORY SYSTEM v5 — Google Apps Script Backend
// ================================================================

const SHEET_ID      = "1iWTa93i4gWDZobBbY9DFga1EzR6vF-Ak-mHkoZmeH8o";
const GDRIVE_FOLDER = "1ajqJv_QDWi03OuLGK5hoh-4nsSjzwR0H";
const WAREHOUSE_EXT = "1XWaC-OaYFjwu1ZuqH_X6wabYMQdbVplBQJQw1M2rqHs";
// GROQ_KEY is in config.gs
const WAREHOUSE_TAB = "Mirroring data";

const ss = SpreadsheetApp.openById(SHEET_ID);

const SH = {
  users     : "Users",
  equipment : "Equipment",
  parts     : "SpareParts",
  wo        : "WorkOrders",
  partsUsed : "PartsUsed",
  history   : "History",
  warehouse : "Warehouse",
};

// ── Response ──────────────────────────────────────────────────
function jsonRes(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ROUTER GET ────────────────────────────────────────────────
function doGet(e) {
  const a = e.parameter.action;
  try {
    if (a === "getEquipment")  return jsonRes(getEquipment());
    if (a === "getParts")      return jsonRes(getParts());
    if (a === "getWorkOrders") return jsonRes(getWorkOrders());
    if (a === "getHistory")    return jsonRes(getHistory(e.parameter.equipmentId));
    if (a === "getDashboard")  return jsonRes(getDashboard());
    if (a === "getWarehouse")  return jsonRes(getWarehouseMirror());
    if (a === "getPartsUsed")  return jsonRes(sheetToJSON(SH.partsUsed));
    if (a === "getMechanics")  return jsonRes(getMechanics());
    return jsonRes({ error: "Unknown GET action" });
  } catch(err) { return jsonRes({ error: err.message }); }
}

// ── ROUTER POST ───────────────────────────────────────────────
function doPost(e) {
  const p = JSON.parse(e.postData.contents);
  const a = p.action;
  try {
    if (a === "login")             return jsonRes(login(p.data));
    if (a === "getUsers")          return jsonRes(getUsers());
    if (a === "addUser")           return jsonRes(addUser(p.data));
    if (a === "editUser")          return jsonRes(editUser(p.data));
    if (a === "deleteUser")        return jsonRes(deleteRow(SH.users, "Username", p.data.id));
    if (a === "uploadPhoto")       return jsonRes(uploadPhoto(p.data));
    if (a === "deletePhoto")       return jsonRes(deletePhoto(p.data));
    if (a === "addEquipment")      return jsonRes(addEquipment(p.data));
    if (a === "editEquipment")     return jsonRes(editEquipment(p.data));
    if (a === "deleteEquipment")   return jsonRes(deleteRow(SH.equipment, "EquipmentID", p.data.id));
    if (a === "addPart")           return jsonRes(addPart(p.data));
    if (a === "editPart")          return jsonRes(editPart(p.data));
    if (a === "deletePart")        return jsonRes(deleteRow(SH.parts, "PartID", p.data.id));
    if (a === "addWorkOrder")      return jsonRes(addWorkOrder(p.data));
    if (a === "editWorkOrder")     return jsonRes(editWorkOrder(p.data));
    if (a === "deleteWorkOrder")   return jsonRes(deleteRow(SH.wo, "WOID", p.data.id));
    if (a === "updateWOStatus")    return jsonRes(updateWOStatus(p.data));
    if (a === "confirmWODone")     return jsonRes(confirmWODone(p.data));
    if (a === "addPartsUsed")      return jsonRes(addPartsUsed(p.data));
    if (a === "ai")                return jsonRes(aiHandler(p.data));
    return jsonRes({ error: "Unknown POST action" });
  } catch(err) { return jsonRes({ error: err.message }); }
}

// ── HELPERS ───────────────────────────────────────────────────
function sheetToJSON(sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  const keys = rows[0];
  return rows.slice(1).map(r => {
    const obj = {};
    keys.forEach((k, i) => {
      obj[k] = (r[i] instanceof Date)
        ? Utilities.formatDate(r[i], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
        : r[i];
    });
    return obj;
  });
}

function genId(p) { return p + "-" + new Date().getTime(); }
function nowStr() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); }

function findRow(sheetName, colName, val) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return null;
  const rows = sh.getDataRange().getValues();
  const col  = rows[0].indexOf(colName);
  if (col < 0) return null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][col]) === String(val))
      return { sh, rows, ri: i + 1, col, hdr: rows[0], row: rows[i] };
  }
  return null;
}

function setCell(sheetName, colName, val, id, idCol) {
  const f = findRow(sheetName, idCol, id);
  if (!f) return false;
  const c = f.hdr.indexOf(colName);
  if (c < 0) return false;
  f.sh.getRange(f.ri, c + 1).setValue(val);
  return true;
}

function deleteRow(sheetName, colName, val) {
  const f = findRow(sheetName, colName, val);
  if (!f) return { error: "Record not found" };
  f.sh.deleteRow(f.ri);
  return { success: true };
}

// ── PHOTO UPLOAD ──────────────────────────────────────────────
function uploadPhoto(d) {
  try {
    const folder  = DriveApp.getFolderById(GDRIVE_FOLDER);
    const decoded = Utilities.base64Decode(d.base64);
    const blob    = Utilities.newBlob(decoded, d.mimeType, d.filename);
    const file    = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId  = file.getId();
    // Use thumbnail URL — renders reliably in <img> tags
    const viewUrl = "https://lh3.googleusercontent.com/d/" + fileId;
    return { success: true, url: viewUrl, fileId };
  } catch(err) {
    return { error: "Upload failed: " + err.message };
  }
}

// Delete old photo from Drive when replacing
function deletePhoto(d) {
  try {
    if (!d.fileId) return { success: true };
    DriveApp.getFileById(d.fileId).setTrashed(true);
    return { success: true };
  } catch(err) {
    return { success: true }; // non-fatal
  }
}

// Extract fileId from known URL formats
function extractFileId(url) {
  if (!url) return null;
  // lh3.googleusercontent.com/d/FILE_ID
  let m = url.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  // drive.google.com/uc?id=FILE_ID
  m = url.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  return null;
}

// ── AUTH ──────────────────────────────────────────────────────
function login(d) {
  const users = sheetToJSON(SH.users);
  const u = users.find(x =>
    String(x.Username) === String(d.username) &&
    String(x.Password) === String(d.password) &&
    (x.Active === true || String(x.Active).toUpperCase() === "TRUE")
  );
  if (!u) return { error: "Invalid username or password" };
  return { success:true, username:u.Username, role:u.Role, fullName:u.FullName, photo:u.Photo||"" };
}

function getMechanics() {
  return sheetToJSON(SH.users)
    .filter(u => u.Role === "mechanic" || u.Role === "supervisor")
    .map(u => ({ username: u.Username, fullName: u.FullName }));
}

function getUsers() {
  return sheetToJSON(SH.users).map(u => ({
    Username: u.Username, FullName: u.FullName, Role: u.Role,
    Active: u.Active, Photo: u.Photo || "", CreatedAt: u.CreatedAt
  }));
}

function addUser(d) {
  ss.getSheetByName(SH.users)
    .appendRow([d.username, d.password, d.fullName, d.role, true, d.photo||"", nowStr()]);
  return { success: true };
}

function editUser(d) {
  const f = findRow(SH.users, "Username", d.username);
  if (!f) return { error: "User not found" };
  const map = { Password:d.password, FullName:d.fullName, Role:d.role, Active:d.active, Photo:d.photo };
  Object.entries(map).forEach(([k, v]) => {
    if (v !== undefined && v !== null && v !== "")
      setCell(SH.users, k, v, d.username, "Username");
  });
  return { success: true };
}

// ── EQUIPMENT ─────────────────────────────────────────────────
function getEquipment() { return sheetToJSON(SH.equipment); }

function addEquipment(d) {
  const id = genId("EQ");
  ss.getSheetByName(SH.equipment).appendRow([
    id, d.name, d.type, d.brand, d.model, d.serialNo, d.year,
    d.status||"Active", d.location||"", d.specs||"",
    d.imgFront||"", d.imgSide||"", d.imgIso||"", nowStr()
  ]);
  return { success: true, id };
}

function editEquipment(d) {
  const f = findRow(SH.equipment, "EquipmentID", d.id);
  if (!f) return { error: "Not found" };
  const map = { Name:d.name, Type:d.type, Brand:d.brand, Model:d.model, SerialNo:d.serialNo,
    Year:d.year, Status:d.status, Location:d.location, Specs:d.specs,
    ImgFront:d.imgFront, ImgSide:d.imgSide, ImgIso:d.imgIso };
  Object.entries(map).forEach(([k,v]) => {
    if (v !== undefined) setCell(SH.equipment, k, v, d.id, "EquipmentID");
  });
  return { success: true };
}

// ── SPARE PARTS ───────────────────────────────────────────────
function getParts() { return sheetToJSON(SH.parts); }

function addPart(d) {
  const id = genId("PT");
  ss.getSheetByName(SH.parts).appendRow([
    id, d.name, d.partNo, d.equipmentId||"General",
    d.unit, d.minStock||0, d.specs||"",
    d.imgFront||"", d.imgSide||"", d.imgIso||"", nowStr()
  ]);
  return { success: true, id };
}

function editPart(d) {
  const f = findRow(SH.parts, "PartID", d.id);
  if (!f) return { error: "Not found" };
  const map = { Name:d.name, PartNo:d.partNo, EquipmentID:d.equipmentId,
    Unit:d.unit, MinStock:d.minStock, Specs:d.specs,
    ImgFront:d.imgFront, ImgSide:d.imgSide, ImgIso:d.imgIso };
  Object.entries(map).forEach(([k,v]) => {
    if (v !== undefined) setCell(SH.parts, k, v, d.id, "PartID");
  });
  return { success: true };
}

// ── WORK ORDERS ───────────────────────────────────────────────
function getWorkOrders() { return sheetToJSON(SH.wo); }

function addWorkOrder(d) {
  const id = genId("WO");
  let status = "Open";
  const neededParts = d.neededParts || "";
  if (neededParts) {
    const wh = getWarehouseMirror();
    const ptList = neededParts.split(",").map(s => s.trim().toLowerCase());
    const hasEmpty = ptList.some(pt =>
      wh.some(w => w.Material && w.Material.toLowerCase().includes(pt) && Number(w.TotalStock) <= 0)
    );
    if (hasEmpty) status = "Waiting for Sparepart";
  }
  ss.getSheetByName(SH.wo).appendRow([
    id, nowStr(), d.equipmentId, d.mechanic,
    d.collaborators||"", d.activityType, d.description,
    status, d.priority||"Normal",
    neededParts, d.neededPartsDesc||"", d.neededPartsImg||"",
    "", "", "", ""
  ]);
  logHistory(d.equipmentId, id, d.mechanic, d.activityType, d.description);
  syncEqStatus(d.equipmentId, status);
  return { success: true, id };
}

function editWorkOrder(d) {
  const f = findRow(SH.wo, "WOID", d.id);
  if (!f) return { error: "Not found" };
  const map = { EquipmentID:d.equipmentId, Mechanic:d.mechanic, Collaborators:d.collaborators,
    ActivityType:d.activityType, Description:d.description, Priority:d.priority,
    NeededParts:d.neededParts, NeededPartsDesc:d.neededPartsDesc, NeededPartsImg:d.neededPartsImg };
  Object.entries(map).forEach(([k,v]) => {
    if (v !== undefined) setCell(SH.wo, k, v, d.id, "WOID");
  });
  return { success: true };
}

function updateWOStatus(d) {
  const f = findRow(SH.wo, "WOID", d.woId);
  if (!f) return { error: "WO not found" };
  const hdr = f.hdr, ri = f.ri, sh = f.sh;
  sh.getRange(ri, hdr.indexOf("Status") + 1).setValue(d.status);
  if (d.status === "In Progress") {
    sh.getRange(ri, hdr.indexOf("StartTime") + 1).setValue(nowStr());
    syncEqStatus(f.row[hdr.indexOf("EquipmentID")], "Maintenance");
  } else if (d.status === "Pending Confirmation") {
    sh.getRange(ri, hdr.indexOf("DoneTime") + 1).setValue(nowStr());
  } else if (d.status === "Waiting for Sparepart") {
    syncEqStatus(f.row[hdr.indexOf("EquipmentID")], "Waiting for Sparepart");
  } else if (d.status === "Cancelled") {
    syncEqStatus(f.row[hdr.indexOf("EquipmentID")], "Active");
  }
  return { success: true };
}

function confirmWODone(d) {
  const f = findRow(SH.wo, "WOID", d.woId);
  if (!f) return { error: "WO not found" };
  const hdr = f.hdr, ri = f.ri, sh = f.sh;
  sh.getRange(ri, hdr.indexOf("Status") + 1).setValue("Done");
  sh.getRange(ri, hdr.indexOf("ConfirmedBy") + 1).setValue(d.confirmedBy);
  sh.getRange(ri, hdr.indexOf("ConfirmedAt") + 1).setValue(nowStr());
  if (!f.row[hdr.indexOf("DoneTime")] || f.row[hdr.indexOf("DoneTime")] === "")
    sh.getRange(ri, hdr.indexOf("DoneTime") + 1).setValue(nowStr());
  syncEqStatus(f.row[hdr.indexOf("EquipmentID")], "Active");
  return { success: true };
}

function syncEqStatus(equipmentId, woStatus) {
  const map = {
    "Open":"Maintenance", "In Progress":"Maintenance",
    "Pending Confirmation":"Maintenance",
    "Waiting for Sparepart":"Waiting for Sparepart",
    "Done":"Active", "Cancelled":"Active",
    "Maintenance":"Maintenance", "Active":"Active"
  };
  setCell(SH.equipment, "Status", map[woStatus]||woStatus, equipmentId, "EquipmentID");
}

function addPartsUsed(d) {
  ss.getSheetByName(SH.partsUsed)
    .appendRow([genId("PU"), d.woId, d.partId, d.qty, nowStr()]);
  return { success: true };
}

// ── HISTORY ───────────────────────────────────────────────────
function logHistory(eqId, woId, mechanic, activity, desc) {
  ss.getSheetByName(SH.history)
    .appendRow([genId("HX"), eqId, woId, mechanic, activity, desc, nowStr()]);
}

function getHistory(eqId) {
  const all = sheetToJSON(SH.history);
  return eqId ? all.filter(r => r.EquipmentID === eqId) : all;
}

// ── WAREHOUSE MIRROR ──────────────────────────────────────────
// External sheet MUST be shared as Viewer with the Google account running this script.
// Expected columns in external sheet: Material, Unit, TotalStock  (exact spelling)
function getWarehouseMirror() {
  try {
    const ext  = SpreadsheetApp.openById(WAREHOUSE_EXT);
    const sh   = ext.getSheetByName(WAREHOUSE_TAB);
    if (!sh) return [{ _error: "Tab '" + WAREHOUSE_TAB + "' not found in external sheet." }];
    const rows = sh.getDataRange().getValues();
    if (rows.length < 2) return [];

    // Normalize header names (trim whitespace)
    const rawKeys = rows[0].map(k => String(k).trim());

    // Find column indices flexibly
    const findCol = (...names) => {
      for (const n of names) {
        const i = rawKeys.findIndex(k => k.toLowerCase() === n.toLowerCase());
        if (i >= 0) return i;
      }
      return -1;
    };

    const matCol   = findCol("Material","Nama Material","Name","Item","Nama");
    const unitCol  = findCol("Unit","Satuan","UOM");
    const stockCol = findCol("TotalStock","Total Stock","Stock","Stok","Qty","Quantity","Jumlah");

    if (matCol < 0) return [{ _error: "Column 'Material' not found. Headers found: " + rawKeys.join(", ") }];

    return rows.slice(1)
      .filter(r => r.some(c => c !== "" && c !== null))
      .map(r => ({
        Material   : r[matCol]   !== undefined ? String(r[matCol])   : "",
        Unit       : unitCol  >= 0 ? String(r[unitCol])  : "—",
        TotalStock : stockCol >= 0 ? r[stockCol] : 0,
      }))
      .filter(r => r.Material !== "");
  } catch(err) {
    return [{ _error: "Cannot access external sheet: " + err.message }];
  }
}

// ── DASHBOARD ─────────────────────────────────────────────────
function getDashboard() {
  const EQ = sheetToJSON(SH.equipment);
  const WO = sheetToJSON(SH.wo);
  const WH = getWarehouseMirror();

  const now2  = new Date();
  const start = new Date(now2.getFullYear(), now2.getMonth(), 1);
  const thisMonthWO = WO.filter(w => w.Date && new Date(w.Date) >= start);

  // Best mechanic
  const mechStats = {};
  thisMonthWO.forEach(w => {
    const names = [w.Mechanic, ...(w.Collaborators ? w.Collaborators.split(",") : [])]
      .map(n => String(n).trim()).filter(Boolean);
    names.forEach(n => {
      if (!mechStats[n]) mechStats[n] = { done:0, totalMin:0, timed:0, total:0 };
      mechStats[n].total++;
      if (String(w.Status) === "Done") {
        mechStats[n].done++;
        if (w.StartTime && w.DoneTime) {
          const mins = (new Date(w.DoneTime) - new Date(w.StartTime)) / 60000;
          if (mins > 0) { mechStats[n].totalMin += mins; mechStats[n].timed++; }
        }
      }
    });
  });
  const bestMechanic = Object.entries(mechStats)
    .map(([name, s]) => ({
      name, done:s.done, total:s.total,
      avgMin : s.timed ? Math.round(s.totalMin / s.timed) : 0,
      score  : s.done * 10 - (s.timed ? Math.round(s.totalMin / s.timed) : 0)
    }))
    .sort((a, b) => b.score - a.score || b.done - a.done)
    .slice(0, 5);

  // Top needed parts (normalized)
  const ptCount = {}, ptNorm = {};
  thisMonthWO.forEach(w => {
    if (!w.NeededParts) return;
    String(w.NeededParts).split(",").forEach(p => {
      const raw = p.trim(); if (!raw) return;
      const norm = raw.toLowerCase().replace(/\s+/g," ");
      ptCount[norm] = (ptCount[norm]||0) + 1;
      if (!ptNorm[norm]) ptNorm[norm] = raw;
    });
  });
  const topParts = Object.entries(ptCount)
    .sort((a,b) => b[1]-a[1]).slice(0,5)
    .map(([norm, count]) => ({ name: ptNorm[norm], count }));

  // Equipment stats
  const eqWO = {};
  WO.forEach(w => {
    if (!eqWO[w.EquipmentID]) eqWO[w.EquipmentID] = { total:0, done:0 };
    eqWO[w.EquipmentID].total++;
    if (String(w.Status) === "Done") eqWO[w.EquipmentID].done++;
  });
  const worstEq = EQ
    .map(e => ({ id:e.EquipmentID, name:e.Name, status:e.Status, woCount:(eqWO[e.EquipmentID]||{}).total||0 }))
    .sort((a,b) => b.woCount - a.woCount).slice(0,5);
  const bestEq = EQ
    .filter(e => e.Status === "Active")
    .map(e => ({ name:e.Name, woCount:(eqWO[e.EquipmentID]||{}).total||0 }))
    .sort((a,b) => a.woCount - b.woCount).slice(0,5);

  // WO Trend last 7 days
  const tz = Session.getScriptTimeZone();
  const trend = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(); d.setDate(d.getDate() - i);
    const ds = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    trend.push({
      date  : ds,
      total : WO.filter(w => w.Date && String(w.Date).slice(0,10) === ds).length,
      done  : WO.filter(w => w.DoneTime && String(w.DoneTime).slice(0,10) === ds).length
    });
  }

  const whArr = Array.isArray(WH) ? WH : [];
  const lowWH = whArr.filter(w => !w._error && Number(w.TotalStock) <= 5)
    .map(w => ({ name:w.Material, stock:w.TotalStock, unit:w.Unit }));

  return {
    totalEquipment  : EQ.length,
    activeEquipment : EQ.filter(e => e.Status === "Active").length,
    maintenance     : EQ.filter(e => e.Status === "Maintenance").length,
    waitingSparepart: EQ.filter(e => e.Status === "Waiting for Sparepart").length,
    openWO          : WO.filter(w => w.Status === "Open").length,
    inProgressWO    : WO.filter(w => w.Status === "In Progress").length,
    pendingWO       : WO.filter(w => w.Status === "Pending Confirmation").length,
    doneWO          : WO.filter(w => w.Status === "Done").length,
    lowStock        : lowWH.length,
    bestMechanic, topParts, worstEq, bestEq, lowWH, trend,
    thisMonthTotal  : thisMonthWO.length,
    thisMonthDone   : thisMonthWO.filter(w => w.Status === "Done").length,
  };
}

// ── AI / GEMINI ───────────────────────────────────────────────
function aiHandler(d) {
  const mode = d.mode || "chat";
  const prompt = d.prompt || "";
  const context = d.context || "";
  let fullPrompt = "";

  if (mode === "suggestWO") {
    fullPrompt = `You are an equipment maintenance assistant. Based on the following work order information, provide:
1. A detailed work description (2-3 sentences, in Indonesian)
2. Suggested priority (Normal/High/Urgent)

Equipment: ${context}
Activity type: ${prompt}
Format: respond with JSON like {"description":"...","priority":"..."}`;

  } else if (mode === "diagnose") {
    fullPrompt = `You are an equipment diagnostic expert. Based on the following information, provide a diagnosis and recommended actions in Indonesian.

Equipment: ${context}
Symptoms/Issue: ${prompt}

Provide response as JSON: {"diagnosis":"...","possibleCauses":["...","..."],"recommendedActions":["...","..."]}`;

  } else {
    // chat mode
    fullPrompt = `Kamu adalah asisten AI support yang ramah, profesional, dan kadang lucu untuk aplikasi SINVENA (Sistem Inventarisasi Alat). 
Gunakan bahasa Indonesia yang santai, hangat, dan sopan seperti customer service professional.

PANDUAN KEPRIBADIAN:
- RAMAH: Gunakan sapaan seperti "Halo!", "Tentu!", "Dengan senang hati!" 
- PROFESIONAL: Jawab dengan informatif, terstruktur, dan jelas
- SUPPORTIF: Jika ada masalah, berikan solusi langkah demi langkah
- LUCU (sesekali): Selipkan jokes ringan atau emoji yang relevan, jangan berlebihan
- SANTUN: Selalu gunakan kata "silakan", "mohon", "terima kasih"
- Jika ditanya di luar konteks SINVENA, jawab dengan sopan bahwa kamu hanya bisa bantu soal inventory

ATURAN FORMAT JAWABAN:
- Gunakan nomor urut (1. 2. 3. ...) untuk setiap item
- SETIAP ITEM HARUS PADA BARIS YANG BERBEDA (gunakan enter/newline)
- Contoh format yang BENAR:
  1. Over Head Crane (OHC) -1 — Zhao Xing | Active | Plant spunpile
  2. Over Head Crane (OHC) -2 — Zhao Xing | Active | Plant spunpile
  3. Over Head Crane (OHC) -3 — Zhao Xing | Active | Plant spunpile
- Jangan pernah menggabungkan beberapa item dalam satu baris
- Gunakan tanda baca yang benar (huruf kapital, titik, koma)
- Jawab dengan ringkas dan padat, jangan bertele-tele
- Setiap list item cukup 1 baris saja

DATA APLIKASI SAAT INI:
${context}

PERTANYAAN USER:
${prompt}

Jawab dengan hangat dan membantu dalam bahasa Indonesia:`;
  }

  return aiGroq(fullPrompt);
}

function aiGroq(prompt) {
  try {
    const url = "https://api.groq.com/openai/v1/chat/completions";
    const payload = {
      model: "llama-3.1-8b-instant",
      messages: [{ role: "user", content: prompt }],
      temperature: 0.3,
      max_tokens: 1024
    };
    const options = {
      method: "post",
      headers: { Authorization: "Bearer " + GROQ_KEY },
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const res = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(res.getContentText());
    if (json.choices && json.choices[0] && json.choices[0].message) {
      const text = json.choices[0].message.content.trim();
      const jsonMatch = text.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        try { return JSON.parse(jsonMatch[0]); } catch(e) {}
      }
      return { success: true, text };
    }
    return { error: json.error?.message || "Empty response" };
  } catch(err) {
    return { error: err.message };
  }
}

// ── SETUP ─────────────────────────────────────────────────────
function setupSheets() {
  const headers = {
    Users      : ["Username","Password","FullName","Role","Active","Photo","CreatedAt"],
    Equipment  : ["EquipmentID","Name","Type","Brand","Model","SerialNo","Year","Status","Location","Specs","ImgFront","ImgSide","ImgIso","CreatedAt"],
    SpareParts : ["PartID","Name","PartNo","EquipmentID","Unit","MinStock","Specs","ImgFront","ImgSide","ImgIso","CreatedAt"],
    WorkOrders : ["WOID","Date","EquipmentID","Mechanic","Collaborators","ActivityType","Description","Status","Priority","NeededParts","NeededPartsDesc","NeededPartsImg","StartTime","DoneTime","ConfirmedBy","ConfirmedAt"],
    PartsUsed  : ["RecordID","WOID","PartID","QtyUsed","Date"],
    History    : ["HistoryID","EquipmentID","WOID","Mechanic","Activity","Description","Date"],
    Warehouse  : ["MaterialID","Material","Unit","TotalStock","UpdatedAt"],
  };
  Object.entries(headers).forEach(([name, cols]) => {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1,1,1,cols.length).setValues([cols])
      .setFontWeight("bold").setBackground("#1e3a5f").setFontColor("#fff");
    sh.setFrozenRows(1);
  });
  const us = ss.getSheetByName("Users");
  if (us.getLastRow() < 2) {
    [["admin","admin123","Administrator","manager",true,""],
     ["mechanic1","mech123","Fajar Herkuntarto","mechanic",true,""],
     ["mechanic2","mech123","Andre Pratama","mechanic",true,""],
     ["supervisor1","super123","Jane Supervisor","supervisor",true,""]]
    .forEach(r => us.appendRow([...r, new Date()]));
  }
  SpreadsheetApp.getUi().alert("✅ Setup complete!");
}