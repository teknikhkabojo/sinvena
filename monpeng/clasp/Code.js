const SHEET_ID = "1-3x-tIulZNyMqNYtbHdgGSLEBh4mfrpiT1xH-et51SU";
const SS = SpreadsheetApp.openById(SHEET_ID);

const SH = {
  customers: "Customers",
  projects: "Projects",
  products: "Products",
  po: "PO",
};

function json(d) {
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const a = e.parameter.action;
  try {
    if (a === "getDashboard") return json(getDashboard());
    if (a === "getCustomers") return json(stj(SH.customers));
    if (a === "getProjects") return json(stj(SH.projects));
    if (a === "getProducts") return json(stj(SH.products));
    if (a === "getPO") return json(stj(SH.po));
    return json({ error: "Unknown action" });
  } catch (err) { return json({ error: err.message }); }
}

function doPost(e) {
  const p = JSON.parse(e.postData.contents);
  const a = p.action;
  try {
    if (a === "addCustomer") return json(addRec(SH.customers, ["Name", "Address", "Phone", "CreatedAt"], [p.name, p.address || "", p.phone || "", now()]));
    if (a === "editCustomer") return json(editRec(SH.customers, "ID", p.id, { Name: p.name, Address: p.address, Phone: p.phone }));
    if (a === "deleteCustomer") return json(delRec(SH.customers, "ID", p.id));
    if (a === "addProject") return json(addRec(SH.projects, ["CustomerID", "ProjectName", "Location", "CreatedAt"], [p.customerId, p.projectName, p.location || "", now()]));
    if (a === "editProject") return json(editRec(SH.projects, "ID", p.id, { CustomerID: p.customerId, ProjectName: p.projectName, Location: p.location }));
    if (a === "deleteProject") return json(delRec(SH.projects, "ID", p.id));
    if (a === "addProduct") return json(addRec(SH.products, ["ProductType", "Diameter", "Length", "Class", "Type", "Unit", "CreatedAt"], [p.productType, p.diameter, p.length, p.pClass, p.pType, p.unit, now()]));
    if (a === "editProduct") return json(editRec(SH.products, "ID", p.id, { ProductType: p.productType, Diameter: p.diameter, Length: p.length, Class: p.pClass, Type: p.pType, Unit: p.unit }));
    if (a === "deleteProduct") return json(delRec(SH.products, "ID", p.id));
    if (a === "addPO") return json(addPO(p));
    if (a === "editPO") return json(editPO(p));
    if (a === "deletePO") return json(delRec(SH.po, "ID", p.id));
    return json({ error: "Unknown action" });
  } catch (err) { return json({ error: err.message }); }
}

function stj(sn) { const sh = SS.getSheetByName(sn); if (!sh) return []; const r = sh.getDataRange().getValues(); if (r.length < 2) return []; const k = r[0]; return r.slice(1).map(row => { const o = {}; k.forEach((c, i) => { o[c] = (row[i] instanceof Date) ? Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : row[i]; }); return o; }); }

function genId(p) { return p + "-" + Date.now(); }
function now() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); }

function findRow(sn, cn, v) {
  const sh = SS.getSheetByName(sn); if (!sh) return null;
  const r = sh.getDataRange().getValues(); const ci = r[0].indexOf(cn); if (ci < 0) return null;
  for (let i = 1; i < r.length; i++) { if (String(r[i][ci]) === String(v)) return { sh, ri: i + 1, h: r[0], row: r[i] }; }
  return null;
}

function setCell(sn, cn, v, id) { const f = findRow(sn, "ID", id); if (!f) return false; const c = f.h.indexOf(cn); if (c < 0) return false; f.sh.getRange(f.ri, c + 1).setValue(v); return true; }

function delRec(sn, cn, v) { const f = findRow(sn, cn, v); if (!f) return { error: "Not found" }; f.sh.deleteRow(f.ri); return { success: true }; }

function addRec(sn, cols, vals) {
  const id = genId(sn.slice(0, 2).toUpperCase());
  SS.getSheetByName(sn).appendRow([id, ...vals]);
  return { success: true, id };
}

function editRec(sn, idCol, id, map) {
  Object.entries(map).forEach(([k, v]) => { if (v !== undefined) setCell(sn, k, v, id); });
  return { success: true };
}

function addPO(d) {
  const id = genId("PO");
  const total = (parseFloat(d.qty) || 0) * (parseFloat(d.unitPrice) || 0);
  SS.getSheetByName(SH.po).appendRow([id, d.customerId, d.projectId, d.poDate, d.productId, d.unitPrice, d.qty, total, d.status || "Open", d.notes || "", now()]);
  return { success: true, id };
}

function editPO(d) {
  const f = findRow(SH.po, "ID", d.id); if (!f) return { error: "Not found" };
  const total = (parseFloat(d.qty) || 0) * (parseFloat(d.unitPrice) || 0);
  const map = { CustomerID: d.customerId, ProjectID: d.projectId, PODate: d.poDate, ProductID: d.productId, UnitPrice: d.unitPrice, Qty: d.qty, Total: total, Status: d.status, Notes: d.notes };
  Object.entries(map).forEach(([k, v]) => { if (v !== undefined) setCell(SH.po, k, v, d.id); });
  return { success: true };
}

function getDashboard() {
  const PO = stj(SH.po);
  const totalPO = PO.length;
  const totalOpen = PO.filter(p => p.Status === "Open").length;
  const totalDone = PO.filter(p => p.Status === "Done").length;
  const totalValue = PO.reduce((s, p) => s + (parseFloat(p.Total) || 0), 0);
  const doneValue = PO.filter(p => p.Status === "Done").reduce((s, p) => s + (parseFloat(p.Total) || 0), 0);
  const progress = totalValue > 0 ? Math.round((doneValue / totalValue) * 10000) / 100 : 0;
  
  const custAgg = {};
  PO.forEach(p => {
    const c = p.CustomerID || "Unknown";
    if (!custAgg[c]) custAgg[c] = { kontrak: 0, realisasi: 0 };
    custAgg[c].kontrak += parseFloat(p.Total) || 0;
    if (p.Status === "Done") custAgg[c].realisasi += parseFloat(p.Total) || 0;
  });
  const topCustomers = Object.entries(custAgg).sort((a, b) => b[1].kontrak - a[1].kontrak).slice(0, 5).map(([id, v]) => ({ id, kontrak: v.kontrak, realisasi: v.realisasi, pct: v.kontrak > 0 ? Math.round((v.realisasi / v.kontrak) * 100) : 0 }));

  return { totalPO, totalOpen, totalDone, totalValue, doneValue, progress, topCustomers };
}

function setupSheets() {
  const headers = {
    Customers: ["ID", "Name", "Address", "Phone", "CreatedAt"],
    Projects: ["ID", "CustomerID", "ProjectName", "Location", "CreatedAt"],
    Products: ["ID", "ProductType", "Diameter", "Length", "Class", "Type", "Unit", "CreatedAt"],
    PO: ["ID", "CustomerID", "ProjectID", "PODate", "ProductID", "UnitPrice", "Qty", "Total", "Status", "Notes", "CreatedAt"],
  };
  Object.entries(headers).forEach(([name, cols]) => {
    let sh = SS.getSheetByName(name) || SS.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff");
    sh.setFrozenRows(1);
  });
  SpreadsheetApp.getUi().alert("✅ Setup complete!");
}
