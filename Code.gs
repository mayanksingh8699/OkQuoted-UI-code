/************ Code.gs (REPLACE FULL FILE) ************/

const CONFIG = {
  SHEETS: {
    ACCESS: "Access",
    PRODUCT_MASTER: "Product_Master",

    LEAD_HEADER: "Lead_Header",
    LEAD_LINES: "Lead_Lines",

    LISTS: "Lists",

    VENDOR_MASTER: "vendor_master",
    VENDOR_RATES: "vendor_rates",

    MATCHING_QUEUE: "Matching_Queue",
    SALES: "Sales_Orders"
  },
  STAGES: ["New Lead", "Qualified", "Sourcing", "Quoted", "Negotiation", "Won", "Lost"]
};

/************ WEB APP ************/
function doGet(e) {
  const email = Session.getActiveUser().getEmail();
  if (!email || !email.endsWith("@okquoted.com")) {
    return HtmlService.createHtmlOutput("Access denied. Use your okquoted.com account.");
  }

  return HtmlService.createTemplateFromFile("WebApp")
    .evaluate()
    .setTitle("OkQuoted Sales Control")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/************ ROLE ************/
function getUserRole_() {
  const email = (Session.getActiveUser().getEmail() || "").toLowerCase();
  if (!email) return "View";

  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.ACCESS);
  if (!sh) return "View";

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowEmail = (data[i][0] || "").toString().toLowerCase();
    if (rowEmail === email) return data[i][1] || "View";
  }
  return "View";
}

function getSessionInfo() {
  const email = Session.getActiveUser().getEmail() || "";
  return { email, role: getUserRole_() };
}

/************ HELPERS ************/
function getSheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error(`Missing sheet: ${name}`);
  return sh;
}

function headers_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

function now_() {
  return new Date();
}

function nextId_(key, prefix) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const next = Number(props.getProperty(key) || 0) + 1;
    props.setProperty(key, next);
    return `${prefix}${String(next).padStart(6, "0")}`;
  } finally {
    lock.releaseLock();
  }
}

/************ LISTS ************/
function getListValues() {
  const sh = getSheet_(CONFIG.SHEETS.LISTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const out = {};
  headers.forEach((h, i) => {
    out[h] = data.slice(1).map(r => r[i]).filter(v => v !== "" && v != null);
  });

  // Safe fallback
  out["Stages"] = out["Stages"] && out["Stages"].length ? out["Stages"] : CONFIG.STAGES;
  out["Lost_Reasons"] = out["Lost_Reasons"] || ["Price", "Lead time", "Payment/Credit", "Specs mismatch", "No response", "Vendor unavailable", "Cancelled"];
  return out;
}

/************ TAXONOMY (AUTOSUGGEST) ************/
function getTaxonomy() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.PRODUCT_MASTER);
  if (!sh) return { rows: [] };

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return { rows: [] };

  const headers = data[0].map(h => (h ?? "").toString().trim());

  const pick = (...names) => {
    for (const n of names) {
      const i = headers.indexOf(n);
      if (i >= 0) return i;
    }
    return -1;
  };

  const iCat  = pick("Category");
  const iSub  = pick("Sub_Category", "Subcategory", "Sub Category");
  const iSub2 = pick("Sub_Sub_Category", "Sub_Subcategory", "Sub Sub Category");
  const iName = pick("Product_Title", "Product_Name", "Product Name", "Product Title");
  const iPid  = pick("Product_ID", "Product Id", "ProductID");

  const rows = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const productName = (iName >= 0 ? row[iName] : "")?.toString().trim();
    if (!productName) continue;

    rows.push({
      category: (iCat >= 0 ? row[iCat] : "")?.toString().trim(),
      subCategory: (iSub >= 0 ? row[iSub] : "")?.toString().trim(),
      subSubCategory: (iSub2 >= 0 ? row[iSub2] : "")?.toString().trim(),
      productName,
      productId: (iPid >= 0 ? row[iPid] : "")?.toString().trim()
    });
  }
  return { rows };
}

/************ LEADS ************/

/**
 * Creates Lead_Header + multiple Lead_Lines.
 * p.lines is required (array). Each line should have:
 * rawRequirement, qty, neededByDate, category, subCategory, subSubCategory, productName, productId
 */
function createNewLead(p) {
  const role = getUserRole_();
  if (!["Admin", "Sales"].includes(role)) throw new Error("You are not allowed to create leads");

  if (!p.owner || !p.buyerPhone) throw new Error("Owner and Buyer Phone are mandatory");
  if (!Array.isArray(p.lines) || p.lines.length === 0) throw new Error("Add at least 1 product line");

  const lh = getSheet_(CONFIG.SHEETS.LEAD_HEADER);
  const ll = getSheet_(CONFIG.SHEETS.LEAD_LINES);

  const leadId = nextId_("SEQ_LEAD", "L");

  // --- Lead_Header ---
  const hH = headers_(lh);
  const hRow = new Array(hH.length).fill("");
  const setH = (k, v) => { const i = hH.indexOf(k); if (i >= 0) hRow[i] = v; };

  setH("Lead_ID", leadId);
  setH("Created_At", now_());
  setH("Owner", p.owner);
  setH("Stage", "New Lead");
  setH("Buyer_Company", p.buyerCompany || "");
  setH("Buyer_Name", p.buyerName || "");
  setH("Buyer_Phone", p.buyerPhone);
  setH("Delivery_City", p.city || "");
  setH("Delivery_State", p.state || "");

  lh.appendRow(hRow);

  // --- Lead_Lines (multiple) ---
  const lH = headers_(ll);
  const lineIds = [];

  const setL = (rowArr, k, v) => { const i = lH.indexOf(k); if (i >= 0) rowArr[i] = v; };

  p.lines.forEach((line, idx) => {
    const lineNo = String(idx + 1).padStart(2, "0");
    const lineId = `LL${leadId.slice(1)}-${lineNo}`;

    const lRow = new Array(lH.length).fill("");
    setL(lRow, "Lead_Line_ID", lineId);
    setL(lRow, "Lead_ID", leadId);

    setL(lRow, "Raw_Requirement", line.rawRequirement || "");
    setL(lRow, "Qty", line.qty || "");
    setL(lRow, "Needed_By_Date", line.neededByDate || "");
    setL(lRow, "Line_Status", "Active");

    setL(lRow, "Category", line.category || "");
    setL(lRow, "Subcategory", line.subCategory || "");
    setL(lRow, "Sub_Subcategory", line.subSubCategory || "");
    setL(lRow, "Product_Title", line.productName || "");
    setL(lRow, "Product_ID", line.productId || "UNMAPPED");

    ll.appendRow(lRow);
    lineIds.push(lineId);
  });

  return { leadId, lineIds };
}

function updateLeadStage(p) {
  if (!CONFIG.STAGES.includes(p.stage)) throw new Error("Invalid stage");

  const role = getUserRole_();
  if (
    (["Qualified", "Won", "Lost"].includes(p.stage) && !["Admin", "Sales"].includes(role)) ||
    (["Sourcing", "Quoted", "Negotiation"].includes(p.stage) && !["Admin", "Procurement"].includes(role))
  ) {
    throw new Error("You are not allowed to update this stage");
  }

  const sh = getSheet_(CONFIG.SHEETS.LEAD_HEADER);
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxLead = h.indexOf("Lead_ID");
  const idxStage = h.indexOf("Stage");

  const rowIdx = data.findIndex((r, i) => i > 0 && r[idxLead] === p.leadId);
  if (rowIdx === -1) throw new Error("Lead not found");

  sh.getRange(rowIdx + 1, idxStage + 1).setValue(p.stage);

  if (p.stage === "Qualified") {
    const qIdx = h.indexOf("Qualified_At");
    if (qIdx >= 0) sh.getRange(rowIdx + 1, qIdx + 1).setValue(now_());
  }

  if (p.stage === "Lost") {
    if (!p.lostReason) throw new Error("Lost Reason mandatory");
    const lostAt = h.indexOf("Lost_At");
    const lostReason = h.indexOf("Lost_Reason");
    const lostNotes = h.indexOf("Lost_Notes");

    if (lostAt >= 0) sh.getRange(rowIdx + 1, lostAt + 1).setValue(now_());
    if (lostReason >= 0) sh.getRange(rowIdx + 1, lostReason + 1).setValue(p.lostReason);
    if (lostNotes >= 0) sh.getRange(rowIdx + 1, lostNotes + 1).setValue(p.lostNotes || "");
  }

  return { ok: true };
}

/** MUST RETURN ARRAY (UI expects array) */
function listRecentLeads(limit) {
  const sh = getSheet_(CONFIG.SHEETS.LEAD_HEADER);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const h = data[0];
  const col = (name) => h.indexOf(name);

  const iLead = col("Lead_ID");
  const iCreated = col("Created_At");
  const iStage = col("Stage");
  const iOwner = col("Owner");
  const iCompany = col("Buyer_Company");
  const iPhone = col("Buyer_Phone");
  const iCity = col("Delivery_City");
  const iState = col("Delivery_State");

  const out = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const r = data[i];
    if (!r[iLead]) continue;
    out.push({
      leadId: r[iLead],
      createdAt: r[iCreated],
      stage: r[iStage],
      owner: r[iOwner],
      buyerCompany: r[iCompany],
      buyerPhone: r[iPhone],
      city: r[iCity],
      state: r[iState]
    });
    if (out.length >= (limit || 25)) break;
  }
  return out;
}

/************ VENDORS ************/
function createVendorMaster(p) {
  const role = getUserRole_();
  if (!["Admin", "Procurement"].includes(role)) {
    throw new Error("Only Admin/Procurement can onboard vendors");
  }

  const sh = getSheet_(CONFIG.SHEETS.VENDOR_MASTER);
  const h = headers_(sh);
  const row = new Array(h.length).fill("");

  const vendorId = nextId_("SEQ_VENDOR", "V");
  const set = (k, v) => { const i = h.indexOf(k); if (i >= 0) row[i] = v; };

  set("Vendor_ID", vendorId);
  set("Created_At", now_());
  set("Vendor_Name", p.vendorName || "");
  set("Vendor_Phone", p.vendorPhone || "");
  set("City", p.city || "");
  set("State", p.state || "");
  set("Categories_Covered", p.categoriesCovered || "");
  set("Notes", p.notes || "");
  set("Status", "Active");

  sh.appendRow(row);
  return { vendorId };
}

function listVendors() {
  const sh = getSheet_(CONFIG.SHEETS.VENDOR_MASTER);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const h = data[0];
  const iId = h.indexOf("Vendor_ID");
  const iName = h.indexOf("Vendor_Name");
  const iPhone = h.indexOf("Vendor_Phone");
  const iCity = h.indexOf("City");
  const iState = h.indexOf("State");
  const iStatus = h.indexOf("Status");

  const out = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const id = r[iId];
    if (!id) continue;
    const status = iStatus >= 0 ? (r[iStatus] || "Active") : "Active";
    if (String(status).toLowerCase() !== "active") continue;

    out.push({
      vendorId: String(id),
      vendorName: String(r[iName] || ""),
      vendorPhone: String(r[iPhone] || ""),
      city: String(r[iCity] || ""),
      state: String(r[iState] || "")
    });
  }
  return out;
}

/**
 * Save multiple vendor rate lines into vendor_rates
 * p: { vendorId, lines:[{productId, productName, brand, model, unit, unitRate, moq, leadTimeDays, gstPercent, validTill, notes}] }
 */
function saveVendorRates(p) {
  const role = getUserRole_();
  if (!["Admin", "Procurement"].includes(role)) throw new Error("Only Admin/Procurement can add vendor rates");
  if (!p.vendorId) throw new Error("Vendor_ID required");
  if (!Array.isArray(p.lines) || p.lines.length === 0) throw new Error("Add at least 1 rate line");

  const sh = getSheet_(CONFIG.SHEETS.VENDOR_RATES);
  const h = headers_(sh);

  const set = (arr, k, v) => { const i = h.indexOf(k); if (i >= 0) arr[i] = v; };

  const createdBy = Session.getActiveUser().getEmail() || "";
  const createdAt = now_();

  const rateIds = [];

  p.lines.forEach(line => {
    const rateId = nextId_("SEQ_RATE", "R");
    const row = new Array(h.length).fill("");

    set(row, "Rate_ID", rateId);
    set(row, "Created_At", createdAt);
    set(row, "Created_By", createdBy);

    set(row, "Vendor_ID", p.vendorId);

    set(row, "Product_ID", line.productId || "");
    set(row, "Product_Title", line.productName || "");
    set(row, "Brand", line.brand || "");
    set(row, "Model", line.model || "");
    set(row, "Unit", line.unit || "");
    set(row, "Unit_Rate", line.unitRate || "");
    set(row, "MOQ", line.moq || "");
    set(row, "Lead_Time_Days", line.leadTimeDays || "");
    set(row, "GST_Percent", line.gstPercent || "");
    set(row, "Valid_Till", line.validTill || "");
    set(row, "Notes", line.notes || "");
    set(row, "Status", "Active");

    sh.appendRow(row);
    rateIds.push(rateId);
  });

  return { ok: true, rateIds };
}

/************ MATCHING (UNCHANGED) ************/
function requestMatching(p) {
  const role = getUserRole_();
  if (!["Admin", "Sales", "Procurement"].includes(role)) throw new Error("Not allowed");
  if (!p.leadId) throw new Error("Lead_ID required");

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.SHEETS.MATCHING_QUEUE);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEETS.MATCHING_QUEUE);
    sh.appendRow(["Request_ID", "Lead_ID", "Requested_By", "Requested_At", "Status"]);
  }

  const reqId = nextId_("SEQ_MATCHREQ", "MR");
  sh.appendRow([reqId, p.leadId, Session.getActiveUser().getEmail(), now_(), "Queued"]);
  return { requestId: reqId };
}
