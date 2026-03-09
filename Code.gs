// =======================================================
// 1BOSSCO QUEUING SYSTEM - COMPLETE Code.gs
// Includes: Client, Display, Staff, Admin, Config + Reports
// =======================================================

// -------------------------------
// SERVICE → COUNTER MAPPING
// (Used by ticket generator routing)
// -------------------------------
const SERVICE_COUNTERS = {
  // National ID
  "N": [1],
  "PI": [1],

  // PSA (Counters 2-4)
  "PS": [2, 3, 4],
  "PA": [2, 3, 4],

  // DTI
  "DT": [5],
  "PT": [5],

  // PRC / Recruitment (Counters 6-9)
  "PR": [6, 7, 8, 9],
  "PC": [6, 7, 8, 9],
 
  // DMW
  "DM": [10, 11],
  "PD": [10, 11],

  // OWWA
  "OW": [12, 13],
  "PO": [12, 13],

  // NOCAP
  "NC": [14],
  "PN": [14],

  // SSS
  "SS": [15],
  "SP": [15],

  // LTO
  "LT": [17],
  "PL": [17],

  // VFS Global
  "VF": [18, 19],
  "PV": [18, 19],

  // PhilHealth
  "LH": [20],
  "PH": [20],

  // Police Clearance
  "CL": [22],
  "PP": [22],

  // LTOPF
  "OP": [23],
  "PF": [23],

  // PESO
  "P": [24, 25],
  "PE": [24, 25]

};

// -------------------------------
// SERVICE CODES → SERVICE NAMES
// (Displayed + stored in QueueData B)
// -------------------------------
const SERVICE_NAMES = {
  "N": "National ID",
  "PI": "Priority National ID",

  "PS": "PSA",
  "PA": "Priority PSA",

  "DT": "DTI",
  "PT": "Priority DTI",

  "PR": "PRC/Recruitment",
  "PC": "Priority PRC/Recruitment",

  "DM": "DMW",
  "PD": "Priority DMW",

  "OW": "OWWA Payment",
  "PO": "Priority OWWA Payment",

  "NC": "NOCAP",
  "PN": "Priority NOCAP",

  "SS": "SSS",
  "SP": "Priority SSS",

  "LT": "LTO",
  "PL": "Priority LTO",

  "VF": "VFS Global",
  "PV": "Priority VFS Global",

  "LH": "PhilHealth",
  "PH": "Priority PhilHealth",

  "CL": "Police Clearance",
  "PP": "Priority Police Clearance",

  "OP": "LTOPF",
  "PF": "Priority LTOPF",

  "P": "PESO",
  "PE": "Priority PESO"
};

// -------------------------------
// Serve HTML pages
// -------------------------------
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : "client";
  var files = { display: "display", staff: "staff", client: "client", admin: "admin" };
  return HtmlService.createHtmlOutputFromFile(files[page] || "client");
}

// =======================================================
// CONFIG (Admin-driven real-time behavior)
// Sheets used:
// - AgencyCounters
// - AgencyServices
// =======================================================

function getAgencyCounters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AgencyCounters");
  return sh ? sh.getDataRange().getValues() : [];
}

function getAgencyServices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AgencyServices");
  return sh ? sh.getDataRange().getValues() : [];
}

// For Client dropdown (Agency list + codes)
function getClientAgencyOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AgencyCounters");
  if (!sh) return [];

  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const [counterNo, agency, normalCode, priorityCode, active, sortOrder] = rows[i];
    const isActive = !(active === false || String(active).toUpperCase() === "FALSE");
    if (!isActive) continue;
    if (!agency) continue;

    out.push({
      sort: Number(sortOrder) || 999,
      agency: String(agency),
      normalCode: String(normalCode || ""),
      priorityCode: String(priorityCode || "")
    });
  }
  out.sort((a, b) => (a.sort - b.sort) || a.agency.localeCompare(b.agency));
  return out;
}

// For Staff checklist (tickboxes based on counter agency)
function getChecklistForCounter(counterNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const c = ss.getSheetByName("AgencyCounters");
  const s = ss.getSheetByName("AgencyServices");
  if (!c || !s) return { agency: "UNCONFIGURED", items: [] };

  const counter = Number(counterNumber);
  const counters = c.getDataRange().getValues();
  let agency = "";
  for (let i = 1; i < counters.length; i++) {
    if (Number(counters[i][0]) === counter) {
      agency = String(counters[i][1] || "");
      break;
    }
  }
  if (!agency) return { agency: "UNCONFIGURED", items: [] };

  const services = s.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < services.length; i++) {
    const [aName, item, active, sort] = services[i];
    const isActive = !(active === false || String(active).toUpperCase() === "FALSE");
    if (String(aName) === agency && isActive && item) {
      items.push({ item: String(item), sort: Number(sort) || 999 });
    }
  }
  items.sort((x, y) => (x.sort - y.sort) || x.item.localeCompare(y.item));
  return { agency, items: items.map(x => x.item) };
}

// Admin save config
function saveAgencyCounters(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("AgencyCounters");
  if (!sh) sh = ss.insertSheet("AgencyCounters");

  sh.clearContents();
  sh.appendRow(["CounterNo", "AgencyName", "ClientNormalCode", "ClientPriorityCode", "Active", "SortOrder"]);
  data.forEach(r => sh.appendRow([
    r.counterNo, r.agency, r.normalCode, r.priorityCode, r.active, r.sortOrder
  ]));
  return "AgencyCounters saved.";
}

function saveAgencyServices(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("AgencyServices");
  if (!sh) sh = ss.insertSheet("AgencyServices");

  sh.clearContents();
  sh.appendRow(["AgencyName", "ServiceItem", "Active", "SortOrder"]);
  data.forEach(r => sh.appendRow([r.agency, r.item, r.active, r.sortOrder]));
  return "AgencyServices saved.";
}

// =======================================================
// SERVICES (for transfer dropdown)
// Reads from Services sheet (easy to edit in spreadsheet template)
// =======================================================
function getServices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Services");
  if (!sh) return [["ServiceCode", "ServiceName"]];
  return sh.getDataRange().getValues();
}

// =======================================================
// QUEUE FUNCTIONS
// Sheets used:
// - QueueData
// - Counters
// =======================================================

// Generates ticket from selected agency + gender + priority (client.html)
function generateTicketFromAgency(agencyName, gender, isPriority) {
  const options = getClientAgencyOptions();
  const found = options.find(x => x.agency === agencyName);
  if (!found) throw new Error("Agency not found or not active: " + agencyName);

  const code = isPriority ? found.priorityCode : found.normalCode;
  if (!code) throw new Error("No service code configured for this agency (check AgencyCounters).");

  return generateTicket(code, gender);
}

// Generate ticket by serviceCode
function generateTicket(serviceCode, gender) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QueueData");
  if (!sheet) throw new Error("QueueData sheet not found!");

  const counterNumber = getNextCounter(serviceCode);
  const counterName = "Counter " + counterNumber;

  const serviceName = SERVICE_NAMES[serviceCode] || serviceCode;

  // Find last ticket number for this service (scan last 500)
  const lastRow = sheet.getLastRow();
  let lastTicketNum = 0;

  if (lastRow > 1) {
    const scan = Math.min(500, lastRow - 1);
    const start = lastRow - scan + 1;
    const allTickets = sheet.getRange(start, 1, scan, 1).getValues();
    for (let i = allTickets.length - 1; i >= 0; i--) {
      const t = allTickets[i][0];
      if (t && String(t).startsWith(serviceCode + "-")) {
        lastTicketNum = parseInt(String(t).split("-")[1], 10);
        break;
      }
    }
  }

  const ticketNumber = serviceCode + "-" + Utilities.formatString("%03d", lastTicketNum + 1);

  // QueueData columns:
  // A Ticket, B Service, C Status, D Counter, E DateTime, F Gender, G Availed Services, H Notes
  sheet.appendRow([
    ticketNumber,
    serviceName,
    "Waiting",
    counterName,
    new Date(),
    gender || "",
    "",
    ""
  ]);

  // Update Counters sheet view (Service column can show last service)
  const countersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Counters");
  if (countersSheet) {
    const counters = countersSheet.getDataRange().getValues();
    const counterRow = counters.findIndex(row => row[0] === counterName);
    if (counterRow !== -1) {
      countersSheet.getRange(counterRow + 1, 3).setValue(serviceName);
    }
  }

  return {
    ticket: ticketNumber,
    service: serviceName,
    counter: counterName
  };
}

// Round-robin counter assignment for multi-counter agencies
function getNextCounter(serviceCode) {
  const counters = SERVICE_COUNTERS[serviceCode];
  if (!counters || counters.length === 0) throw new Error("No counters assigned for service: " + serviceCode);

  if (counters.length === 1) return counters[0];

  const props = PropertiesService.getScriptProperties();
  const key = "LAST_COUNTER_" + serviceCode;
  const lastUsed = props.getProperty(key);
  let nextIndex = 0;

  if (lastUsed !== null) {
    const lastIndex = counters.indexOf(Number(lastUsed));
    nextIndex = (lastIndex + 1) % counters.length;
  }

  const nextCounter = counters[nextIndex];
  props.setProperty(key, nextCounter);
  return nextCounter;
}

// Display page: all counters table
function getCounters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Counters");
  return sheet.getDataRange().getValues();
}

// Display page: last rows of queue
function getQueueData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("QueueData");
  const all = qSheet.getDataRange().getValues();
  const header = all[0];
  let data = all.slice(-80);
  if (data.length < 80) data = all;
  return [header].concat(data);
}

// Status of a counter (staff)
function getCounterStatus(counterNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cSheet = ss.getSheetByName("Counters");
  const values = cSheet.getDataRange().getValues();
  const counterName = "Counter " + Number(counterNo);

  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === counterName) {
      return {
        counter: values[i][0],
        ticket: values[i][1] || "None",
        service: values[i][2] || "-",
        status: values[i][3] || "Idle"
      };
    }
  }
  return { counter: counterName, ticket: "None", service: "-", status: "Idle" };
}

// Call next client (staff)
// Priority handling can be extended later; currently FIFO per counter.
function nextClient(counterNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName("QueueData");
  const countersSheet = ss.getSheetByName("Counters");
  const counterName = "Counter " + Number(counterNumber);

  const lastRow = queueSheet.getLastRow();
  if (lastRow < 2) return "No clients waiting";

  const scanRows = Math.min(400, lastRow - 1);
  const startRow = lastRow - scanRows + 1;
  const queue = queueSheet.getRange(startRow, 1, scanRows, 4).getValues(); // A-D

  // -------------------------------------
  // PRIORITY / REGULAR ALTERNATING LOGIC
  // -------------------------------------
  let waitingRegular = [];
  let waitingPriority = [];

  for (let i = 0; i < queue.length; i++) {
    if (queue[i][2] !== "Waiting") continue;
    if (queue[i][3] !== counterName) continue;

    const service = String(queue[i][1] || "");
    const rowNumber = startRow + i;

    if (service.toLowerCase().includes("priority")) {
      waitingPriority.push(rowNumber);
    } else {
      waitingRegular.push(rowNumber);
    }
  }

  let foundRow = -1;

  // Read remembered last served type for this counter
  const props = PropertiesService.getScriptProperties();
  const lastTypeKey = "LAST_SERVED_TYPE_" + Number(counterNumber);
  const lastType = props.getProperty(lastTypeKey) || "";

  if (lastType === "priority") {
    if (waitingRegular.length > 0) {
      foundRow = waitingRegular[0];
    } else if (waitingPriority.length > 0) {
      foundRow = waitingPriority[0];
    }
  } else if (lastType === "regular") {
    if (waitingPriority.length > 0) {
      foundRow = waitingPriority[0];
    } else if (waitingRegular.length > 0) {
      foundRow = waitingRegular[0];
    }
  } else {
    // First call / no memory yet:
    // if any priority is waiting, serve it next
    if (waitingPriority.length > 0) {
      foundRow = waitingPriority[0];
    } else if (waitingRegular.length > 0) {
      foundRow = waitingRegular[0];
    }
  }

  if (foundRow === -1) return "No clients waiting";

  const ticket = queueSheet.getRange(foundRow, 1).getValue();
  const service = queueSheet.getRange(foundRow, 2).getValue();

  // Update QueueData
  queueSheet.getRange(foundRow, 3, 1, 2).setValues([["Serving", counterName]]);

  // Update Counters sheet row
  const counters = countersSheet.getDataRange().getValues();
  let counterRow = counters.findIndex(row => row[0] === counterName);
  if (counterRow === -1) {
    countersSheet.appendRow([counterName, "-", "-", "Idle", "", "", "", "", 0]);
    counterRow = countersSheet.getLastRow() - 1;
  }
  countersSheet.getRange(counterRow + 1, 2, 1, 3).setValues([[ticket, service, "Serving"]]);

  // Remember type of currently called client
  const servedType = String(service).toLowerCase().includes("priority") ? "priority" : "regular";
  props.setProperty(lastTypeKey, servedType);

  return ticket + " (" + service + ")";
}

// Mark done (staff)
function completeClient(counterNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const countersSheet = ss.getSheetByName("Counters");
  const queueSheet = ss.getSheetByName("QueueData");
  const counterName = "Counter " + Number(counterNumber);

  const counters = countersSheet.getDataRange().getValues();
  const counterRow = counters.findIndex(row => row[0] === counterName);

  if (counterRow !== -1) {
    const ticket = counters[counterRow][1];
    if (!ticket || ticket === "-") return "No client serving in this counter.";

    // Find ticket in last 500 rows
    const lastRow = queueSheet.getLastRow();
    const scanRows = Math.min(500, lastRow - 1);
    const start = lastRow - scanRows + 1;
    const tickets = queueSheet.getRange(start, 1, scanRows, 1).getValues();
    for (let i = 0; i < tickets.length; i++) {
      if (tickets[i][0] === ticket) {
        queueSheet.getRange(start + i, 3).setValue("Completed");
        break;
      }
    }

    countersSheet.getRange(counterRow + 1, 2, 1, 3).setValues([["-", "-", "Idle"]]);
    return "Client " + ticket + " Completed!";
  }
  return "No client serving in this counter.";
}

// Transfer client (staff)
function transferClient(ticket, fromCounter, toCounter, newServiceCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName("QueueData");

  const serviceName = SERVICE_NAMES[newServiceCode] || newServiceCode;

  // Find last ticket for this newServiceCode
  const lastRow = queueSheet.getLastRow();
  let newNumber = 1;
  if (lastRow > 1) {
    const scan = Math.min(500, lastRow - 1);
    const start = lastRow - scan + 1;
    const tickets = queueSheet.getRange(start, 1, scan, 1).getValues();
    for (let i = tickets.length - 1; i >= 0; i--) {
      const t = tickets[i][0];
      if (t && String(t).startsWith(newServiceCode + "-")) {
        newNumber = parseInt(String(t).split("-")[1], 10) + 1;
        break;
      }
    }
  }

  const newTicket = newServiceCode + "-" + Utilities.formatString("%03d", newNumber);
  const assignedCounter = "Counter " + Number(toCounter);

  queueSheet.appendRow([newTicket, serviceName, "Waiting", assignedCounter, new Date(), "", "", "Transferred from " + ticket]);
  return "Transferred " + ticket + " → " + assignedCounter + " (" + newTicket + ")";
}

// =======================================================
// STAFF EXTRA: Call Again + Upcoming Preview
// =======================================================
function getCurrentServing(counterNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cSheet = ss.getSheetByName("Counters");
  const counterName = "Counter " + Number(counterNumber);

  const values = cSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === counterName) {
      return {
        counter: counterName,
        ticket: values[i][1] || "",
        service: values[i][2] || "",
        status: values[i][3] || "Idle"
      };
    }
  }
  return { counter: counterName, ticket: "", service: "", status: "Idle" };
}

function getUpcomingClients(counterNumber, limit) {
  limit = Number(limit) || 5;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("QueueData");
  const counterName = "Counter " + Number(counterNumber);

  const lastRow = qSheet.getLastRow();
  if (lastRow < 2) return [];

  const scanRows = Math.min(400, lastRow - 1);
  const startRow = lastRow - scanRows + 1;
  const data = qSheet.getRange(startRow, 1, scanRows, 4).getValues();

  const upcoming = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][2] === "Waiting" && data[i][3] === counterName) {
      upcoming.push({ ticket: data[i][0], service: data[i][1], counter: data[i][3] });
      if (upcoming.length >= limit) break;
    }
  }
  return upcoming;
}

// =======================================================
// RECORDING: Staff tickboxes → Logs + Transactions
// =======================================================

function _getGenderByTicket_(ticket) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const q = ss.getSheetByName("QueueData");
  if (!q) return "";
  const last = q.getLastRow();
  if (last < 2) return "";

  const scan = Math.min(600, last - 1);
  const start = last - scan + 1;
  const values = q.getRange(start, 1, scan, 6).getValues(); // A-F

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === ticket) return values[i][5] || "";
  }
  return "";
}

function _getMainServiceByTicket_(ticket) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const q = ss.getSheetByName("QueueData");
  if (!q) return "";
  const last = q.getLastRow();
  if (last < 2) return "";

  const scan = Math.min(600, last - 1);
  const start = last - scan + 1;
  const values = q.getRange(start, 1, scan, 2).getValues(); // A-B
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === ticket) return values[i][1] || "";
  }
  return "";
}

function _getAgencyByCounter_(counterNumber) {
  const cfg = getChecklistForCounter(counterNumber);
  return cfg.agency || "UNKNOWN";
}

function appendTransaction_(agency, counterNumber, ticket, gender, mainService, availedServicesCsv) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Transactions");
  if (!sh) {
    sh = ss.insertSheet("Transactions");
    sh.appendRow(["DateTime","Date","Year","Month","Agency","Counter","Ticket","Gender","Main Service","Availed Services"]);
  }

  const now = new Date();
  const tz = ss.getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const year = Number(Utilities.formatDate(now, tz, "yyyy"));
  const month = Number(Utilities.formatDate(now, tz, "M"));

  sh.appendRow([
    now,
    dateStr,
    year,
    month,
    agency,
    "Counter " + Number(counterNumber),
    ticket,
    gender || "",
    mainService || "",
    availedServicesCsv || ""
  ]);
}

function recordChecklistAvailed(counterNumber, ticket, checklistItems) {
  counterNumber = Number(counterNumber);
  const agency = _getAgencyByCounter_(counterNumber);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "LOG_" + String(agency).replace(/\s+/g, "_").toUpperCase();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(["Date","Time","Ticket","Counter","Agency","Gender(auto)","Checklist Selected"]);
  }

  const now = new Date();
  const tz = ss.getSpreadsheetTimeZone();
  const items = Array.isArray(checklistItems) ? checklistItems : [];
  const gender = _getGenderByTicket_(ticket);
  const mainService = _getMainServiceByTicket_(ticket);

  sh.appendRow([
    Utilities.formatDate(now, tz, "yyyy-MM-dd"),
    Utilities.formatDate(now, tz, "HH:mm:ss"),
    ticket,
    "Counter " + counterNumber,
    agency,
    gender,
    items.join(", ")
  ]);

  // write back to QueueData (Availed Services column G)
  const q = ss.getSheetByName("QueueData");
  if (q) {
    const last = q.getLastRow();
    const scan = Math.min(600, last - 1);
    const start = last - scan + 1;
    const values = q.getRange(start, 1, scan, 1).getValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i][0] === ticket) {
        q.getRange(start + i, 7).setValue(items.join(", ")); // column G
        break;
      }
    }
  }

  // master Transactions for Admin reports
  appendTransaction_(agency, counterNumber, ticket, gender, mainService, items.join(", "));

  return `Recorded: ${agency} | ${ticket} | ${items.join(", ")}`;
}

// =======================================================
// ADMIN REPORTING (Daily / Monthly / Yearly)
// Reads from Transactions sheet
// =======================================================
function getAdminReport(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Transactions");
  if (!sh) return { summary:{ total:0 }, byService:[], byAgency:[], byGender:[] };

  const rows = sh.getDataRange().getValues();
  const data = rows.slice(1);

  const mode = params.mode; // "daily" | "monthly" | "yearly"
  const agencyFilter = params.agency || "ALL";

  const y = Number(params.year);
  const m = Number(params.month);
  const d = String(params.date || "");

  const byService = new Map();
  const byAgency = new Map();
  const byGender = new Map();
  let total = 0;

  data.forEach(r => {
    const dateStr = r[1];
    const year = Number(r[2]);
    const month = Number(r[3]);
    const agency = String(r[4] || "");
    const gender = String(r[7] || "Unknown") || "Unknown";
    const availed = String(r[9] || "");

    if (agencyFilter !== "ALL" && agency !== agencyFilter) return;

    if (mode === "daily" && dateStr !== d) return;
    if (mode === "monthly" && !(year === y && month === m)) return;
    if (mode === "yearly" && year !== y) return;

    total++;
    byAgency.set(agency, (byAgency.get(agency) || 0) + 1);
    byGender.set(gender, (byGender.get(gender) || 0) + 1);

    availed.split(",").map(x=>x.trim()).filter(Boolean).forEach(item=>{
      byService.set(item, (byService.get(item) || 0) + 1);
    });
  });

  const toArr = (mp) => Array.from(mp.entries())
    .map(([name,count])=>({name,count}))
    .sort((a,b)=>b.count-a.count);

  return {
    summary: { total },
    byService: toArr(byService),
    byAgency: toArr(byAgency),
    byGender: toArr(byGender)
  };
}

// =======================================================
// ✅ MISSING STAFF FUNCTIONS + CALENDAR (IMPLEMENT NOW)
// Requires sheets:
// CalendarEvents, DailyReset, CounterHistory, AuditLogs, SystemSettings
// =======================================================

function _ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers && headers.length) sh.appendRow(headers);
  return sh;
}

function _logCounterHistory_(counterNumber, ticket, service, action, staffName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = _ensureSheet_("CounterHistory", ["DateTime","Date","Counter","Ticket","Service","Action","Staff"]);
  const now = new Date();
  const tz = ss.getSpreadsheetTimeZone();
  sh.appendRow([
    now,
    Utilities.formatDate(now, tz, "yyyy-MM-dd"),
    "Counter " + Number(counterNumber),
    ticket || "",
    service || "",
    action || "",
    staffName || ("Counter " + Number(counterNumber))
  ]);
}

function _logAudit_(page, user, action, reference, details) {
  const sh = _ensureSheet_("AuditLogs", ["DateTime","Page","User","Action","Reference","Details"]);
  sh.appendRow([new Date(), page || "", user || "", action || "", reference || "", details || ""]);
}

// -------------------------------------------------------
// ✅ Upcoming Summary (used by staff.html)
// staff.html calls: getUpcomingSummary(counter, 5)
// Return: { total:number, list:[{ticket,service,counter}] }
// -------------------------------------------------------
function getUpcomingSummary(counterNumber, limit) {
  limit = Number(limit) || 5;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("QueueData");
  if (!qSheet) return { total: 0, list: [] };

  const counterName = "Counter " + Number(counterNumber);

  const lastRow = qSheet.getLastRow();
  if (lastRow < 2) return { total: 0, list: [] };

  const scanRows = Math.min(800, lastRow - 1);
  const startRow = lastRow - scanRows + 1;
  const data = qSheet.getRange(startRow, 1, scanRows, 4).getValues(); // A-D

  let total = 0;
  const list = [];

  for (let i = 0; i < data.length; i++) {
    const ticket = data[i][0];
    const service = data[i][1];
    const status = data[i][2];
    const counter = data[i][3];

    if (status === "Waiting" && counter === counterName) {
      total++;
      if (list.length < limit) {
        list.push({ ticket, service, counter });
      }
    }
  }

  return { total, list };
}

// -------------------------------------------------------
// ✅ No Show (used by staff.html)
// staff.html calls: markNoShow(counter)
// Behavior:
// - Marks the currently Serving ticket for that counter as "No Show" in QueueData
// - Clears Counters row back to Idle
// - Logs to CounterHistory + AuditLogs
// -------------------------------------------------------
function markNoShow(counterNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const countersSheet = ss.getSheetByName("Counters");
  const queueSheet = ss.getSheetByName("QueueData");
  if (!countersSheet || !queueSheet) throw new Error("Missing Counters or QueueData sheet.");

  const counterName = "Counter " + Number(counterNumber);
  const counters = countersSheet.getDataRange().getValues();
  const counterRow = counters.findIndex(r => r[0] === counterName);

  if (counterRow === -1) return "Counter not found.";

  const ticket = counters[counterRow][1];
  const service = counters[counterRow][2];

  if (!ticket || ticket === "-" || ticket === "None") {
    return "No client serving in this counter.";
  }

  // Find ticket in QueueData (scan last 800)
  const lastRow = queueSheet.getLastRow();
  const scanRows = Math.min(800, lastRow - 1);
  const start = lastRow - scanRows + 1;
  const values = queueSheet.getRange(start, 1, scanRows, 8).getValues(); // A-H

  let found = false;
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === ticket) {
      queueSheet.getRange(start + i, 3).setValue("No Show"); // Status col C
      // append note
      const oldNote = values[i][7] || "";
      const note = (oldNote ? (oldNote + " | ") : "") + ("No Show at " + counterName);
      queueSheet.getRange(start + i, 8).setValue(note); // Notes col H
      found = true;
      break;
    }
  }

  // Clear counter
  countersSheet.getRange(counterRow + 1, 2, 1, 3).setValues([["-", "-", "Idle"]]);

  _logCounterHistory_(counterNumber, ticket, service, "NO_SHOW", counterName);
  _logAudit_("STAFF", counterName, "NO_SHOW", ticket, `Counter=${counterName} Service=${service || ""}`);

  return found ? ("Marked No Show: " + ticket) : ("Marked No Show (ticket not found in recent scan): " + ticket);
}

// -------------------------------------------------------
// ✅ Transfer Same Ticket (used by staff.html)
// staff.html calls: transferClientSameTicket(currentTicket, fromCounter, toAgency)
// Behavior:
// - Ticket stays the same
// - Status becomes Waiting
// - Counter auto-assigned based on AgencyCounters rows (agency→counter list)
// - Service column updated to the agency name (clean for display)
// - Logs to CounterHistory + AuditLogs
// -------------------------------------------------------
function transferClientSameTicket(ticket, fromCounterNumber, toAgencyName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName("QueueData");
  const aSheet = ss.getSheetByName("AgencyCounters");
  const countersSheet = ss.getSheetByName("Counters");

  if (!qSheet) throw new Error("QueueData sheet not found.");
  if (!aSheet) throw new Error("AgencyCounters sheet not found.");
  if (!ticket) throw new Error("Ticket is required.");

  const fromCounterName = "Counter " + Number(fromCounterNumber);
  const toAgency = String(toAgencyName || "").trim();
  if (!toAgency) throw new Error("Target agency is required.");

  // ✅ Case-insensitive agency match
  const toAgencyKey = toAgency.toLowerCase();

  // 1) Eligible counters for that agency
  const cfg = aSheet.getDataRange().getValues();
  const eligible = [];
  for (let i = 1; i < cfg.length; i++) {
    const counterNo = Number(cfg[i][0]);
    const agencyName = String(cfg[i][1] || "").trim();
    const active = cfg[i][4];
    const isActive = !(active === false || String(active).toUpperCase() === "FALSE");

    if (isActive && agencyName.toLowerCase() === toAgencyKey && counterNo) eligible.push(counterNo);
  }
  if (eligible.length === 0) throw new Error("No active counters configured for agency: " + toAgency);

  // 2) Round-robin choose counter
  eligible.sort((a, b) => a - b);
  const props = PropertiesService.getScriptProperties();
  const key = "LAST_AGENCY_COUNTER_" + toAgencyKey.replace(/\s+/g, "_").toUpperCase();
  const lastUsed = Number(props.getProperty(key) || "");
  let idx = 0;

  if (lastUsed) {
    const lastIdx = eligible.indexOf(lastUsed);
    idx = (lastIdx >= 0) ? ((lastIdx + 1) % eligible.length) : 0;
  }

  const chosen = eligible[idx];
  props.setProperty(key, String(chosen));
  const assignedCounter = "Counter " + chosen;

  // 3) Find the ticket row (scan last 2000)
  const lastRow = qSheet.getLastRow();
  const scanRows = Math.min(2000, lastRow - 1);
  const start = lastRow - scanRows + 1;
  const values = qSheet.getRange(start, 1, scanRows, 8).getValues(); // A-H

  let rowIndex = -1;
  let rowData = null;

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === String(ticket)) {
      rowIndex = start + i;
      rowData = values[i].slice(); // copy A-H
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Ticket not found in QueueData: " + ticket);

  // ✅ KEEP column B unchanged (do NOT rename agency/service)
  rowData[2] = "Waiting";         // C Status
  rowData[3] = assignedCounter;   // D Counter
  rowData[4] = new Date();        // E DateTime (transfer time)

  const oldNotes = String(rowData[7] || "");
  const note = `Transferred from ${fromCounterName} to ${assignedCounter}`;
  rowData[7] = oldNotes ? (oldNotes + " | " + note) : note;

  // ✅ Make it the LAST in target queue: append then delete old row
  qSheet.appendRow(rowData);
  qSheet.deleteRow(rowIndex);

  // Clear FROM counter in Counters sheet
  if (countersSheet) {
    const cVals = countersSheet.getDataRange().getValues();
    const fromRow = cVals.findIndex(r => String(r[0]) === fromCounterName);
    if (fromRow !== -1) {
      countersSheet.getRange(fromRow + 1, 2, 1, 3).setValues([["-", "-", "Idle"]]);
    }
  }

  return `Transferred ${ticket} → ${toAgency} (${assignedCounter})`;
}

// -------------------------------------------------------
// ✅ Calendar (used by staff.html)
// staff.html calls:
// - addAgencyCalendarEvent(counter, date, type, note)
// - getCalendarEvents(from, to, 80)
// -------------------------------------------------------
function addAgencyCalendarEvent(counterNumber, dateStr, type, notice) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = _ensureSheet_("CalendarEvents", ["Date","Agency","Counter","Type","Notice","CreatedAt","CreatedBy"]);

  const counterName = "Counter " + Number(counterNumber);
  const cfg = getChecklistForCounter(counterNumber); // returns { agency, items }
  const agency = (cfg && cfg.agency) ? cfg.agency : "UNKNOWN";

  const d = String(dateStr || "").trim();
  if (!d) throw new Error("Date is required (YYYY-MM-DD).");

  const t = String(type || "").trim();
  const n = String(notice || "").trim();

  const createdBy = counterName;
  sh.appendRow([d, agency, counterName, t, n, new Date(), createdBy]);

  _logAudit_("STAFF", counterName, "CALENDAR_ADD", d, `Agency=${agency} Type=${t}`);
  return "Saved calendar notice for " + agency + " (" + d + ").";
}

function getCalendarEvents(fromDateStr, toDateStr, limit) {
  limit = Number(limit) || 80;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("CalendarEvents");
  if (!sh) return [];

  const rows = sh.getDataRange().getValues();
  const data = rows.slice(1);

  const from = fromDateStr ? new Date(fromDateStr) : null;
  const to = toDateStr ? new Date(toDateStr) : null;

  const out = [];
  for (let i = 0; i < data.length; i++) {
    const dStr = String(data[i][0] || "").trim(); // Date (yyyy-mm-dd)
    if (!dStr) continue;

    const d = new Date(dStr);
    if (from && d < from) continue;
    if (to && d > to) continue;

    out.push({
      date: dStr,
      agency: String(data[i][1] || ""),
      counter: String(data[i][2] || ""),
      type: String(data[i][3] || ""),
      notice: String(data[i][4] || "")
    });
  }

  // Sort ascending by date (then agency)
  out.sort((a,b)=> (a.date.localeCompare(b.date)) || (a.agency.localeCompare(b.agency)));

  return out.slice(0, limit);
}

function triggerRecall(counterNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cSheet = ss.getSheetByName("Counters");
  const counterName = "Counter " + Number(counterNumber);

  if (!cSheet) throw new Error("Counters sheet not found.");

  const values = cSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === counterName) {
      const ticket = values[i][1] || "";
      const service = values[i][2] || "";
      const status = values[i][3] || "Idle";

      if (!ticket || ticket === "-" || ticket === "None" || status !== "Serving") {
        throw new Error("No serving client to call again.");
      }

      const props = PropertiesService.getScriptProperties();
      const payload = {
        counter: counterName,
        ticket: ticket,
        service: service,
        ts: new Date().getTime()
      };
      props.setProperty("RECALL_" + Number(counterNumber), JSON.stringify(payload));

      return payload;
    }
  }

  throw new Error("Counter not found.");
}

function getRecallSignals() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const out = [];

  Object.keys(props).forEach(key => {
    if (key.startsWith("RECALL_")) {
      try {
        out.push(JSON.parse(props[key]));
      } catch (e) {}
    }
  });

  return out;
}