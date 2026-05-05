// ============================================================
// ROYAL GLASS — Work Order Dashboard
// Version 2.0 | May 2026
//
// Future DB reference:
//   Work Orders   → jobs table          (PK: col G / uuid)
//   Job Materials → job_materials table (PK: col R / uuid)
//   Audit Log     → audit_log table
//
// SM8 columns are read-only — refreshed from ServiceM8 API.
// Manual columns are never overwritten by refresh.
// ============================================================


// ============================================================
// CONFIG — All settings live here. Nothing else needs changing.
// ============================================================
const CONFIG = {

  sheets: {
    workOrders:   "Work Orders",
    jobMaterials: "Job Materials",
    auditLog:     "Audit Log",
  },

  // Column index map (1-based). Matches future DB field names.
  workOrders: {
    jobNo: 1, clientName: 2, siteAddress: 3, projectDesc: 4,
    dateCreated: 5, status: 6, uuid: 7,
    sm8ColCount: 5,   // Columns A–E are SM8-managed (never touch manually)
    totalCols:   7,
  },

  jobMaterials: {
    jobNo: 1, clientName: 2, siteAddress: 3, dateCreated: 4,
    projectDesc: 5, itemName: 6, shortLabel: 7, installer: 8,
    productType: 9, designStatus: 10, expectedDate: 11, confirmedDate: 12,
    glassStatus: 13, hardwareStatus: 14, siteReady: 15, overallStatus: 16,
    remarks: 17, uuid: 18,
    sm8ColCount: 6,   // Columns A–F are SM8-managed
    totalCols:   18,
  },

  // Items matching any keyword are excluded from the dashboard entirely
  excludeKeywords: ["partial invoice", "invoice", "deposit"],

  ai: {
    model:      "claude-haiku-4-5-20251001",
    maxTokens:  120,
    scriptPropKey: "ANTHROPIC_API_KEY",
  },

  sm8: {
    scriptPropKey: "SM8_API_KEY",
    baseUrl:       "https://api.servicem8.com/api_1.0",
  },

  theme: {
    headerBg:   "#2c5f8a",
    headerFont: "#ffffff",
    sm8Bg:      "#f5f5f5",   // Light grey — read-only SM8 columns
    labelBg:    "#e8f4fd",   // Light blue — AI label column
    manualBg:   "#ffffff",   // White — manual input columns
  },

  dropdowns: {
    woStatus:       ["In Progress", "On Hold", "Ready to Install", "Completed"],
    installer:      ["Royal Glass", "Mike"],
    productType:    ["Custom Glass", "Shower Glass", "Shower & General Glass", "Toughened Glass", "Partition Glass", "Glass + Hardware", "Non-production"],
    designStatus:   ["Confirmed", "Pending", "Not Required"],
    glassStatus:    ["Not Ordered", "Ordered", "Partial Supplied", "Delivered", "N/A"],
    hardwareStatus: ["Not Purchased", "Ordered", "In Stock", "N/A"],
    overallStatus:  ["In Progress", "On Hold", "Completed"],
  },

};

// ============================================================
// LOAD BACK UP
// ============================================================

function loadBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const uniqueDates = [...new Set(
    ss.getSheets()
      .filter(s => s.getName().startsWith("Backup —"))
      .sort((a, b) => b.getName().localeCompare(a.getName()))
      .map(s => s.getName().split(" | ")[0].replace("Backup — ", ""))
  )];

  if (uniqueDates.length === 0) { ui.alert("No backups found."); return; }

  const list   = uniqueDates.map((d, i) => `${i + 1}. ${d}`).join("\n");
  const result = ui.prompt("Load Backup", "Available backups:\n\n" + list + "\n\nType the number to restore:", ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const choice = parseInt(result.getResponseText().trim()) - 1;
  if (isNaN(choice) || choice < 0 || choice >= uniqueDates.length) {
    ui.alert("Invalid selection.");
    return;
  }

  const selectedDate = uniqueDates[choice];
  const confirm = ui.alert("Restore Backup", "Overwrite current data with backup from:\n" + selectedDate + "\n\nThis cannot be undone. Continue?", ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  [CONFIG.sheets.workOrders, CONFIG.sheets.jobMaterials].forEach(name => {
    const backup = ss.getSheetByName("Backup — " + selectedDate + " | " + name);
    const target = ss.getSheetByName(name);
    if (!backup || !target) return;
  
    const lastRow = backup.getLastRow();
    const lastCol = backup.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;
  
    const range = backup.getRange(1, 1, lastRow, lastCol);
  
    const values = range.getValues();
    const notes  = range.getNotes();   // ← captures all hover tooltips
  
    target.clearContents();
    target.getRange(1, 1, lastRow, lastCol).setValues(values);
    target.getRange(1, 1, lastRow, lastCol).setNotes(notes);  // ← restores them
  });

  ui.alert("Restored backup from " + selectedDate);
}


// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔧 Royal Glass")
    .addItem("🔄 Refresh — Last 7 Days",  "refreshLast7Days")
    .addItem("🔄 Refresh — Last 30 Days", "refreshLast30Days")
    .addItem("🗂️ Full Refresh (All Jobs)", "fullRefresh")
    //.addSeparator()
    //.addItem("🏷️ Generate Short Labels",  "generateShortLabels")
    .addSeparator()
    .addItem("💾 Backup Data",            "backupData")
    .addItem("📂 Load Backup", "loadBackup")
    .addSeparator()
    .addItem("📋 List Backups", "listBackups")
    //.addItem("🗑️ Clear Data",             "clearData")
    //.addSeparator()
    //.addItem("🔧 Setup Sheets (First Run)", "setupSheets")
    .addToUi();
}


// ============================================================
// 1. SETUP
//    Run once on first use. Creates both sheets with all
//    formatting, dropdowns, and the Audit Log.
//    Safe to re-run — prompts for confirmation first.
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const woExists  = !!ss.getSheetByName(CONFIG.sheets.workOrders);
  const matExists = !!ss.getSheetByName(CONFIG.sheets.jobMaterials);

  if (woExists || matExists) {
    const ok = ui.alert(
      "Setup Warning",
      "Sheets already exist. Setup will clear and recreate them.\n\nBackup first if you have data to keep.\n\nContinue?",
      ui.ButtonSet.YES_NO
    );
    if (ok !== ui.Button.YES) return;
  }

  _createWorkOrdersSheet(ss);
  _createJobMaterialsSheet(ss);
  _ensureAuditLogSheet(ss);

  ui.alert("Setup complete!\n\nNext:\n1. Run Full Refresh to load all jobs.\n2. Labels generate automatically during refresh.");
}


// ============================================================
// 2. REFRESH — Menu triggers
//    Each calls the core engine with a day window.
//    null = no date filter (full refresh).
// ============================================================
function refreshLast7Days()  { _runRefresh(7);    }
function refreshLast30Days() { _runRefresh(30);   }
function fullRefresh() {
  const ok = SpreadsheetApp.getUi().alert(
    "Full Refresh",
    "This pulls all active Work Orders with no date limit. May take longer.\n\nContinue?",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (ok === SpreadsheetApp.getUi().Button.YES) _runRefresh(null);
}


// ============================================================
// 3. CORE REFRESH ENGINE
//    Fetches SM8 data, syncs both sheets, removes stale rows,
//    excludes invoice items, then generates labels.
// ============================================================
function _runRefresh(days) {
  _archiveAuditLogIfNeeded();   // for archiving old audit log entries
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sm8Key = PropertiesService.getScriptProperties().getProperty(CONFIG.sm8.scriptPropKey);

  if (!sm8Key) {
    SpreadsheetApp.getUi().alert("SM8_API_KEY not found in Script Properties.");
    return;
  }

  const woSheet  = ss.getSheetByName(CONFIG.sheets.workOrders);
  const matSheet = ss.getSheetByName(CONFIG.sheets.jobMaterials);

  if (!woSheet || !matSheet) {
    SpreadsheetApp.getUi().alert("Sheets not found. Run Setup first.");
    return;
  }

  // --- Fetch from SM8 ---
  const companyMap   = _fetchCompanyMap(sm8Key);
  const jobs         = _fetchJobs(sm8Key, days);
  if (!jobs) return; // Error already alerted inside _fetchJobs

  const rawMaterials = _fetchMaterials(sm8Key);
  const materials    = _filterExcludedItems(rawMaterials);

  const activeJobUuids = new Set(jobs.map(j => j.uuid));

  // Build a lookup of job info for enriching material rows
  const jobMap = {};
  jobs.forEach(job => {
    jobMap[job.uuid] = {
      jobNo:       job.generated_job_id || "",
      clientName:  companyMap[job.company_uuid] || "Unknown",
      address:     job.job_address || "TBC",
      dateCreated: job.date ? job.date.split(" ")[0] : "",
      shortDesc:   (job.job_description || "").split("\n")[0].trim(),
    };
  });

  // --- Sync sheets ---
  const woResult  = _syncWorkOrdersSheet(woSheet, jobs, jobMap, activeJobUuids);
  const matResult = _syncJobMaterialsSheet(matSheet, materials, jobMap, activeJobUuids, jobs);

  // --- Generate labels for any new unlabelled items ---
  generateShortLabels();

  // --- Save last sync timestamp ---
  PropertiesService.getScriptProperties().setProperty("LAST_SYNC", new Date().toLocaleString());

  ss.toast(
    `Work Orders: ${woResult.synced} synced, ${woResult.removed} removed | ` +
    `Materials: ${matResult.synced} synced, ${matResult.removed} removed | ` +
    `${new Date().toLocaleTimeString()}`,
    "✅ Sync Complete", 6
  );
}


// ============================================================
// 4. SHEET SYNC — Work Orders
//    Updates SM8 columns (A–E) on existing rows.
//    Appends new rows. Deletes rows no longer in SM8.
//    Never touches col F (Status) — that is manual.
// ============================================================
function _syncWorkOrdersSheet(sheet, jobs, jobMap, activeJobUuids) {
  const WO      = CONFIG.workOrders;
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, WO.totalCols).getValues()
    : [];

  // UUID → row index in existing array
  const existingMap = {};
  existing.forEach((row, i) => {
    const uuid = row[WO.uuid - 1];
    if (uuid) existingMap[uuid] = i;
  });

  let synced = 0;
  let removed = 0;

  // Update or append
  jobs.forEach(job => {
    const uuid = job.uuid;
    if (!uuid) return;

    const info      = jobMap[uuid];
    const sm8Values = [info.jobNo, info.clientName, info.address, info.shortDesc, info.dateCreated];

    if (uuid in existingMap) {
      sheet.getRange(existingMap[uuid] + 2, 1, 1, WO.sm8ColCount).setValues([sm8Values]);
    } else {
      sheet.appendRow([...sm8Values, "", uuid]);
    }
    synced++;
  });

  // Remove stale rows — loop bottom to top to avoid row index shifting
  for (let i = existing.length - 1; i >= 0; i--) {
    const uuid = existing[i][WO.uuid - 1];
    if (uuid && !activeJobUuids.has(uuid)) {
      sheet.deleteRow(i + 2);
      removed++;
    }
  }

  // Sort: Client Name A→Z, Date Created newest first
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, WO.totalCols)
      .sort([{ column: WO.clientName, ascending: true }, { column: WO.dateCreated, ascending: false }]);
  }

  return { synced, removed };
}


// ============================================================
// 5. SHEET SYNC — Job Materials
//    Updates SM8 columns (A–F) on existing rows.
//    Appends new rows. Deletes rows for inactive/removed items.
//    Never touches col G (label) or H–Q (manual columns).
// ============================================================
function _syncJobMaterialsSheet(matSheet, materials, jobMap, activeJobUuids, jobs) {
  const MAT     = CONFIG.jobMaterials;
  const lastRow = matSheet.getLastRow();
  const existing = lastRow > 1
    ? matSheet.getRange(2, 1, lastRow - 1, MAT.totalCols).getValues()
    : [];

  const existingMap = {};
  existing.forEach((row, i) => {
    const uuid = row[MAT.uuid - 1];
    if (uuid) existingMap[uuid] = i;
  });

  const activeMatUuids    = new Set();
  const jobsWithMaterials = new Set();
  let synced  = 0;
  let removed = 0;

  // Sync material rows
  materials.forEach(mat => {
    if (!activeJobUuids.has(mat.job_uuid)) return;
    jobsWithMaterials.add(mat.job_uuid);

    const matUuid = mat.uuid || "";
    if (matUuid) activeMatUuids.add(matUuid);

    const info      = jobMap[mat.job_uuid] || {};
    const sm8Values = [
      info.jobNo || "", info.clientName || "", info.address || "",
      info.dateCreated || "", info.shortDesc || "", mat.name || "",
    ];

    if (matUuid && matUuid in existingMap) {
      // A–F only — col G (label) and H–Q (manual) are never overwritten
      matSheet.getRange(existingMap[matUuid] + 2, 1, 1, MAT.sm8ColCount).setValues([sm8Values]);
    } else {
      // New row — 12 empty placeholders for G through Q, then UUID
      matSheet.appendRow([...sm8Values, "", "", "", "", "", "", "", "", "", "", "", matUuid]);
    }
    synced++;
  });

  // Placeholder row for jobs with no materials yet
  jobs.forEach(job => {
    if (jobsWithMaterials.has(job.uuid)) return;
    const info    = jobMap[job.uuid] || {};
    const fakeKey = "job_" + job.uuid;
    activeMatUuids.add(fakeKey);

    const sm8Values = [
      info.jobNo || "", info.clientName || "", info.address || "",
      info.dateCreated || "", info.shortDesc || "", "",
    ];

    if (fakeKey in existingMap) {
      matSheet.getRange(existingMap[fakeKey] + 2, 1, 1, MAT.sm8ColCount).setValues([sm8Values]);
    } else {
      matSheet.appendRow([...sm8Values, "", "", "", "", "", "", "", "", "", "", "", fakeKey]);
    }
  });

  // Remove stale rows — bottom to top
  for (let i = existing.length - 1; i >= 0; i--) {
    const uuid = existing[i][MAT.uuid - 1];
    if (uuid && !activeMatUuids.has(uuid)) {
      matSheet.deleteRow(i + 2);
      removed++;
    }
  }

  // Sort: Client Name A→Z, Date Created newest first
  if (matSheet.getLastRow() > 1) {
    matSheet.getRange(2, 1, matSheet.getLastRow() - 1, MAT.totalCols)
      .sort([{ column: MAT.clientName, ascending: true }, { column: MAT.dateCreated, ascending: false }]);
  }

  return { synced, removed };
}


// ============================================================
// 6. GENERATE SHORT LABELS
//    Runs automatically at end of every refresh.
//    Only processes rows where col G is blank.
//    To redo one label: clear that cell in G, then refresh.
// ============================================================
function generateShortLabels() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const matSheet     = ss.getSheetByName(CONFIG.sheets.jobMaterials);
  const anthropicKey = PropertiesService.getScriptProperties().getProperty(CONFIG.ai.scriptPropKey);
  const MAT          = CONFIG.jobMaterials;

  if (!anthropicKey) {
    SpreadsheetApp.getUi().alert("ANTHROPIC_API_KEY not found in Script Properties.");
    return;
  }

  const lastRow = matSheet.getLastRow();
  if (lastRow < 2) return;

  const data = matSheet.getRange(2, 1, lastRow - 1, MAT.shortLabel).getValues();
  let processed = 0;
  let skipped   = 0;

  data.forEach((row, i) => {
    const rawItem      = row[MAT.itemName - 1];    // col F
    const existingLabel = row[MAT.shortLabel - 1]; // col G
    const sheetRow     = i + 2;

    // Skip if no item name or label already exists
    if (!rawItem || existingLabel) { skipped++; return; }

    try {
      const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type":      "application/json",
          "x-api-key":         anthropicKey,
          "anthropic-version": "2023-06-01",
        },
        payload: JSON.stringify({
          model:      CONFIG.ai.model,
          max_tokens: CONFIG.ai.maxTokens,
          messages:   [{ role: "user", content: _buildLabelPrompt(rawItem) }],
        }),
        muteHttpExceptions: true,
      });

      const result = JSON.parse(response.getContentText());
      const label  = result.content[0].text.trim();

      const cell = matSheet.getRange(sheetRow, MAT.shortLabel);
      cell.setValue(label);
      cell.setNote(rawItem.toString()); // Full item name shows on hover
      cell.setWrap(true);
      processed++;

      Utilities.sleep(300); // Avoid API rate limits

    } catch (e) {
      Logger.log("Label error row " + sheetRow + ": " + e.message);
    }
  });

  if (processed > 0) {
    ss.toast(`Generated: ${processed} | Skipped: ${skipped}`, "🏷️ Labels Done", 5);
  }
}


// ============================================================
// 7. BACKUP
//    Copies both sheets to hidden timestamped sheets.
//    To restore: right-click the tab → Show Sheet.
// ============================================================
function backupData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const label = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm");
  const MAX_BACKUPS = 5;

   [CONFIG.sheets.workOrders, CONFIG.sheets.jobMaterials].forEach(name => {
    const source = ss.getSheetByName(name);
    if (!source) return;
    const copy = source.copyTo(ss);
    copy.setName("Backup — " + label + " | " + name);
    copy.hideSheet();
  });

  // Get all backups sorted oldest first, delete if over limit
  const allBackups = ss.getSheets()
    .filter(s => s.getName().startsWith("Backup —"))
    .sort((a, b) => a.getName().localeCompare(b.getName()));

  const uniqueDates = [...new Set(allBackups.map(s => s.getName().split(" | ")[0]))];

  if (uniqueDates.length > MAX_BACKUPS) {
    const oldest = uniqueDates[0];
    allBackups
      .filter(s => s.getName().startsWith(oldest))
      .forEach(s => ss.deleteSheet(s));
    Logger.log("Deleted old backup: " + oldest);
  }

  ss.toast("Backup saved. Keeping latest " + MAX_BACKUPS + " backups.", "💾 Backup Done", 5);
}


// ============================================================
// 8. CLEAR DATA
//    Prompts for backup first. Then deletes all data rows
//    from both sheets, leaving only the header row.
// ============================================================
function clearData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const backup = ui.alert("Backup First?", "Create a backup before clearing?", ui.ButtonSet.YES_NO);
  if (backup === ui.Button.YES) backupData();

  const confirm = ui.alert(
    "Clear All Data",
    "This permanently deletes all rows from both sheets including manual columns.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  _deleteAllDataRows(ss.getSheetByName(CONFIG.sheets.workOrders));
  _deleteAllDataRows(ss.getSheetByName(CONFIG.sheets.jobMaterials));

  ss.toast("Both sheets cleared.", "✅ Done", 3);
}


// ============================================================
// 9. AUDIT LOG — Tracks all manual edits
//    Fires on every cell edit in manual columns.
//    Locked sheet — no manual editing of log entries.
// ============================================================
function onEdit(e) {
  if (!e || !e.source) return;

  const sheet     = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const col       = e.range.getColumn();
  const row       = e.range.getRow();

  if (row === 1) return; // Ignore header row

  const isWOManual  = sheetName === CONFIG.sheets.workOrders  && col === CONFIG.workOrders.status;
  const isMatManual = sheetName === CONFIG.sheets.jobMaterials && col >= CONFIG.jobMaterials.installer;

  if (!isWOManual && !isMatManual) return;

  // "Other" handler — clears cell and prompts free text
  if (sheetName === CONFIG.sheets.jobMaterials) {
    const dropdownCols = [
      CONFIG.jobMaterials.installer, CONFIG.jobMaterials.productType,
      CONFIG.jobMaterials.designStatus, CONFIG.jobMaterials.glassStatus,
      CONFIG.jobMaterials.hardwareStatus, CONFIG.jobMaterials.overallStatus,
    ];
    if (dropdownCols.includes(col) && e.range.getValue() === "Other") {
      e.range.clearContent();
      e.source.toast("Type your custom value in the cell.", "✏️ Other selected", 3);
    }
  }

  // Write to Audit Log
  _writeAuditLog(e, sheetName, row, col, sheet);
}


// ============================================================
// SM8 FETCH HELPERS
// ============================================================

function _fetchCompanyMap(apiKey) {
  const map = {};
  try {
    const res = UrlFetchApp.fetch(CONFIG.sm8.baseUrl + "/company.json",
      { method: "GET", headers: { "X-API-Key": apiKey }, muteHttpExceptions: true });
    JSON.parse(res.getContentText()).forEach(c => {
      if (c.uuid) map[c.uuid] = c.name || "Unknown";
    });
  } catch (e) { Logger.log("Company fetch failed: " + e.message); }
  return map;
}

function _fetchJobs(apiKey, days) {
  let filter = "status eq 'Work Order' and active eq 1";
  if (days) filter += " and date gt '" + _getDateDaysAgo(days) + "'";
  try {
    const res = UrlFetchApp.fetch(
      CONFIG.sm8.baseUrl + "/job.json?$filter=" + encodeURIComponent(filter),
      { method: "GET", headers: { "X-API-Key": apiKey }, muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert("SM8 API error: " + res.getContentText());
      return null;
    }
    return JSON.parse(res.getContentText());
  } catch (e) {
    SpreadsheetApp.getUi().alert("SM8 connection failed: " + e.message);
    return null;
  }
}

function _fetchMaterials(apiKey) {
  try {
    const res = UrlFetchApp.fetch(
      CONFIG.sm8.baseUrl + "/jobmaterial.json?$filter=active eq 1",
      { method: "GET", headers: { "X-API-Key": apiKey }, muteHttpExceptions: true }
    );
    return JSON.parse(res.getContentText());
  } catch (e) {
    Logger.log("Materials fetch failed: " + e.message);
    return [];
  }
}

function _filterExcludedItems(materials) {
  const keywords = CONFIG.excludeKeywords;
  return materials.filter(mat => {
    const name     = (mat.name || "").toLowerCase().trim();
    const excluded = keywords.some(kw => name.includes(kw));
    if (excluded) Logger.log("EXCLUDED: " + mat.name);
    return !excluded;
  });
}


// ============================================================
// SHEET CREATION HELPERS (called by setupSheets)
// ============================================================

function _createWorkOrdersSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.sheets.workOrders);
  if (!sheet) sheet = ss.insertSheet(CONFIG.sheets.workOrders);
  else sheet.clear();

  const WO = CONFIG.workOrders;
  const T  = CONFIG.theme;

  const headers = ["Job No", "Client Name", "Site Address", "Project Description", "Date Created", "Status", "Item ID"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(T.headerBg).setFontColor(T.headerFont).setFontWeight("bold").setFontSize(11);

  sheet.setFrozenRows(1);
  sheet.hideColumns(WO.uuid);
  sheet.getBandings().forEach(b => b.remove());
  sheet.getRange(1, 1, sheet.getMaxRows(), WO.totalCols)
    .applyRowBanding()
    .setHeaderRowColor(T.headerBg)
    .setFirstRowColor("#ffffff")
    .setSecondRowColor("#eef3f8");

  const rows = Math.max(sheet.getMaxRows(), 100);
  sheet.getRange(2, WO.status, rows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.dropdowns.woStatus, true)
      .setAllowInvalid(false)
      .build()
  );
}

function _createJobMaterialsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.sheets.jobMaterials);
  if (!sheet) sheet = ss.insertSheet(CONFIG.sheets.jobMaterials);
  else sheet.clear();

  const MAT = CONFIG.jobMaterials;
  const T   = CONFIG.theme;
  const D   = CONFIG.dropdowns;

  const headers = [
    "Job No", "Client Name", "Site Address", "Date Created", "Project Description",
    "Item Name", "Item - Short Label", "Installer", "Product Type", "Design Status",
    "Expected Completion Date", "Confirmed Date", "Glass Status", "Hardware Status",
    "Site Ready", "Overall Status", "Remarks", "Item UUID"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(T.headerBg).setFontColor(T.headerFont).setFontWeight("bold").setFontSize(11);

  sheet.setFrozenRows(1);
  sheet.hideColumns(MAT.itemName); // Hidden — raw SM8 name (col F)
  sheet.hideColumns(MAT.uuid);     // Hidden — primary key (col R)
  sheet.setColumnWidth(MAT.shortLabel, 320);

  sheet.getBandings().forEach(b => b.remove());
  sheet.getRange(1, 1, sheet.getMaxRows(), MAT.totalCols)
    .applyRowBanding()
    .setHeaderRowColor(T.headerBg)
    .setFirstRowColor("#ffffff")
    .setSecondRowColor("#eef3f8");

  const rows     = Math.max(sheet.getMaxRows(), 100);
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();

  function dropdown(col, options) {
    sheet.getRange(2, col, rows - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList([...options, "Other"], true)
        .setAllowInvalid(true)
        .build()
    );
  }

  dropdown(MAT.installer,      D.installer);
  dropdown(MAT.productType,    D.productType);
  dropdown(MAT.designStatus,   D.designStatus);
  sheet.getRange(2, MAT.expectedDate,  rows - 1, 1).setDataValidation(dateRule).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(2, MAT.confirmedDate, rows - 1, 1).setDataValidation(dateRule).setNumberFormat("yyyy-mm-dd");
  dropdown(MAT.glassStatus,    D.glassStatus);
  dropdown(MAT.hardwareStatus, D.hardwareStatus);
  sheet.getRange(2, MAT.siteReady, rows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );
  dropdown(MAT.overallStatus, D.overallStatus);
}

function _ensureAuditLogSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.sheets.auditLog);
  if (sheet) return; // Already exists — don't recreate

  sheet = ss.insertSheet(CONFIG.sheets.auditLog);
  sheet.getRange(1, 1, 1, 8).setValues([[
    "Timestamp", "Edited By", "Sheet", "Job No", "Client Name", "Column", "Old Value", "New Value"
  ]]).setBackground(CONFIG.theme.headerBg).setFontColor(CONFIG.theme.headerFont).setFontWeight("bold");
  sheet.setFrozenRows(1);

  // Warning-only protection — prevents accidental edits, not malicious ones
  sheet.protect().setDescription("Audit Log — do not edit manually").setWarningOnly(true);
}


// ============================================================
// GENERAL HELPERS
// ============================================================

function _getDateDaysAgo(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return d.toISOString().split("T")[0];
}

function _deleteAllDataRows(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;
  const lastRow = sheet.getLastRow();
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  if (lastRow > 2) {
    sheet.deleteRows(3, lastRow - 2);
  }
}

function _writeAuditLog(e, sheetName, row, col, sheet) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName(CONFIG.sheets.auditLog);
  if (!auditSheet) return;

  const colLetter = String.fromCharCode(64 + col);
  const jobNo     = sheet.getRange(row, 1).getValue();
  const client    = sheet.getRange(row, 2).getValue();
  const colName   = _getColumnName(sheetName, col);

  auditSheet.appendRow([
    new Date(),
    Session.getActiveUser().getEmail(),
    sheetName,
    jobNo,
    client,
    colLetter + " — " + colName,
    e.oldValue || "",
    e.value || e.range.getValue(),
  ]);
}

function _getColumnName(sheetName, col) {
  if (sheetName === CONFIG.sheets.workOrders) {
    return { 6: "Status" }[col] || "Col " + col;
  }
  if (sheetName === CONFIG.sheets.jobMaterials) {
    return {
      8: "Installer", 9: "Product Type", 10: "Design Status",
      11: "Expected Date", 12: "Confirmed Date", 13: "Glass Status",
      14: "Hardware Status", 15: "Site Ready", 16: "Overall Status", 17: "Remarks",
    }[col] || "Col " + col;
  }
  return "Col " + col;
}

function _buildLabelPrompt(rawItem) {
  return `You are labelling glass installation job items for Royal Glass NZ production dashboard.

Return one line per product scope using this format:
[Short Name] | L: [length] | H: [height] | [Finish]

RULES:

Short Name
- Product type in 2-3 words
- Include quantity with × if multiple identical units (e.g. Shower Glass ×8)

Dimensions
- L: is total run length. H: is height above floor
- Use only the dimension that applies — you do not need both
- Copy numbers exactly as written. Do NOT convert units. 14.6m stays 14.6m. 1200mm stays 1200mm
- Two segments of the same product: combine with + (e.g. L: 5.3m + 8.4m)
- Many individual panel sizes with a total area: use the total area instead (e.g. 11.6sqm)
- Unknown dimension: write TBC

Finish
- e.g. Chrome, Black, Polished, Satin, Matt Black, Powder Coated, Gun Metal
- Unknown: write TBC

Multiple scopes
- If the raw description contains more than one distinct product, return one line per product
- Each line follows the same format

Examples:

Single scope:
"Mini Post System 14.6m chrome internal stair timber"
→ Mini Post System | L: 14.6m | H: 1000mm | Chrome

Many panel sizes with total area:
"Glass Splashback 3000x770, 4400x770 — 11.6sqm polished"
→ Glass Splashback | 11.6sqm | Polished

Two segments:
"Duo Double Anchor stair 5.3m + landing 8.4m, 1m high, black"
→ Duo Double Anchor System | L: 5.3m + 8.4m | H: 1m | Black

Multiple scopes:
"Shower Glass 2 hinged 1960mm chrome + Stair Handrail 6.5m stainless"
→ Shower Glass ×2 | H: 1960mm | Chrome
→ Stair Handrail | L: 6.5m | Chrome

Unknown dimensions:
"replace small glass above shower — 3 showers"
→ Shower Glass ×3 | TBC | TBC

Raw description:
${rawItem.toString().substring(0, 600)}

Return the label only. No explanation. No markdown.`;
}


// ============================================================
// UTILITY — Run manually when needed
// ============================================================

// Backfills hover notes on existing labels (run once after migration)
function backfillNotes() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(CONFIG.sheets.jobMaterials);
  const MAT     = CONFIG.jobMaterials;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, MAT.itemName, lastRow - 1, 2).getValues();
  data.forEach((row, i) => {
    const rawItem = row[0]; // col F
    const label   = row[1]; // col G
    if (label && rawItem) sheet.getRange(i + 2, MAT.shortLabel).setNote(rawItem.toString());
  });

  ss.toast("Notes backfilled.", "✅ Done", 3);
}

function listBackups() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const backups = ss.getSheets()
    .filter(s => s.getName().startsWith("Backup —"))
    .map(s => s.getName());

  if (backups.length === 0) {
    SpreadsheetApp.getUi().alert("No backups found.");
    return;
  }

  SpreadsheetApp.getUi().alert(
    "Existing Backups (" + backups.length + ")",
    backups.join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function _archiveAuditLogIfNeeded() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName(CONFIG.sheets.auditLog);
  if (!auditSheet || auditSheet.getLastRow() < 2) return;

  const cutoff  = new Date();
  cutoff.setDate(cutoff.getDate() - 30);

  const data     = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 8).getValues();
  const toKeep   = [];
  const toArchive = [];

  data.forEach(row => {
    const ts = new Date(row[0]);
    (ts < cutoff ? toArchive : toKeep).push(row);
  });

  if (toArchive.length === 0) return;

  // Write archive
  const monthLabel  = Utilities.formatDate(cutoff, Session.getScriptTimeZone(), "MMM yyyy");
  const archiveName = "Audit Archive — " + monthLabel;
  let archiveSheet  = ss.getSheetByName(archiveName);

  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(archiveName);
    archiveSheet.getRange(1, 1, 1, 8).setValues([[
      "Timestamp", "Edited By", "Sheet", "Job No", "Client Name", "Column", "Old Value", "New Value"
    ]]).setBackground(CONFIG.theme.headerBg).setFontColor(CONFIG.theme.headerFont).setFontWeight("bold");
    archiveSheet.setFrozenRows(1);
    archiveSheet.hideSheet();
  }

  const archiveLastRow = archiveSheet.getLastRow();
  archiveSheet.getRange(archiveLastRow + 1, 1, toArchive.length, 8).setValues(toArchive);

  // Rewrite active log with only recent rows
  auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 8).clearContent();
  if (toKeep.length > 0) {
    auditSheet.getRange(2, 1, toKeep.length, 8).setValues(toKeep);
  }

  Logger.log("Archived " + toArchive.length + " audit rows to " + archiveName);
}

function deleteAllBackups() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const ui      = SpreadsheetApp.getUi();
  const backups = ss.getSheets().filter(s => s.getName().startsWith("Backup —"));

  if (backups.length === 0) { ui.alert("No backups found."); return; }

  const confirm = ui.alert(
    "Delete All Backups",
    `This will permanently delete ${backups.length} backup sheet(s). Cannot be undone.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  backups.forEach(s => ss.deleteSheet(s));
  ss.toast(`Deleted ${backups.length} backup(s).`, "🗑️ Done", 3);
}