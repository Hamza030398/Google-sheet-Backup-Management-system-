/**************************************************
 * BACKUP MANAGER — HOURLY + WORKING HOURS + SETTINGS SHEET
 **************************************************/


const TIMEZONE = "Asia/Karachi";


/**************************************************
 * ON OPEN: Create Menu
 **************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🛡 Backup Manager")
    .addItem("Restore Latest Backup", "restoreLatestBackup")
    .addItem("Open Settings Sheet", "openSettingsSheet")
    .addSeparator()
    .addItem("Reinstall Backup System", "installBackupSystem")
    .addToUi();
}


/**************************************************
 * Create default Settings sheet if missing
 **************************************************/
function getSettings() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Backup_Settings");


  // If missing → create AND initialize safely
  if (!sh) {
    sh = ss.insertSheet();
    sh.setName("Backup_Settings");


    const defaults = [
      ["Setting", "Value"],
      ["SOURCE_SHEET_NAME", "Data"],
      ["WORK_START_HOUR", 8],
      ["WORK_END_HOUR", 17],
      ["MAX_BACKUPS", 6],
      ["MIN_CELL_CHANGES", 10]
    ];


    // IMPORTANT: Never write out-of-grid range → use default size
    sh.getRange(1, 1, defaults.length, 2).setValues(defaults);
  }


  // Read settings into object
  const settings = {};
  const data = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();


  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];


    if (key) settings[key] = value;
  }


  return settings;
}




/**************************************************
 * Install HOURLY trigger
 **************************************************/
function installBackupSystem() {
  deleteAllBackupTriggers();


  ScriptApp.newTrigger("hourlyBackupController")
    .timeBased()
    .everyHours(1)
    .create();


  logEvent("SYSTEM", "Installed hourly backup triggers");
  SpreadsheetApp.getUi().alert("Hourly Backup System Installed Successfully!");
}


/**************************************************
 * Delete old backup triggers
 **************************************************/
function deleteAllBackupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "hourlyBackupController") {
      ScriptApp.deleteTrigger(t);
    }
  });
}


/**************************************************
 * MAIN CONTROLLER — Runs every hour
 * Checks time window before allowing backup
 **************************************************/
function hourlyBackupController() {
  const settings = getSettings();


  const start = Number(settings.WORK_START_HOUR);
  const end = Number(settings.WORK_END_HOUR);


  const now = new Date();
  const hourNow = Number(Utilities.formatDate(now, TIMEZONE, "HH"));


  // Outside working hours → skip
  if (hourNow < start || hourNow >= end) {
    logEvent("SKIPPED", `Outside working hours (${start}:00 - ${end}:00)`);
    return;
  }


  // Within working hours → create backup
  createBackup();
}


/**************************************************
 * CREATE BACKUP
 **************************************************/
function createBackup() {
  try {
    const settings = getSettings();
    const ss = SpreadsheetApp.getActive();
    const sourceName = settings.SOURCE_SHEET_NAME;
    const original = ss.getSheetByName(sourceName);


    if (!original) {
      logEvent("ERROR", `Source sheet '${sourceName}' not found.`);
      return;
    }


    const now = new Date();
    const tag = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd_HH:mm");
    const backupName = `Backup_${sourceName}_${tag}`;


    // Create backup
    const newSheet = original.copyTo(ss);
    newSheet.setName(backupName);


    logEvent("SUCCESS", `Backup created: ${backupName}`);


    // Enforce retention
    cleanupOldBackups();


  } catch (err) {
    logEvent("ERROR", err.message);
  }
}


/**************************************************
 * RETENTION — Keep last X backups
 **************************************************/
function cleanupOldBackups() {
  const settings = getSettings();
  const MAX_BACKUPS = Number(settings.MAX_BACKUPS);
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  const sourceName = settings.SOURCE_SHEET_NAME;


  const backups = sheets.filter(sh =>
    sh.getName().startsWith(`Backup_${sourceName}_`)
  );


  // Sort newest → oldest using timestamp in sheet name
  backups.sort((a, b) => {
    const aTime = parseBackupTimestamp(a.getName(), sourceName);
    const bTime = parseBackupTimestamp(b.getName(), sourceName);
    return bTime - aTime; // newest first
  });


  // Delete old backups beyond MAX_BACKUPS
  if (backups.length > MAX_BACKUPS) {
    const toDelete = backups.slice(MAX_BACKUPS);
    toDelete.forEach(sh => {
      logEvent("INFO", `Deleting old backup: ${sh.getName()}`);
      ss.deleteSheet(sh);
    });
  }
}
function parseBackupTimestamp(name, sourceName) {
  try {
    const parts = name.replace(`Backup_${sourceName}_`, '');
    // Format: yyyy-MM-dd_HH:mm
    const dt = Utilities.parseDate(parts, TIMEZONE, "yyyy-MM-dd_HH:mm");
    return dt.getTime();
  } catch (e) {
    return 0; // fallback: treat invalid names as very old
  }
}
/**************************************************
 * RESTORE LATEST BACKUP
 **************************************************/
function restoreLatestBackup() {
  const settings = getSettings();
  const sourceName = settings.SOURCE_SHEET_NAME;


  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();


  const backups = sheets.filter(sh =>
    sh.getName().startsWith(`Backup_${sourceName}`)
  );


  if (backups.length === 0) {
    SpreadsheetApp.getUi().alert("No backups available.");
    return;
  }


  // Sort newest first (by name timestamp)
  backups.sort((a, b) => b.getName().localeCompare(a.getName()));


  const latest = backups[0];


  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    "Restore Backup",
    `This will replace '${sourceName}' with '${latest.getName()}'. Continue?`,
    ui.ButtonSet.YES_NO
  );


  if (confirm !== ui.Button.YES) return;


  const original = ss.getSheetByName(sourceName);
  if (original) ss.deleteSheet(original);


  const restored = latest.copyTo(ss);
  restored.setName(sourceName);


  ui.alert(`Restored from: ${latest.getName()}`);
  logEvent("RESTORE", `Restored from: ${latest.getName()}`);
}


/**************************************************
 * Logging
 **************************************************/
function logEvent(type, message) {
  const ss = SpreadsheetApp.getActive();
  let log = ss.getSheetByName("Backup_Log");


  if (!log) {
    log = ss.insertSheet("Backup_Log");
    log.appendRow(["Timestamp", "Type", "Message"]);
  }


  log.appendRow([
    Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss"),
    type,
    message
  ]);
}


/**************************************************
 * Open Settings Sheet
 **************************************************/
function openSettingsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Backup_Settings");


  if (!sh) {
    // Call getSettings() to auto-create it
    getSettings();
    sh = ss.getSheetByName("Backup_Settings");
  }


  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert("Backup Settings sheet is ready.");
}








For Backup in the other spreadsheet
/**************************************************
 * BACKUP MANAGER — HOURLY + WORKING HOURS + REPOSITORY (Option C)
 **************************************************/


const TIMEZONE = "Asia/Karachi";


/**************************************************
 * ON OPEN: Create Menu
 **************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🛡 Backup Manager")
    .addItem("Restore Latest Backup", "restoreLatestBackup")
    .addItem("Open Settings Sheet", "openSettingsSheet")
    .addSeparator()
    .addItem("Reinstall Backup System", "installBackupSystem")
    .addToUi();
}


/**************************************************
 * Create default Settings sheet if missing
 **************************************************/
function getSettings() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Backup_Settings");


  // If missing → create AND initialize safely
  if (!sh) {
    sh = ss.insertSheet();
    sh.setName("Backup_Settings");


    const defaults = [
      ["Setting", "Value"],
      ["SOURCE_SHEET_NAME", "Data"],
      ["WORK_START_HOUR", 8],
      ["WORK_END_HOUR", 17],
      ["MAX_BACKUPS", 6],
      ["MIN_CELL_CHANGES", 10],
      ["BACKUP_SPREADSHEET_ID", ""] // empty = use current spreadsheet
    ];


    sh.getRange(1, 1, defaults.length, 2).setValues(defaults);
  }


  // Read settings into object
  const settings = {};
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();


  data.forEach(row => {
    const key = row[0];
    const value = row[1];
    if (key) settings[key] = value;
  });


  // Normalize numeric settings
  settings.WORK_START_HOUR = Number(settings.WORK_START_HOUR) || 8;
  settings.WORK_END_HOUR = Number(settings.WORK_END_HOUR) || 17;
  settings.MAX_BACKUPS = Number(settings.MAX_BACKUPS) || 6;
  settings.MIN_CELL_CHANGES = Number(settings.MIN_CELL_CHANGES) || 10;
  settings.SOURCE_SHEET_NAME = settings.SOURCE_SHEET_NAME || "Data";
  settings.BACKUP_SPREADSHEET_ID = settings.BACKUP_SPREADSHEET_ID || "";


  return settings;
}


/**************************************************
 * Install HOURLY trigger
 **************************************************/
function installBackupSystem() {
  deleteAllBackupTriggers();


  ScriptApp.newTrigger("hourlyBackupController")
    .timeBased()
    .everyHours(1)
    .create();


  logEvent("SYSTEM", "Installed hourly backup trigger");
  SpreadsheetApp.getUi().alert("Hourly Backup System Installed Successfully!");
}


/**************************************************
 * Delete old backup triggers
 **************************************************/
function deleteAllBackupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "hourlyBackupController") {
      ScriptApp.deleteTrigger(t);
    }
  });
}


/**************************************************
 * MAIN CONTROLLER — Runs every hour
 * Checks time window before allowing backup
 **************************************************/
function hourlyBackupController() {
  const settings = getSettings();


  const start = Number(settings.WORK_START_HOUR);
  const end = Number(settings.WORK_END_HOUR);


  const now = new Date();
  const hourNow = Number(Utilities.formatDate(now, TIMEZONE, "HH"));


  // Outside working hours → skip
  if (hourNow < start || hourNow >= end) {
    logEvent("SKIPPED", `Outside working hours (${start}:00 - ${end}:00)`);
    return;
  }


  // Within working hours → create backup
  createBackup();
}


/**************************************************
 * CREATE BACKUP (writes to repository when provided)
 **************************************************/
function createBackup() {
  try {
    const settings = getSettings();
    const activeSS = SpreadsheetApp.getActive();
    const sourceName = settings.SOURCE_SHEET_NAME;
    const original = activeSS.getSheetByName(sourceName);


    if (!original) {
      logEvent("ERROR", `Source sheet '${sourceName}' not found.`);
      return;
    }


    const now = new Date();
    const tag = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd_HH:mm");
    const backupName = `Backup_${sourceName}_${tag}`;


    // Determine target spreadsheet (repository) — Option C: use repository if provided
    let targetSS;
    if (settings.BACKUP_SPREADSHEET_ID && String(settings.BACKUP_SPREADSHEET_ID).trim()) {
      try {
        targetSS = SpreadsheetApp.openById(String(settings.BACKUP_SPREADSHEET_ID).trim());
      } catch (e) {
        logEvent("ERROR", `Invalid BACKUP_SPREADSHEET_ID: ${settings.BACKUP_SPREADSHEET_ID}`);
        targetSS = activeSS;
      }
    } else {
      targetSS = activeSS;
    }


    // If same spreadsheet target and name collision: delete existing
    const existingInTarget = targetSS.getSheetByName(backupName);
    if (existingInTarget) {
      targetSS.deleteSheet(existingInTarget);
    }


    // Copy to target spreadsheet
    const copied = original.copyTo(targetSS);
    // setName may throw if duplicate or invalid; ensure unique
    let finalName = backupName;
    if (targetSS.getSheetByName(finalName)) {
      // append seconds to make unique
      const suffix = Utilities.formatDate(new Date(), TIMEZONE, "HHmmss");
      finalName = `${backupName}_${suffix}`;
    }
    copied.setName(finalName);


    logEvent("SUCCESS", `Backup created: ${finalName} (to ${targetSS.getName()})`);


    // Run retention cleanup on the target spreadsheet
    cleanupOldBackups(targetSS);


  } catch (err) {
    logEvent("ERROR", err.message || String(err));
  }
}








/**************************************************
 * RETENTION — Keep last X backups (runs on targetSS)
 **************************************************/
function cleanupOldBackups(targetSS) {
  const settings = getSettings();
  const MAX_BACKUPS = Number(settings.MAX_BACKUPS);
  const sourceName = settings.SOURCE_SHEET_NAME;


  if (!targetSS) targetSS = SpreadsheetApp.getActive();


  const sheets = targetSS.getSheets();


  const backups = sheets.filter(sh =>
    sh.getName().startsWith(`Backup_${sourceName}_`)
  );


  // Parse timestamp and sort newest -> oldest
  backups.sort((a, b) => {
    const aTime = parseBackupTimestamp(a.getName(), sourceName);
    const bTime = parseBackupTimestamp(b.getName(), sourceName);
    return bTime - aTime; // newest first
  });


  // Delete old backups beyond MAX_BACKUPS
  if (backups.length > MAX_BACKUPS) {
    const toDelete = backups.slice(MAX_BACKUPS);
    toDelete.forEach(sh => {
      try {
        // avoid deleting active sheet in the targetSS
        targetSS.setActiveSheet(targetSS.getSheets()[0]);
        targetSS.deleteSheet(sh);
        logEvent("INFO", `Deleted old backup in '${targetSS.getName()}': ${sh.getName()}`);
      } catch (e) {
        logEvent("ERROR", `Failed to delete ${sh.getName()} in ${targetSS.getName()}: ${e}`);
      }
    });
  }
}


/**************************************************
 * Parse timestamp from backup sheet name
 * Expected sheet name: Backup_<source>_yyyy-MM-dd_HH:mm (exact format)
 **************************************************/
function parseBackupTimestamp(name, sourceName) {
  try {
    const prefix = `Backup_${sourceName}_`;
    if (!name.startsWith(prefix)) return 0;
    const parts = name.substring(prefix.length).split("_"); // ["yyyy-MM-dd","HH:mm"] or with suffix
    if (parts.length < 2) return 0;
    const datePart = parts[0]; // yyyy-MM-dd
    const timePart = parts[1].split("_")[0]; // HH:mm (ignore extra suffix if exists)


    const d = datePart.split("-");
    const t = timePart.split(":");
    if (d.length !== 3 || t.length < 1) return 0;


    const year = Number(d[0]);
    const month = Number(d[1]) - 1;
    const day = Number(d[2]);
    const hour = Number(t[0]) || 0;
    const minute = Number(t[1]) || 0;


    const dt = new Date(year, month, day, hour, minute);
    return dt.getTime();
  } catch (e) {
    return 0;
  }
}


/**************************************************
 * RESTORE LATEST BACKUP (Option C: repo if set, else current)
 **************************************************/
function restoreLatestBackup() {
  const settings = getSettings();
  const sourceName = settings.SOURCE_SHEET_NAME;
  const activeSS = SpreadsheetApp.getActive();


  // Determine target repository to read backups from
  let repoSS;
  if (settings.BACKUP_SPREADSHEET_ID && String(settings.BACKUP_SPREADSHEET_ID).trim()) {
    try {
      repoSS = SpreadsheetApp.openById(String(settings.BACKUP_SPREADSHEET_ID).trim());
    } catch (e) {
      logEvent("ERROR", `Invalid BACKUP_SPREADSHEET_ID: ${settings.BACKUP_SPREADSHEET_ID}`);
      repoSS = activeSS;
    }
  } else {
    repoSS = activeSS;
  }


  const sheets = repoSS.getSheets();
  const backups = sheets.filter(sh =>
    sh.getName().startsWith(`Backup_${sourceName}_`)
  );


  if (backups.length === 0) {
    SpreadsheetApp.getUi().alert("No backups available to restore.");
    return;
  }


  // Sort newest first by parsed timestamp
  backups.sort((a, b) => parseBackupTimestamp(b.getName(), sourceName) - parseBackupTimestamp(a.getName(), sourceName));
  const latestBackupSheet = backups[0];


  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    "Restore Backup",
    `This will replace sheet '${sourceName}' in the active spreadsheet with '${latestBackupSheet.getName()}' from '${repoSS.getName()}'. Continue?`,
    ui.ButtonSet.YES_NO
  );


  if (confirm !== ui.Button.YES) return;


  try {
    // Copy latest backup into active spreadsheet (restore)
    // If a sheet with sourceName exists in activeSS, delete it first
    const existing = activeSS.getSheetByName(sourceName);
    if (existing) activeSS.deleteSheet(existing);


    // copyTo returns a new sheet in activeSS
    const copied = latestBackupSheet.copyTo(activeSS);


    // ensure unique name if conflict, then rename to sourceName
    let restoredName = sourceName;
    if (activeSS.getSheetByName(restoredName)) {
      // should not happen since we deleted existing, but just in case
      restoredName = `${sourceName}_restored_${Utilities.formatDate(new Date(), TIMEZONE, "HHmmss")}`;
    }
    copied.setName(restoredName);


    // Move restored sheet to front
    activeSS.setActiveSheet(copied);
    activeSS.moveActiveSheet(1);


    ui.alert(`Restored '${restoredName}' from '${latestBackupSheet.getName()}'`);
    logEvent("RESTORE", `Restored ${restoredName} from ${repoSS.getName()} / ${latestBackupSheet.getName()}`);
  } catch (e) {
    logEvent("ERROR", "Restore failed: " + e);
    SpreadsheetApp.getUi().alert("Restore failed: " + e);
  }
}


/**************************************************
 * Logging
 **************************************************/
function logEvent(type, message) {
  const ss = SpreadsheetApp.getActive();
  let log = ss.getSheetByName("Backup_Log");


  if (!log) {
    log = ss.insertSheet("Backup_Log");
    log.appendRow(["Timestamp", "Type", "Message"]);
  }


  log.appendRow([
    Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss"),
    type,
    message
  ]);
}


/**************************************************
 * Open Settings Sheet
 **************************************************/
function openSettingsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Backup_Settings");


  if (!sh) {
    getSettings(); // will create it
    sh = ss.getSheetByName("Backup_Settings");
  }


  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert("Backup Settings sheet is ready.");
}



