/**
 * TEBO STUDIO'S — Google Apps Script Backend
 * ============================================
 * This file goes into your Google Apps Script project.
 * It reads/writes data from Google Sheets.
 * 
 * SHEETS REQUIRED (create these tabs in your spreadsheet):
 * 1. EventOrders  — event order data
 * 2. StudioOrders — studio work data
 * 3. Staff        — staff names & phones
 * 4. Settings     — app settings (review link, etc.)
 */

// ═══ SERVE THE WEB APP ═══
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Tebo Studio's — Order Management")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width,initial-scale=1.0,maximum-scale=1.0');
}

// ═══ HELPERS ═══
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Add headers based on sheet type
    if (name === 'EventOrders') {
      sheet.appendRow([
        'id','name','phone','email','odate','otime','contacts',
        'subEvents','reqs','pkgCat','pkgKey','particulars',
        'quoted','discount','total','quotationConfirmed',
        'a1_amt','a1_date','a1_meth','a2_amt','a2_date','a2_meth',
        'balance','baldate','received','paynotes','payStatus',
        'teamMembers','vendors',
        'arrivedAtVenue','photosCopied','nasCopied',
        'smugShared','smugLink','pixShared','pixLink',
        'dels','deldate',
        'proofLink','proofOk',
        'deliveryNotes','deliveryDate','deliveryTime','reviewLink',
        'status','notes','month','workDoneBy','crew',
        'savedAt','lastrem','remCount','imported'
      ]);
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    } else if (name === 'StudioOrders') {
      sheet.appendRow([
        'id','name','phone','odate','deldate','reqs','notes','total','status','savedAt'
      ]);
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold');
    } else if (name === 'Staff') {
      sheet.appendRow(['name', 'phone']);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    } else if (name === 'Settings') {
      sheet.appendRow(['key', 'value']);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      sheet.appendRow(['reviewLink', '']);
      sheet.appendRow(['nextEventSeq', '113']);
      sheet.appendRow(['nextStudioSeq', '1']);
    }
  }
  return sheet;
}

// ═══ EVENT ORDERS ═══

function getAllEventOrders() {
  const sheet = getSheet('EventOrders');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const orders = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      // Parse JSON fields
      if (['subEvents','reqs','teamMembers','vendors','dels','contacts'].includes(headers[j])) {
        try { val = JSON.parse(val); } catch { val = val || []; }
      }
      // Parse booleans
      if (['quotationConfirmed','arrivedAtVenue','photosCopied','nasCopied','smugShared','pixShared','proofOk','imported'].includes(headers[j])) {
        val = val === true || val === 'true' || val === 'TRUE';
      }
      // Parse a1/a2 compound fields
      if (headers[j] === 'a1_amt') {
        row.a1 = { amt: val || '', date: data[i][j+1] || '', meth: data[i][j+2] || '' };
      } else if (headers[j] === 'a2_amt') {
        row.a2 = { amt: val || '', date: data[i][j+1] || '', meth: data[i][j+2] || '' };
      } else if (!['a1_date','a1_meth','a2_date','a2_meth'].includes(headers[j])) {
        row[headers[j]] = val;
      }
      // Parse numbers
      if (['quoted','discount','total','received','balance','remCount'].includes(headers[j])) {
        row[headers[j]] = val || 0;
      }
    }
    row.type = 'event';
    row._row = i + 1; // 1-indexed sheet row for updates
    orders.push(row);
  }
  
  return orders;
}

function saveEventOrder(orderJson) {
  const order = JSON.parse(orderJson);
  const sheet = getSheet('EventOrders');
  
  // Check duplicate ID
  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (existing[i][0] === order.id) {
      return { error: 'ID already exists: ' + order.id };
    }
  }
  
  const row = [
    order.id, order.name, order.phone, order.email || '',
    order.odate || '', order.otime || '',
    JSON.stringify(order.contacts || []),
    JSON.stringify(order.subEvents || []),
    JSON.stringify(order.reqs || []),
    order.pkgCat || '', order.pkgKey || '', order.particulars || '',
    order.quoted || 0, order.discount || 0, order.total || 0,
    order.quotationConfirmed || false,
    order.a1?.amt || '', order.a1?.date || '', order.a1?.meth || '',
    order.a2?.amt || '', order.a2?.date || '', order.a2?.meth || '',
    order.balance || 0, order.baldate || '',
    order.received || 0, order.paynotes || '', order.payStatus || '',
    JSON.stringify(order.teamMembers || []),
    JSON.stringify(order.vendors || []),
    order.arrivedAtVenue || false,
    order.photosCopied || false, order.nasCopied || false,
    order.smugShared || false, order.smugLink || '',
    order.pixShared || false, order.pixLink || '',
    JSON.stringify(order.dels || []), order.deldate || '',
    order.proofLink || '', order.proofOk || false,
    order.deliveryNotes || '', order.deliveryDate || '',
    order.deliveryTime || '', order.reviewLink || '',
    order.status || 'New', order.notes || '',
    order.month || '', order.workDoneBy || '', order.crew || '',
    new Date().toISOString(), '', 0, false
  ];
  
  sheet.appendRow(row);
  
  // Update sequence counter
  updateSetting('nextEventSeq', getNextSeq('EventOrders'));
  
  return { success: true, id: order.id };
}

function updateEventOrder(orderJson) {
  const order = JSON.parse(orderJson);
  const sheet = getSheet('EventOrders');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === order.id) {
      // Update the specific fields that change frequently
      const headers = data[0];
      for (let j = 0; j < headers.length; j++) {
        const h = headers[j];
        if (h === 'status' && order.status !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(order.status);
        }
        if (h === 'lastrem' && order.lastrem !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(order.lastrem);
        }
        if (h === 'remCount' && order.remCount !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(order.remCount);
        }
        if (h === 'smugShared' && order.smugShared !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(order.smugShared);
        }
        if (h === 'pixShared' && order.pixShared !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(order.pixShared);
        }
        if (h === 'photosCopied' && order.photosCopied !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(order.photosCopied);
        }
        if (h === 'nasCopied' && order.nasCopied !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(order.nasCopied);
        }
      }
      return { success: true };
    }
  }
  return { error: 'Order not found: ' + order.id };
}

function deleteEventOrder(id) {
  const sheet = getSheet('EventOrders');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

// ═══ STUDIO ORDERS ═══

function getAllStudioOrders() {
  const sheet = getSheet('StudioOrders');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  return data.slice(1).map((row, idx) => {
    const obj = { type: 'studio', _row: idx + 2 };
    headers.forEach((h, j) => {
      let val = row[j];
      if (h === 'reqs') { try { val = JSON.parse(val); } catch { val = val ? [val] : []; } }
      obj[h] = val;
    });
    return obj;
  });
}

function saveStudioOrder(orderJson) {
  const order = JSON.parse(orderJson);
  const sheet = getSheet('StudioOrders');
  
  // Check duplicate
  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (existing[i][0] === order.id) return { error: 'ID exists' };
  }
  
  sheet.appendRow([
    order.id, order.name, order.phone, order.odate || '',
    order.deldate || '', JSON.stringify(order.reqs || []),
    order.notes || '', order.total || 0,
    order.status || 'New', new Date().toISOString()
  ]);
  
  return { success: true, id: order.id };
}

function updateStudioOrder(orderJson) {
  const order = JSON.parse(orderJson);
  const sheet = getSheet('StudioOrders');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === order.id) {
      if (order.status !== undefined) sheet.getRange(i + 1, 9).setValue(order.status);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

function deleteStudioOrder(id) {
  const sheet = getSheet('StudioOrders');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Not found' };
}

// ═══ STAFF ═══

function getAllStaff() {
  const sheet = getSheet('Staff');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(r => ({ name: r[0] || '', phone: r[1] || '' }));
}

function addStaffMember(name, phone) {
  const sheet = getSheet('Staff');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) return { error: 'Already exists' };
  }
  sheet.appendRow([name, phone || '']);
  return { success: true };
}

function removeStaffMember(index) {
  const sheet = getSheet('Staff');
  if (index >= 0 && index < sheet.getLastRow() - 1) {
    sheet.deleteRow(index + 2); // +2 for header + 0-index
    return { success: true };
  }
  return { error: 'Invalid index' };
}

// ═══ SETTINGS ═══

function getSetting(key) {
  const sheet = getSheet('Settings');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return '';
}

function updateSetting(key, value) {
  const sheet = getSheet('Settings');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getNextSeq(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]);
    const m = id.match(/[ES](\d+)$/);
    if (m) maxSeq = Math.max(maxSeq, parseInt(m[1]));
  }
  return maxSeq + 1;
}

// ═══ CHECK DUPLICATE ID ═══

function checkDuplicateId(id) {
  const evSheet = getSheet('EventOrders');
  const stSheet = getSheet('StudioOrders');
  
  const evData = evSheet.getDataRange().getValues();
  for (let i = 1; i < evData.length; i++) {
    if (evData[i][0] === id) return true;
  }
  
  const stData = stSheet.getDataRange().getValues();
  for (let i = 1; i < stData.length; i++) {
    if (stData[i][0] === id) return true;
  }
  
  return false;
}

// ═══ IMPORT EXISTING DATA ═══
// Call this once to import your existing Excel data into Sheets

function importExistingData(dataJson) {
  const orders = JSON.parse(dataJson);
  const sheet = getSheet('EventOrders');
  
  let imported = 0;
  orders.forEach(order => {
    try {
      saveEventOrder(JSON.stringify(order));
      imported++;
    } catch (e) {
      Logger.log('Skip ' + order.id + ': ' + e.message);
    }
  });
  
  return { imported: imported, total: orders.length };
}

// ═══ BULK LOAD (for initial page load) ═══

function loadAllData() {
  return {
    eventOrders: getAllEventOrders(),
    studioOrders: getAllStudioOrders(),
    staff: getAllStaff(),
    settings: {
      reviewLink: getSetting('reviewLink'),
      nextEventSeq: getSetting('nextEventSeq') || 113,
      nextStudioSeq: getSetting('nextStudioSeq') || 1
    }
  };
}

// ═══ DAILY AUTOMATION TRIGGER ═══
// Set this up once: Run setupDailyTrigger() from the editor
// It will run every day at 9 AM and flag orders needing action

function setupDailyTrigger() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyAutomation') ScriptApp.deleteTrigger(t);
  });
  // Create new daily trigger at 9 AM IST
  ScriptApp.newTrigger('dailyAutomation')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .inTimezone('Asia/Kolkata')
    .create();
  Logger.log('Daily trigger set for 9 AM IST');
}

function dailyAutomation() {
  const sheet = getSheet('EventOrders');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  const headers = data[0];
  const today = new Date();
  const dateCol = headers.indexOf('subEvents');
  const statusCol = headers.indexOf('status');
  const smugCol = headers.indexOf('smugShared');
  const pixCol = headers.indexOf('pixShared');
  const lastremCol = headers.indexOf('lastrem');
  const remCountCol = headers.indexOf('remCount');
  const nameCol = headers.indexOf('name');
  const idCol = 0;
  
  let summary = '📊 Daily Tebo Studio Report\n\n';
  let teamReminders = [];
  let photoReminders = [];
  let vendorAlerts = [];
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][statusCol];
    if (status === 'Complete') continue;
    
    const id = data[i][idCol];
    const name = data[i][nameCol];
    
    // Parse event date from subEvents JSON
    let evDate = null;
    try {
      const subEv = JSON.parse(data[i][dateCol]);
      if (subEv && subEv[0] && subEv[0].evdate) {
        evDate = new Date(subEv[0].evdate + 'T00:00:00');
      }
    } catch {}
    
    if (!evDate) continue;
    
    const daysTo = Math.ceil((evDate - today) / (1000 * 60 * 60 * 24));
    
    // Team reminders: 3 days before & 1 day before
    if (daysTo === 3) {
      teamReminders.push({ id, name, days: 3 });
    }
    if (daysTo === 1) {
      teamReminders.push({ id, name, days: 1 });
    }
    
    // Photo selection: 15+ days after event, weekly
    if (daysTo <= -15) {
      const smugShared = data[i][smugCol] === true || data[i][smugCol] === 'true';
      const pixShared = data[i][pixCol] === true || data[i][pixCol] === 'true';
      if (!smugShared || !pixShared) {
        const lastRem = data[i][lastremCol];
        const lastRemDate = lastRem ? new Date(lastRem) : null;
        const daysSinceRem = lastRemDate ? Math.floor((today - lastRemDate) / (1000 * 60 * 60 * 24)) : 999;
        if (daysSinceRem >= 7) {
          photoReminders.push({ id, name, daysSince: Math.abs(daysTo), row: i + 1 });
          // Auto-update lastrem date
          sheet.getRange(i + 1, lastremCol + 1).setValue(today.toISOString().split('T')[0]);
          sheet.getRange(i + 1, remCountCol + 1).setValue((data[i][remCountCol] || 0) + 1);
        }
      }
    }
    
    // Vendor payment check: 7+ days after event
    if (daysTo <= -7) {
      try {
        const team = JSON.parse(data[i][headers.indexOf('teamMembers')]);
        if (team) {
          team.forEach(m => {
            if (m.vendorAmt && Number(m.vendorAmt) > 0 && m.vendorStatus !== 'Paid') {
              vendorAlerts.push({ id, name, member: m.name, amt: m.vendorAmt });
            }
          });
        }
      } catch {}
    }
  }
  
  // Build summary
  if (teamReminders.length) {
    summary += '👥 TEAM REMINDERS:\n';
    teamReminders.forEach(r => {
      summary += `  • ${r.id} (${r.name}) — ${r.days} day(s) to event\n`;
    });
    summary += '\n';
  }
  
  if (photoReminders.length) {
    summary += '📸 PHOTO SELECTION REMINDERS (auto-sent):\n';
    photoReminders.forEach(r => {
      summary += `  • ${r.id} (${r.name}) — ${r.daysSince} days since event\n`;
    });
    summary += '\n';
  }
  
  if (vendorAlerts.length) {
    summary += '💰 VENDOR PAYMENTS OVERDUE:\n';
    vendorAlerts.forEach(r => {
      summary += `  • ${r.id} (${r.name}) — ${r.member}: ₹${r.amt}\n`;
    });
    summary += '\n';
  }
  
  if (!teamReminders.length && !photoReminders.length && !vendorAlerts.length) {
    summary += '✅ No pending actions today.\n';
  }
  
  // Log the summary
  Logger.log(summary);
  
  // Optional: Email the summary to admin
  // Uncomment and set your email:
  // MailApp.sendEmail('tebophotography@gmail.com', 'Tebo Studio Daily Report', summary);
  
  return summary;
}

// ═══════════════════════════════════════════
// DAILY EXCEL BACKUP — 10 AM and 10 PM IST
// ═══════════════════════════════════════════

function setupBackupTriggers() {
  // Remove existing backup triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['morningBackup','eveningBackup','dailyAutomation'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 10 AM IST daily backup
  ScriptApp.newTrigger('morningBackup')
    .timeBased().everyDays(1).atHour(10)
    .inTimezone('Asia/Kolkata').create();

  // 10 PM IST daily backup
  ScriptApp.newTrigger('eveningBackup')
    .timeBased().everyDays(1).atHour(22)
    .inTimezone('Asia/Kolkata').create();

  // Daily automation check at 9 AM
  ScriptApp.newTrigger('dailyAutomation')
    .timeBased().everyDays(1).atHour(9)
    .inTimezone('Asia/Kolkata').create();

  Logger.log('✅ All triggers set: 9 AM automation, 10 AM backup, 10 PM backup');
}

function morningBackup() {
  exportToExcel('Morning Backup');
}

function eveningBackup() {
  exportToExcel('Evening Backup');
}

function exportToExcel(label) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm');
    const dateStr = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
    
    // Create Excel export using Sheets API
    const ssId = ss.getId();
    const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&portrait=true`;
    
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    });
    
    // Save to Google Drive in a "Tebo Backups" folder
    let backupFolder;
    const folders = DriveApp.getFoldersByName('Tebo Studio Backups');
    if (folders.hasNext()) {
      backupFolder = folders.next();
    } else {
      backupFolder = DriveApp.createFolder('Tebo Studio Backups');
      Logger.log('Created backup folder: Tebo Studio Backups');
    }
    
    // Save file with timestamp
    const fileName = `Tebo_Studio_${dateStr}_${label.replace(' ','_')}.xlsx`;
    
    // Delete old backup with same name if exists (keep only latest per day per slot)
    const existing = backupFolder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);
    
    // Save new backup
    backupFolder.createFile(response.getBlob().setName(fileName));
    
    Logger.log(`✅ ${label} saved: ${fileName}`);
    
    // Keep only last 7 days of backups (auto-cleanup)
    cleanOldBackups(backupFolder, 7);
    
    return { success: true, file: fileName, time: timestamp };
    
  } catch(e) {
    Logger.log('❌ Backup failed: ' + e.message);
    return { error: e.message };
  }
}

function cleanOldBackups(folder, keepDays) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - keepDays);
  
  const files = folder.getFiles();
  let deleted = 0;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < cutoff) {
      file.setTrashed(true);
      deleted++;
    }
  }
  if (deleted > 0) Logger.log(`Cleaned ${deleted} old backup(s)`);
}

// Manual backup — call anytime from editor
function manualBackup() {
  const result = exportToExcel('Manual');
  if (result.success) {
    Logger.log('✅ Manual backup saved to Google Drive > Tebo Studio Backups > ' + result.file);
  } else {
    Logger.log('❌ Failed: ' + result.error);
  }
}

// ═══════════════════════════════════════════
// SYNC APP DATA TO SHEETS
// Called from the app when saving orders
// ═══════════════════════════════════════════

function syncAllOrders(ordersJson) {
  try {
    const orders = JSON.parse(ordersJson);
    const evOrders = orders.filter(o => o.type !== 'studio');
    const stOrders = orders.filter(o => o.type === 'studio');
    
    // Clear and rewrite EventOrders sheet
    const evSheet = getSheet('EventOrders');
    const evData = evSheet.getDataRange().getValues();
    const headers = evData[0];
    
    // Keep headers, clear data rows
    if (evSheet.getLastRow() > 1) {
      evSheet.deleteRows(2, evSheet.getLastRow() - 1);
    }
    
    // Write all event orders
    evOrders.forEach(o => {
      const se = (o.subEvents && o.subEvents[0]) || {};
      const team = (o.teamMembers || []).filter(m => m.name)
        .map(m => `${m.name}(${m.role||''})`).join(', ');
      evSheet.appendRow([
        o.id, o.name, o.phone||'', o.email||'',
        o.evdate||se.evdate||'', o.month||'',
        se.evtype||'', se.venue||'', o.pkgKey||'',
        (o.reqs||[]).join(', '),
        o.quoted||0, o.discount||0, o.total||0,
        (o.a1&&o.a1.amt)||0, (o.a2&&o.a2.amt)||0, o.balance||0,
        o.status||'New',
        o.photosCopied?'Yes':'No',
        o.nasCopied?'Yes':'No',
        o.smugShared?'Yes':'No', o.smugLink||'',
        o.pixShared?'Yes':'No', o.pixLink||'',
        team,
        o.deliveryNotes||'', o.deldate||'',
        o.proofOk?'Yes':'No',
        new Date().toISOString()
      ]);
    });
    
    // Write studio orders
    const stSheet = getSheet('StudioOrders');
    if (stSheet.getLastRow() > 1) {
      stSheet.deleteRows(2, stSheet.getLastRow() - 1);
    }
    stOrders.forEach(o => {
      stSheet.appendRow([
        o.id, o.name, o.phone||'', o.odate||'',
        o.deldate||'', (o.reqs||[]).join(', '),
        o.notes||'', o.total||0, o.status||'New'
      ]);
    });
    
    Logger.log(`Synced ${evOrders.length} events + ${stOrders.length} studio orders`);
    return { success: true, events: evOrders.length, studio: stOrders.length };
    
  } catch(e) {
    Logger.log('Sync error: ' + e.message);
    return { error: e.message };
  }
}
