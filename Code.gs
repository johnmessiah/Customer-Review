/****************************************************************
 * APP REVIEW DASHBOARD - MULTI-SHEET BACKEND (V2)
 * Developed by: John Messiah
 * Changes: Preference tab, People API, Enhanced User Base,
 *          Login page, Profile picture, Day-wise sessions
 ****************************************************************/

const ADMIN_EMAIL = "john.messiah@theporter.in";
const ALLOWED_DOMAIN = "@theporter.in";
const MASTER_SHEET_ID = "1jUdZLeWFxCaoUShlVrGRUSPEHeR3VIAS-SCzL8rhSZE";

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('App Review')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * GET CURRENT USER INFO (for login page)
 * Uses People API to fetch name and profile picture
 */
function getCurrentUserInfo() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return { success: false, error: 'No email found. Please sign in with your Google account.' };
    
    // Check domain
    if (!userEmail.endsWith(ALLOWED_DOMAIN) && userEmail !== ADMIN_EMAIL) {
      return { success: false, error: 'Access denied. Only @theporter.in accounts are allowed.' };
    }
    
    // Get user profile via People API
    let userName = userEmail.split('@')[0].split('.').map(s => s.charAt(0).toUpperCase() + s.slice(1)).join(' ');
    let photoUrl = '';
    
    try {
      const people = People.People.get('people/me', { personFields: 'names,photos' });
      if (people.names && people.names.length > 0) {
        userName = people.names[0].displayName || userName;
      }
      if (people.photos && people.photos.length > 0) {
        photoUrl = people.photos[0].url || '';
      }
    } catch (e) {
      // People API might fail, use fallback
      Logger.log('People API error: ' + e.message);
    }
    
    return {
      success: true,
      email: userEmail,
      name: userName,
      photoUrl: photoUrl
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * 1. DATA FETCHER & INITIALIZER
 */
function fetchDashboardData() {
  const result = {
    success: false, error: null, data: [], configList: [], accessList: [],
    logList: [], presetList: [], userList: [], user: {}, stats: {},
    preferences: {}
  };
  
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const userEmail = Session.getActiveUser().getEmail();

    // Get user profile via People API
    let userName = userEmail.split('@')[0].split('.').map(s => s.charAt(0).toUpperCase() + s.slice(1)).join(' ');
    let photoUrl = '';
    
    try {
      const people = People.People.get('people/me', { personFields: 'names,photos' });
      if (people.names && people.names.length > 0) {
        userName = people.names[0].displayName || userName;
      }
      if (people.photos && people.photos.length > 0) {
        photoUrl = people.photos[0].url || '';
      }
    } catch (e) {
      Logger.log('People API error: ' + e.message);
    }

    // 1. Auth 
    const accessSheet = getOrCreateSheet(ss, 'Access');
    const accessData = accessSheet.getDataRange().getDisplayValues();
    
    if (accessData.length <= 1) { 
       if (accessData.length === 0) accessSheet.appendRow(['Unique ID', 'User Mail ID', 'Role', 'Added By', 'Timestamp']);
       accessSheet.appendRow([Utilities.getUuid(), ADMIN_EMAIL, 'Admin', 'System', new Date()]); 
    }

    let role = 'Viewer';
    let hasAccess = false;
    
    if (userEmail === ADMIN_EMAIL) { role = 'Admin'; hasAccess = true; }
    else {
       const userRow = accessData.find(r => r[1] === userEmail);
       if (userRow) { role = userRow[2]; hasAccess = true; }
       else if (userEmail.endsWith(ALLOWED_DOMAIN)) { role = 'Viewer'; hasAccess = true; }
    }
    if (!hasAccess) throw new Error("Access Denied.");
    result.user = {
      email: userEmail,
      role: role,
      name: userName,
      photoUrl: photoUrl,
      isAdminOrUpdater: (role === 'Admin' || role === 'Updater')
    };

    // 2. User Base Tracking (Enhanced - Day-wise sessions)
    const userSheet = getOrCreateSheet(ss, 'User');
    let userData = userSheet.getDataRange().getValues();
    const userHeaders = ['Unique ID', 'User Gmail ID', 'Role', 'First Access', 'Last Access', 'Session History', 'Total Sessions', 'Number of Preset', 'Walkthrough', 'User Name'];
    
    if (userData.length === 0 || (userData[0] && userData[0][0] !== 'Unique ID')) {
      // Clear and re-create with proper headers
      userSheet.clearContents();
      userSheet.appendRow(userHeaders);
      userData = userSheet.getDataRange().getValues();
    }
    
    const now = new Date();
    const todayStr = Utilities.formatDate(now, 'Asia/Kolkata', 'dd/MM/yyyy');
    const timeStr = Utilities.formatDate(now, 'Asia/Kolkata', 'hh:mm:ss a');
    
    let userRowIdx = userData.findIndex(r => r[1] === userEmail);
    if (userRowIdx === -1) {
      // New User - Session History format: "dd/MM/yyyy:count|dd/MM/yyyy:count"
      const sessionHistory = todayStr + ':1';
      userSheet.appendRow([
        Utilities.getUuid(), userEmail, role, now, now,
        sessionHistory, 1, 0, '', userName
      ]);
      result.user.walkthrough = '';
    } else {
      // Returning User - Update session tracking
      let sessionHistory = (userData[userRowIdx][5] || '').toString();
      let totalSessions = parseInt(userData[userRowIdx][6]) || 0;
      
      // Parse session history
      let sessions = {};
      if (sessionHistory) {
        sessionHistory.split('|').forEach(entry => {
          const parts = entry.split(':');
          if (parts.length === 2) sessions[parts[0]] = parseInt(parts[1]) || 0;
        });
      }
      
      // Increment today's session count
      sessions[todayStr] = (sessions[todayStr] || 0) + 1;
      totalSessions++;
      
      // Rebuild session history string
      const newHistory = Object.entries(sessions).map(([d, c]) => d + ':' + c).join('|');
      
      userSheet.getRange(userRowIdx + 1, 5).setValue(now); // Last Access
      userSheet.getRange(userRowIdx + 1, 6).setValue(newHistory); // Session History
      userSheet.getRange(userRowIdx + 1, 7).setValue(totalSessions); // Total Sessions
      userSheet.getRange(userRowIdx + 1, 3).setValue(role); // Update role
      userSheet.getRange(userRowIdx + 1, 10).setValue(userName); // Update name
      
      result.user.walkthrough = userData[userRowIdx][8] || '';
    }

    // Fetch user list for Admin
    if (result.user.role === 'Admin') {
      const refreshedUsers = userSheet.getDataRange().getDisplayValues();
      result.userList = refreshedUsers.slice(1).map(r => ({
        id: r[0], email: r[1], role: r[2],
        firstAccess: r[3], lastAccess: r[4],
        sessionHistory: r[5], totalSessions: r[6],
        presets: r[7], walk: r[8], name: r[9]
      }));
    }

    // 3. Preferences (New Feature)
    const prefSheet = getOrCreateSheet(ss, 'Preference');
    let prefData = prefSheet.getDataRange().getValues();
    const prefHeaders = ['User Email', 'Filter Position', 'Theme', 'Trend Metric', 'Result By', 'Percentage Decimal', 'Perspective', 'Filter View', 'Visible Filters', 'Font Size', 'Font Style'];
    
    if (prefData.length === 0) {
      prefSheet.appendRow(prefHeaders);
      prefData = prefSheet.getDataRange().getValues();
    }
    
    let prefRowIdx = prefData.findIndex(r => r[0] === userEmail);
    if (prefRowIdx === -1) {
      const defaults = [userEmail, 'Left', 'White', 'Status', 'Count + Percentage', '0.0%', 'Detail Analysis', 'Basic', '', '14px', 'DM Sans'];
      prefSheet.appendRow(defaults);
      result.preferences = {
        filterPosition: 'Left', theme: 'White', trendMetric: 'Status',
        resultBy: 'Count + Percentage', percentageDecimal: '0.0%',
        perspective: 'Detail Analysis', filterView: 'Basic',
        visibleFilters: '', fontSize: '14px', fontStyle: 'DM Sans'
      };
    } else {
      const p = prefData[prefRowIdx];
      result.preferences = {
        filterPosition: p[1] || 'Left', theme: p[2] || 'White',
        trendMetric: p[3] || 'Status', resultBy: p[4] || 'Count + Percentage',
        percentageDecimal: p[5] || '0.0%', perspective: p[6] || 'Detail Analysis',
        filterView: p[7] || 'Basic', visibleFilters: p[8] || '',
        fontSize: p[9] || '14px', fontStyle: p[10] || 'DM Sans'
      };
    }

    // 4. User Presets Fetcher
    const presetSheet = getOrCreateSheet(ss, 'User Preset');
    let presetData = presetSheet.getDataRange().getDisplayValues();
    if (presetData.length === 0) {
      presetSheet.appendRow(['Unique ID', 'Preset Name', 'Creator', 'Visibility', 'Notes', 'Theme', 'Filters JSON', 'Timestamp']);
      presetData = presetSheet.getDataRange().getDisplayValues();
    }
    
    let userPresetCount = 0;
    for (let i = 1; i < presetData.length; i++) {
      const p = presetData[i];
      if (p[2] === userEmail) userPresetCount++;
      if (p[3] === 'Public' || p[2] === userEmail) {
        result.presetList.push({
          id: p[0], name: p[1], creator: p[2], visibility: p[3],
          notes: p[4], theme: p[5], filters: JSON.parse(p[6] || '{}'), timestamp: p[7]
        });
      }
    }
    
    // Update active user's preset count
    userRowIdx = userSheet.getDataRange().getValues().findIndex(r => r[1] === userEmail);
    if (userRowIdx > -1) userSheet.getRange(userRowIdx + 1, 8).setValue(userPresetCount);

    // 5. Configuration List & Data Fetching
    const configSheet = getOrCreateSheet(ss, 'Data Validation');
    const configRaw = configSheet.getDataRange().getDisplayValues();
    const configList = [];
    
    if (configRaw.length > 1) {
      for (let i = 1; i < configRaw.length; i++) {
        const row = configRaw[i];
        if (row[1] && row[2]) {
          configList.push({
            id: row[0], sheetId: row[1], tabName: row[2], timestamp: row[3],
            cols: {
              star: row[4], status: row[5], type: row[6], text: row[7],
              l1: row[8], l2: row[9], l3: row[10],
              veh: row[11], comp: row[12], city: row[13], 
              month: row[14], country: row[15], source: row[16]
            }
          });
        }
      }
    }
    result.configList = configList;

    let allData = [];
    let connected = 0;
    let failed = 0;
    let newLogs = []; 

    const getIdx = (char) => {
      if (!char) return -1;
      const s = char.toString().trim().toUpperCase();
      if (!s) return -1;
      let n = 0;
      for (let p = 0; p < s.length; p++) n = s.charCodeAt(p) - 64 + n * 26;
      return n - 1;
    };

    configList.forEach(conf => {
      try {
        const extSS = SpreadsheetApp.openById(conf.sheetId);
        const extSheet = extSS.getSheetByName(conf.tabName);
        if (!extSheet) throw new Error("Tab not found in spreadsheet.");

        const rawVals = extSheet.getDataRange().getDisplayValues();
        if (rawVals.length > 1) {
          const idx = {
            rating: getIdx(conf.cols.star), status: getIdx(conf.cols.status),
            type: getIdx(conf.cols.type), text: getIdx(conf.cols.text),
            l1: getIdx(conf.cols.l1), l2: getIdx(conf.cols.l2), l3: getIdx(conf.cols.l3),
            veh: getIdx(conf.cols.veh), comp: getIdx(conf.cols.comp), city: getIdx(conf.cols.city),
            month: getIdx(conf.cols.month), country: getIdx(conf.cols.country), source: getIdx(conf.cols.source)
          };

          const mapped = rawVals.slice(1).map(r => ({
            uid: Utilities.getUuid(),
            rating: parseFloat(r[idx.rating]) || 0,
            status: r[idx.status] || '', type: r[idx.type] || '', text: r[idx.text] || '', 
            l1: r[idx.l1] || '', l2: r[idx.l2] || '', l3: r[idx.l3] || '',
            vehicle: r[idx.veh] || '', comp: r[idx.comp] || '', city: r[idx.city] || '',
            month: r[idx.month] || '', country: r[idx.country] || '', source: r[idx.source] || ''
          }));
          allData = allData.concat(mapped);
          connected++;
        }
      } catch (e) {
        failed++;
        newLogs.push([Utilities.getUuid(), userEmail, new Date(), 'Fetch Error [' + conf.tabName + ']: ' + e.message]);
      }
    });

    result.data = allData;
    result.stats = { total: allData.length, connected: connected, failed: failed, sheets: configList.length };

    // 6. Access & Log Lists
    if (result.user.isAdminOrUpdater) {
      const refreshedAccess = accessSheet.getDataRange().getDisplayValues();
      result.accessList = refreshedAccess.slice(1).map(x => ({ email: x[1], role: x[2], addedBy: x[3], timestamp: x[4] }));
    }

    const logSheet = getOrCreateSheet(ss, 'Log');
    if (logSheet.getDataRange().getValues().length === 0) logSheet.appendRow(['Unique ID', 'User', 'Error Timestamp', 'Error Message']);
    if (newLogs.length > 0) logSheet.getRange(logSheet.getLastRow() + 1, 1, newLogs.length, 4).setValues(newLogs);
    
    const logDisplay = logSheet.getDataRange().getDisplayValues();
    result.logList = logDisplay.slice(1).map(x => ({ id: x[0], user: x[1], timestamp: x[2], message: x[3] }));
    
    result.success = true;
  } catch (e) {
    result.success = false;
    result.error = e.message;
  }
  return result;
}

/**
 * 2. SAVE USER PREFERENCES
 */
function savePreferences(prefs) {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = getOrCreateSheet(ss, 'Preference');
    const userEmail = Session.getActiveUser().getEmail();
    const data = sheet.getDataRange().getValues();
    
    const rowIdx = data.findIndex(r => r[0] === userEmail);
    const row = [
      userEmail,
      prefs.filterPosition || 'Left',
      prefs.theme || 'White',
      prefs.trendMetric || 'Status',
      prefs.resultBy || 'Count + Percentage',
      prefs.percentageDecimal || '0.0%',
      prefs.perspective || 'Detail Analysis',
      prefs.filterView || 'Basic',
      prefs.visibleFilters || '',
      prefs.fontSize || '14px',
      prefs.fontStyle || 'DM Sans'
    ];
    
    if (rowIdx > 0) {
      sheet.getRange(rowIdx + 1, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * 3. USER PRESET ACTIONS
 */
function managePreset(action, payload) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'User Preset');
  const userEmail = Session.getActiveUser().getEmail();
  const data = sheet.getDataRange().getValues();

  if (action === 'save') {
    if (payload.id) {
      const idx = data.findIndex(r => r[0] === payload.id);
      if (idx > 0) sheet.getRange(idx + 1, 2, 1, 7).setValues([[payload.name, userEmail, payload.visibility, payload.notes, payload.theme, JSON.stringify(payload.filters), new Date()]]);
    } else {
      sheet.appendRow([Utilities.getUuid(), payload.name, userEmail, payload.visibility, payload.notes, payload.theme, JSON.stringify(payload.filters), new Date()]);
    }
  } else if (action === 'delete') {
    const idx = data.findIndex(r => r[0] === payload.id);
    if (idx > 0) sheet.deleteRow(idx + 1);
  }
  return { success: true };
}

function completeWalkthrough() {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'User');
  const userEmail = Session.getActiveUser().getEmail();
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[1] === userEmail);
  if (idx > -1) sheet.getRange(idx + 1, 9).setValue('Completed');
  return { success: true };
}

/**
 * 4. ADMIN ACTIONS
 */
function saveConfigs(configArray) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'Data Validation');
  const existingRaw = sheet.getDataRange().getValues();
  const existingMap = {};
  if (existingRaw.length > 1) {
    for (let i = 1; i < existingRaw.length; i++) {
      const dataSignature = [existingRaw[i][1], existingRaw[i][2], ...existingRaw[i].slice(4)].join('|');
      existingMap[existingRaw[i][0]] = { timestamp: existingRaw[i][3], signature: dataSignature };
    }
  }
  
  sheet.clearContents();
  sheet.appendRow(['Unique ID', 'Sheet ID', 'Tab Name', 'Timestamp', 'Star Rating', 'Status', 'Order Type', 'Review Text', 'L1', 'L2', 'L3', 'Vehicle Type', 'Competitor', 'City', 'Month', 'Country', 'Source']);
  
  if (configArray.length > 0) {
    const rows = configArray.map(c => {
      const newRowData = [c.sheetId, c.tabName, c.cols.star, c.cols.status, c.cols.type, c.cols.text, c.cols.l1, c.cols.l2, c.cols.l3, c.cols.veh, c.cols.comp, c.cols.city, c.cols.month, c.cols.country, c.cols.source];
      const newSignature = newRowData.join('|');
      let tsToSave = new Date(); 
      if (c.id && existingMap[c.id] && existingMap[c.id].signature === newSignature) tsToSave = existingMap[c.id].timestamp || tsToSave;
      return [c.id || Utilities.getUuid(), c.sheetId, c.tabName, tsToSave, ...newRowData.slice(2)];
    });
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  return { success: true };
}

function manageAccess(action, payload) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'Access');
  const user = Session.getActiveUser().getEmail();

  if (action === 'add') {
    if (!payload.email.includes(ALLOWED_DOMAIN)) return { success: false, error: 'Domain must be ' + ALLOWED_DOMAIN };
    sheet.appendRow([Utilities.getUuid(), payload.email, payload.role, user, new Date()]);
  } else {
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i > 0; i--) if (data[i][1] === payload.email) sheet.deleteRow(i + 1);
  }
  
  const newData = sheet.getDataRange().getDisplayValues();
  return { success: true, list: newData.slice(1).map(x => ({ email: x[1], role: x[2], addedBy: x[3], timestamp: x[4] })) };
}

function getOrCreateSheet(ss, name) {
  let s = ss.getSheetByName(name);
  return s ? s : ss.insertSheet(name);
}
