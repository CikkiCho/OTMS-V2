// ============================================
// OT MONITORING SYSTEM - MAIN CODE
// ============================================

// ======== CONFIG ========
const CONFIG = {
  SPREADSHEET_ID: '1o1uyzh5osT6kqGI6cAQMEF5YQF2rPeeLIqHFEll6ZRs',
  SHEET_STAFF: 'Users',
  SHEET_OT: 'OT_Applications',
  SHEET_CLAIMS: 'OT_Claim',
  MAX_OT_HOURS: 104,
  AMBER_THRESHOLD: 90,
  SESSION_TIMEOUT_MINUTES: 30
};

// Legacy constants for backward compatibility
const SHEET_STAFF = CONFIG.SHEET_STAFF;
const SHEET_OT = CONFIG.SHEET_OT;
const SHEET_CLAIMS = CONFIG.SHEET_CLAIMS;
const MAX_OT_HOURS = CONFIG.MAX_OT_HOURS;
const AMBER_THRESHOLD = CONFIG.AMBER_THRESHOLD;
const SESSION_TIMEOUT_MINUTES = CONFIG.SESSION_TIMEOUT_MINUTES;

// ======== doGet ========
function doGet(e) {
  try {
    const webAppUrl = ScriptApp.getService().getUrl();
    const page = e.parameter ? e.parameter.page : 'login';
    const sessionId = e.parameter ? e.parameter.sessionId : null;
    
    Logger.log('====================================');
    Logger.log('doGet called at: ' + new Date().toISOString());
    Logger.log('Page: ' + page);
    Logger.log('SessionId: ' + sessionId);
    Logger.log('SessionId type: ' + typeof sessionId);
    Logger.log('====================================');
    
    // Route based on page parameter
    switch(page) {
      case 'staff-dashboard':
        Logger.log('Loading staff dashboard, sessionId: ' + sessionId);
        const staffSession = validateSession(sessionId);
        Logger.log('Staff session valid: ' + staffSession.valid);
        Logger.log('Staff session user: ' + JSON.stringify(staffSession.user));
        if (!staffSession.valid) {
          return redirectToLogin(webAppUrl, 'Session expired. Please login again.');
        }
        return loadStaffDashboard(webAppUrl, sessionId, staffSession.user);
        
      case 'management-dashboard':
        const mgmtSession = validateSession(sessionId);
        if (!mgmtSession.valid) {
          return redirectToLogin(webAppUrl, 'Session expired. Please login again.');
        }
        return loadManagementDashboard(webAppUrl, sessionId, mgmtSession.user);
        
      case 'ot-application':
        const otSession = validateSession(sessionId);
        if (!otSession.valid) {
          return redirectToLogin(webAppUrl, 'Session expired. Please login again.');
        }
        return loadOTApplicationPage(webAppUrl, sessionId, otSession.user);
        
      case 'ot-claim':
        const claimSession = validateSession(sessionId);
        if (!claimSession.valid) {
          return redirectToLogin(webAppUrl, 'Session expired. Please login again.');
        }
        return loadOTClaimPage(webAppUrl, sessionId, claimSession.user);
        
      case 'leave-application':
        const leaveSession = validateSession(sessionId);
        if (!leaveSession.valid) {
          return redirectToLogin(webAppUrl, 'Session expired. Please login again.');
        }
        return loadLeaveApplicationPage(webAppUrl, sessionId, leaveSession.user);
        
      case 'login':
      default:
        return loadLoginPage(webAppUrl);
    }
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>System Error</h1><p>' + error.message + '</p>');
  }
}

// ======== PAGE LOADERS ========
function loadLoginPage(webAppUrl) {
  const template = HtmlService.createTemplateFromFile('login');
  template.webAppUrl = webAppUrl;
  return template.evaluate()
    .setTitle('OTMS - Login')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function loadStaffDashboard(webAppUrl, sessionId, user) {
  Logger.log('=== loadStaffDashboard called ===');
  Logger.log('User object received: ' + JSON.stringify(user));
  
  if (!user) {
    Logger.log('ERROR: User object is null/undefined');
    return redirectToLogin(webAppUrl, 'Session invalid. Please login again.');
  }
  
  Logger.log('User email: ' + user.email);
  Logger.log('User name: ' + user.name);
  Logger.log('User team: ' + user.team);
  
  if (!user.email || user.email === 'undefined') {
    Logger.log('ERROR: User email is invalid');
    return redirectToLogin(webAppUrl, 'Session data corrupted. Please login again.');
  }
  
  const template = HtmlService.createTemplateFromFile('staff');
  template.webAppUrl = webAppUrl;
  template.sessionId = sessionId;
  template.userEmail = user.email || '';
  template.userName = user.name || 'Unknown';
  template.userTeam = user.team || 'N/A';
  template.cacheBreaker = new Date().getTime();
  
  Logger.log('Template variables set - Email: ' + template.userEmail + ', Name: ' + template.userName);
  
  return template.evaluate()
    .setTitle('OTMS - Staff Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function loadManagementDashboard(webAppUrl, sessionId, user) {
  Logger.log('=== loadManagementDashboard called ===');
  Logger.log('User object: ' + JSON.stringify(user));
  
  if (!user || !user.email) {
    Logger.log('ERROR: Invalid user data in loadManagementDashboard');
    return redirectToLogin(webAppUrl, 'Session invalid. Please login again.');
  }
  
  const template = HtmlService.createTemplateFromFile('management');
  template.webAppUrl = webAppUrl;
  template.sessionId = sessionId;
  template.userEmail = user.email || '';
  template.userName = user.name || 'Unknown';
  
  Logger.log('Management template set - Email: ' + template.userEmail);
  return template.evaluate()
    .setTitle('OTMS - Management Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function loadOTApplicationPage(webAppUrl, sessionId, user) {
  Logger.log('=== loadOTApplicationPage called ===');
  Logger.log('User object: ' + JSON.stringify(user));
  
  if (!user || !user.email) {
    Logger.log('ERROR: Invalid user data in loadOTApplicationPage');
    return redirectToLogin(webAppUrl, 'Session invalid. Please login again.');
  }
  
  const template = HtmlService.createTemplateFromFile('ot-application');
  template.webAppUrl = webAppUrl;
  template.sessionId = sessionId;
  template.userEmail = user.email || '';
  template.userName = user.name || 'Unknown';
  template.userTeam = user.team || 'N/A';
  
  Logger.log('OT App template set - Email: ' + template.userEmail);
  return template.evaluate()
    .setTitle('OTMS - OT Application')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function loadOTClaimPage(webAppUrl, sessionId, user) {
  const template = HtmlService.createTemplateFromFile('ot-claim');
  template.webAppUrl = webAppUrl;
  template.sessionId = sessionId;
  template.userEmail = user.email;
  template.userName = user.name;
  template.userTeam = user.team || 'N/A';
  return template.evaluate()
    .setTitle('OTMS - OT Claim')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function loadLeaveApplicationPage(webAppUrl, sessionId, user) {
  const template = HtmlService.createTemplateFromFile('leave-application');
  template.webAppUrl = webAppUrl;
  template.sessionId = sessionId;
  template.userEmail = user.email;
  template.userName = user.name;
  template.userTeam = user.team || 'N/A';
  return template.evaluate()
    .setTitle('OTMS - Leave Request')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function redirectToLogin(webAppUrl, message) {
  return HtmlService.createHtmlOutput(`
    <script>
      alert('${message}');
      window.top.location.href = '${webAppUrl}?page=login';
    </script>
  `);
}

// ======== SESSION MANAGEMENT ========
function generateSessionId() {
  return Utilities.getUuid();
}

function createSession(email, name, role, team) {
  const sessionId = generateSessionId();
  const sessionData = {
    email: email,
    name: name,
    role: role,
    team: team,
    created: new Date().toISOString(),
    expires: new Date(Date.now() + SESSION_TIMEOUT_MINUTES * 60 * 1000).toISOString()
  };
  
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('session_' + sessionId, JSON.stringify(sessionData));
  
  Logger.log('Session created: ' + sessionId + ' for ' + email);
  return sessionId;
}

function validateSession(sessionId) {
  if (!sessionId) {
    return { valid: false, user: null };
  }
  
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const sessionDataStr = scriptProperties.getProperty('session_' + sessionId);
    
    if (!sessionDataStr) {
      Logger.log('Session not found: ' + sessionId);
      return { valid: false, user: null };
    }
    
    const sessionData = JSON.parse(sessionDataStr);
    const now = new Date();
    const expires = new Date(sessionData.expires);
    
    if (now > expires) {
      Logger.log('Session expired: ' + sessionId);
      deleteSession(sessionId);
      return { valid: false, user: null };
    }
    
    // Session valid - extend expiration
    sessionData.expires = new Date(Date.now() + SESSION_TIMEOUT_MINUTES * 60 * 1000).toISOString();
    scriptProperties.setProperty('session_' + sessionId, JSON.stringify(sessionData));
    
    return { 
      valid: true, 
      user: {
        email: sessionData.email,
        name: sessionData.name,
        role: sessionData.role,
        team: sessionData.team
      }
    };
  } catch (error) {
    Logger.log('Error validating session: ' + error.toString());
    return { valid: false, user: null };
  }
}

function deleteSession(sessionId) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('session_' + sessionId);
    Logger.log('Session deleted: ' + sessionId);
    return { success: true };
  } catch (error) {
    Logger.log('Error deleting session: ' + error.toString());
    return { success: false, error: error.message };
  }
}

// ======== LOGIN FUNCTION ========
function login(email) {
  Logger.log('=== LOGIN ATTEMPT ===');
  Logger.log('Email received: "' + email + '"');
  Logger.log('Email type: ' + typeof email);
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  if (!ss) {
    Logger.log('ERROR: Cannot open spreadsheet by ID');
    return { error: 'Cannot access database' };
  }
  
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (!staffSheet) {
    Logger.log('ERROR: Users sheet not found');
    return { success: false, role: null, sessionId: null, message: 'Users sheet not found' };
  }
  
  const data = staffSheet.getDataRange().getValues();
  const emailToFind = (email || '').toString().trim().toLowerCase();
  
  Logger.log('Looking for email: "' + emailToFind + '"');
  Logger.log('Total rows to scan: ' + (data.length - 1));
  
  for (let i = 1; i < data.length; i++) {
    const sheetEmail = (data[i][1] || '').toString().trim().toLowerCase();
    const name = data[i][0]; // Column A = Name
    const role = data[i][2]; // Column C = Role
    const team = data[i][3]; // Column D = Team
    
    Logger.log('Row ' + (i + 1) + ': Sheet email="' + sheetEmail + '" | Name="' + name + '" | Role="' + role + '"');
    
    if (sheetEmail === emailToFind) {
      Logger.log('✅ MATCH FOUND at row ' + (i + 1));
      
      // Create session with original email format
      const sessionId = createSession(data[i][1], name, role, team);
      
      Logger.log('Login successful for: ' + name + ' (' + role + ')');
      Logger.log('Session ID: ' + sessionId);
      
      return { 
        success: true, 
        role: role,
        sessionId: sessionId
      };
    }
  }
  
  Logger.log('❌ Email not found in Users sheet');
  Logger.log('Available emails:');
  for (let i = 1; i < Math.min(data.length, 6); i++) {
    Logger.log('  Row ' + (i + 1) + ': "' + data[i][1] + '"');
  }
  
  return { success: false, role: null, sessionId: null, message: 'Email not found. Please check your email address.' };
}

// ======== LOGOUT FUNCTION ========
function logout(sessionId) {
  return deleteSession(sessionId);
}

// ======== STAFF DASHBOARD ========
function getStaffDashboard(email) {
  const callTime = new Date().toISOString();
  Logger.log('=== getStaffDashboard START at ' + callTime + ' ===');
  Logger.log('Email received: ' + email);
  Logger.log('Email type: ' + typeof email);
  
  // Validate email parameter
  if (!email || email === 'undefined' || email === '' || email === null) {
    Logger.log('ERROR: Invalid email parameter');
    const errorResponse = {
      totalOTHours: 0,
      remainingHours: 104,
      remainingColor: 'green',
      applicationCount: 0,
      pendingCount: 0,
      latestClaim: null,
      applications: [],
      error: 'Invalid email. Please login again.'
    };
    Logger.log('Returning error response: ' + JSON.stringify(errorResponse));
    return errorResponse;
  }
  
  try {
    // Use OT Calculation Module to get user stats
    const otStats = calculateUserOTStats(email);
    
    if (!otStats.success) {
      Logger.log('ERROR: Failed to calculate OT stats - ' + otStats.error);
      return {
        totalOTHours: 0,
        remainingHours: 104,
        remainingColor: 'green',
        latestClaim: null,
        applications: [],
        error: otStats.error
      };
    }
    
    Logger.log('OT Stats calculated: Total=' + otStats.totalOTHours + 'h, Remaining=' + otStats.remainingHours + 'h');
    
    // Now get the applications list
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return { error: 'Cannot access database' };
    }
    
    const usersSheet = ss.getSheetByName(SHEET_STAFF);
    const otSheet = ss.getSheetByName(SHEET_OT);
    
    if (!usersSheet || !otSheet) {
      Logger.log('ERROR: Required sheets not found');
      return {
        totalOTHours: otStats.totalOTHours,
        remainingHours: otStats.remainingHours,
        remainingColor: otStats.remainingColor,
        latestClaim: null,
        applications: []
      };
    }
    
    // Get userName from email
    const usersData = usersSheet.getDataRange().getValues();
    let userName = null;
    const emailToFind = (email || '').toString().trim().toLowerCase();
    
    for (let i = 1; i < usersData.length; i++) {
      const sheetEmail = (usersData[i][1] || '').toString().trim().toLowerCase();
      if (sheetEmail === emailToFind) {
        userName = usersData[i][0].toString().trim();
        Logger.log('Found user: ' + userName + ' for email: ' + email);
        break;
      }
    }
    
    if (!userName) {
      Logger.log('ERROR: User not found for email: ' + email);
      return {
        totalOTHours: otStats.totalOTHours || 0,
        remainingHours: otStats.remainingHours || 104,
        remainingColor: otStats.remainingColor || 'green',
        applicationCount: 0,
        pendingCount: 0,
        latestClaim: null,
        applications: [],
        error: 'User not found in system'
      };
    }
    
    // Build applications array
    const otData = otSheet.getDataRange().getValues();
    const applications = [];
    let pendingCount = 0;

    Logger.log('--- SCANNING OT_APPLICATIONS SHEET ---');

    // Start from row 1 (skip header at row 0)
    for (let i = 1; i < otData.length; i++) {
      const row = otData[i];

      // Skip empty rows
      if (!row[0]) continue;

      const rowName = row[0] ? row[0].toString().trim() : '';
      const rowStatus = row[10] ? row[10].toString().trim() : 'Pending';

      // Match by name
      if (rowName === userName) {
        const hours = parseFloat(row[5]) || 0;

        // Count pending applications
        if (rowStatus === 'Pending') {
          pendingCount++;
        }

        // Convert date to string format for serialization
        let dateStr = '';
        if (row[2]) {
          if (row[2] instanceof Date) {
            dateStr = Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'yyyy-MM-dd');
          } else {
            dateStr = row[2].toString();
          }
        }

        // Format time fields as HH:MM
        let startTimeStr = '';
        if (row[3]) {
          if (row[3] instanceof Date) {
            startTimeStr = Utilities.formatDate(row[3], Session.getScriptTimeZone(), 'HH:mm');
          } else {
            startTimeStr = row[3].toString();
          }
        }

        let endTimeStr = '';
        if (row[4]) {
          if (row[4] instanceof Date) {
            endTimeStr = Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'HH:mm');
          } else {
            endTimeStr = row[4].toString();
          }
        }

        // Build application object
        applications.push({
          rowNumber: i + 1,
          Name: row[0] || '',
          Team: row[1] || '',
          Date: dateStr,
          StartTime: startTimeStr,
          EndTime: endTimeStr,
          TotalHours: hours,
          OT_Type: row[6] || '',
          Is_Public_Holiday: row[7] || '',
          Proof_Attendance: row[8] || '',
          Reason: row[9] || '',
          Status: rowStatus,
          ApprovedBy: row[11] || '',
          Approval_Date: row[12] ? row[12].toString() : ''
        });
      }
    }
    
    Logger.log('--- SCAN COMPLETE ---');
    Logger.log('Total applications found: ' + applications.length);
    Logger.log('Pending applications: ' + pendingCount);
    
    // Return result using OT Calculation Module stats
    const result = {
      totalOTHours: otStats.totalOTHours,
      remainingHours: otStats.remainingHours,
      remainingColor: otStats.remainingColor,
      applicationCount: applications.length,
      pendingCount: pendingCount,
      latestClaim: null,
      applications: applications
    };
    
    Logger.log('=== getStaffDashboard END ===');
    Logger.log('Returning: totalOTHours=' + otStats.totalOTHours + ', remainingHours=' + otStats.remainingHours + ', apps=' + applications.length + ', pending=' + pendingCount);
    Logger.log('First 3 applications:');
    for (let i = 0; i < Math.min(applications.length, 3); i++) {
      Logger.log('  App ' + (i+1) + ': ' + applications[i].Name + ' (' + applications[i].Date + ') Status="' + applications[i].Status + '"');
    }
    Logger.log('Result object keys: ' + Object.keys(result).join(', '));

    return result;
    
  } catch (error) {
    Logger.log('=== ERROR in getStaffDashboard ===');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    
    // Always return valid structure, never null
    return {
      totalOTHours: 0,
      remainingHours: 104,
      remainingColor: 'green',
      applicationCount: 0,
      pendingCount: 0,
      latestClaim: null,
      applications: []
    };
  }
}

// ======== EDIT OT APPLICATION ========
function editOTApplication(formData) {
  try {
    Logger.log('=== editOTApplication START ===');
    Logger.log('Editing row: ' + formData.rowNumber);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('OT_Applications');
    
    if (!sheet) {
      return { success: false, message: 'OT_Applications sheet not found' };
    }
    
    const rowNum = parseInt(formData.rowNumber);
    
    // Update the row with new data
    sheet.getRange(rowNum, 3).setValue(formData.date);           // Date (Column C)
    sheet.getRange(rowNum, 4).setValue(formData.startTime);      // Start Time (Column D)
    sheet.getRange(rowNum, 5).setValue(formData.endTime);        // End Time (Column E)
    sheet.getRange(rowNum, 6).setValue(formData.totalHours);     // Total Hours (Column F)
    sheet.getRange(rowNum, 7).setValue(formData.otType);         // OT Type (Column G)
    sheet.getRange(rowNum, 8).setValue(formData.isPublicHoliday); // Public Holiday (Column H)
    sheet.getRange(rowNum, 10).setValue(formData.reason);        // Reason (Column J)

    // Ensure changes are written immediately
    SpreadsheetApp.flush();

    // Read back the updated row for verification
    const saved = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('Post-edit verification for row ' + rowNum + ': ' + JSON.stringify(saved));

    Logger.log('Application updated successfully');
    return { success: true, message: 'Application updated successfully', rowNumber: rowNum, saved: saved };
    
  } catch (error) {
    Logger.log('Error in editOTApplication: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ======== DELETE OT APPLICATION ========
function deleteOTApplication(rowNumber) {
  try {
    Logger.log('=== deleteOTApplication START ===');
    Logger.log('Deleting row: ' + rowNumber);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('OT_Applications');
    
    if (!sheet) {
      return { success: false, message: 'OT_Applications sheet not found' };
    }
    
    const rowNum = parseInt(rowNumber);
    // Delete the row
    sheet.deleteRow(rowNum);
    // Ensure deletion is applied
    SpreadsheetApp.flush();

    Logger.log('Application deleted successfully (row ' + rowNum + ')');
    return { success: true, message: 'Application deleted successfully', deletedRow: rowNum };
    
  } catch (error) {
    Logger.log('Error in deleteOTApplication: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ======== TEST FUNCTION ========
function testStaffDashboard() {
  const testEmail = 'alihariz@malaysiaairports.com.my';
  Logger.log('=== TESTING getStaffDashboard ===');
  Logger.log('Test email: ' + testEmail);

  const result = getStaffDashboard(testEmail);

  Logger.log('=== RESULT ===');
  Logger.log('Result is null: ' + (result === null));
  Logger.log('Result is undefined: ' + (result === undefined));
  Logger.log('Result type: ' + typeof result);
  Logger.log('Total OT Hours: ' + (result ? result.totalOTHours : 'N/A'));
  Logger.log('Remaining Hours: ' + (result ? result.remainingHours : 'N/A'));
  Logger.log('Applications count: ' + (result && result.applications ? result.applications.length : 'N/A'));
  Logger.log('Pending count: ' + (result ? result.pendingCount : 'N/A'));

  if (result && result.applications && result.applications.length > 0) {
    Logger.log('\n=== APPLICATIONS RETURNED ===');
    result.applications.forEach((app, idx) => {
      Logger.log('App ' + (idx+1) + ': ' + app.Name + ' | Date: ' + app.Date + ' | Hours: ' + app.TotalHours + ' | Status: ' + app.Status);
    });
  } else {
    Logger.log('⚠️ NO APPLICATIONS RETURNED');
  }

  return result;
}

// ======== TEST MANAGEMENT DASHBOARD ========
function testManagementDashboard() {
  Logger.log('=== TESTING getManagementDashboard ===');

  const result = getManagementDashboard();

  Logger.log('=== RESULT ===');
  Logger.log('Total Staff: ' + result.totalStaff);
  Logger.log('Total OT Hours: ' + result.totalOTHours);
  Logger.log('Exceed Count: ' + result.exceedCount);
  Logger.log('Pending Approvals: ' + result.pendingApprovals);
  Logger.log('Applications count: ' + (result.applications ? result.applications.length : 'N/A'));

  if (result.applications && result.applications.length > 0) {
    Logger.log('\n=== FIRST 5 APPLICATIONS ===');
    for (let i = 0; i < Math.min(result.applications.length, 5); i++) {
      const app = result.applications[i];
      Logger.log((i+1) + '. ' + app.Name + ' | Status: ' + app.Status + ' | Hours: ' + app.TotalHours);
    }
  }

  return result;
}

// ======== MANAGEMENT DASHBOARD ========
// WORKFLOW: Fetches ALL OT applications across all staff
// CRITICAL: Must show newly submitted Pending applications immediately
function getManagementDashboard() {
  const callTime = new Date().toISOString();
  Logger.log('=== getManagementDashboard START at ' + callTime + ' ===');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return { error: 'Cannot access database' };
    }
    
    // Use OTCalc module for aggregated statistics
    const allStats = calculateAllUsersOTStats();
    
    if (!allStats.success) {
      Logger.log('ERROR: Failed to calculate OT stats - ' + allStats.error);
      return {
        totalStaff: 0,
        totalOTHours: 0,
        exceedCount: 0,
        pendingApprovals: 0,
        applications: []
      };
    }
    
    Logger.log('OT Stats from module: totalOTHours=' + allStats.totalOTHours + ', staffCount=' + allStats.staffCount + ', exceedCount=' + allStats.exceedCount);
    
    // Retrieve applications for display
    const otSheet = ss.getSheetByName(SHEET_OT);
    
    if (!otSheet) {
      Logger.log('ERROR: OT_Applications sheet not found');
      return {
        totalStaff: allStats.staffCount,
        totalOTHours: allStats.totalOTHours,
        exceedCount: allStats.exceedCount,
        pendingApprovals: 0,
        applications: []
      };
    }
    
    const otData = otSheet.getDataRange().getValues().slice(1);
    
    // Filter out empty rows and rows with invalid data
    const validOtData = otData.filter(row => {
      return row[0] && row[0] !== '' && row.length >= 6;
    });
    
    Logger.log('Total OT rows: ' + otData.length + ', Valid rows: ' + validOtData.length);
    Logger.log('--- CHECKING PENDING APPLICATIONS ---');
    
    // Count pending applications
    let pendingApplicationsCount = 0;
    
    validOtData.forEach((row, idx) => {
      const name = row[0];
      const hours = parseFloat(row[5]);
      const status = row[10] ? row[10].toString().trim() : 'Pending';
      
      // Log each row for debugging
      Logger.log('Row ' + (idx + 2) + ': Name="' + name + '" Hours=' + hours + ' Status="' + status + '"');
      
      // CRITICAL: Count pending with exact match
      if (status === 'Pending') {
        pendingApplicationsCount++;
        Logger.log('  ✅ Counted as PENDING');
      }
    });
    
    Logger.log('Total Pending Applications: ' + pendingApplicationsCount);
    Logger.log('Staff Exceeding Limit (from module): ' + allStats.exceedCount);
    
    // Prepare applications array
    const applications = validOtData.map(row => {
      // Convert date to string format for serialization
      let dateStr = '';
      if (row[2]) {
        if (row[2] instanceof Date) {
          dateStr = Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          dateStr = row[2].toString();
        }
      }

      // Format time fields as HH:MM
      let startTimeStr = '';
      if (row[3]) {
        if (row[3] instanceof Date) {
          startTimeStr = Utilities.formatDate(row[3], Session.getScriptTimeZone(), 'HH:mm');
        } else {
          startTimeStr = row[3].toString();
        }
      }

      let endTimeStr = '';
      if (row[4]) {
        if (row[4] instanceof Date) {
          endTimeStr = Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'HH:mm');
        } else {
          endTimeStr = row[4].toString();
        }
      }

      return {
        Name: row[0] || '',
        Team: row[1] || '',
        Date: dateStr,
        StartTime: startTimeStr,
        EndTime: endTimeStr,
        TotalHours: row[5] || 0,
        OT_Type: row[6] || '',
        Is_Public_Holiday: row[7] || '',
        Proof_Attendance: row[8] || '',
        Reason: row[9] || '',
        Status: row[10] ? row[10].toString() : '',
        ApprovedBy: row[11] || '',
        Approval_Date: row[12] ? row[12].toString() : ''
      };
    });
    
    Logger.log('=== getManagementDashboard END ===');
    Logger.log('Returning: totalStaff=' + allStats.staffCount + ', totalOTHours=' + allStats.totalOTHours.toFixed(1) + ', exceedCount=' + allStats.exceedCount + ', pendingApprovals=' + pendingApplicationsCount);
    Logger.log('Total applications to return: ' + applications.length);
    Logger.log('First 3 applications:');
    for (let i = 0; i < Math.min(applications.length, 3); i++) {
      Logger.log('  App ' + (i+1) + ': ' + applications[i].Name + ' Status="' + applications[i].Status + '"');
    }

    return {
      totalStaff: allStats.staffCount,
      totalOTHours: allStats.totalOTHours,
      exceedCount: allStats.exceedCount,
      pendingApprovals: pendingApplicationsCount,
      applicationCount: applications.length,
      applications: applications
    };
  } catch (error) {
    Logger.log('Error in getManagementDashboard: ' + error.toString());
    return {
      totalStaff: 0,
      totalOTHours: 0,
      exceedCount: 0,
      pendingApprovals: 0,
      applicationCount: 0,
      applications: []
    };
  }
}

// ======== OT APPLICATION SUBMISSION ========
// WORKFLOW: Form Submit → Save to Sheet → Status='Pending' → Dashboard Refresh
// ISSUE FIXED: Status was not explicitly set, causing dashboard filter to miss records
function submitOTApplication(formData) {
  Logger.log('=== submitOTApplication START ===');
  Logger.log('Form data received: ' + JSON.stringify(formData));
  
  // Guard: Cannot be called directly from script editor
  if (!formData) {
    Logger.log('ERROR: formData is undefined. This function must be called from the web app form.');
    return { success: false, message: 'This function cannot be run directly. Please use the web app form.' };
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return { success: false, error: 'Cannot access database' };
    }
    
    const otSheet = ss.getSheetByName(SHEET_OT);
    
    if (!otSheet) {
      Logger.log('ERROR: OT_Applications sheet not found');
      return { success: false, message: 'OT Applications sheet not found' };
    }
    
    // CRITICAL: Verify sheet has 13 columns (A-M) before proceeding
    const numCols = otSheet.getLastColumn();
    if (numCols < 13) {
      Logger.log('ERROR: Sheet structure invalid - has ' + numCols + ' columns, needs 13');
      Logger.log('CRITICAL: Missing columns will cause data misalignment!');
      return { 
        success: false, 
        message: 'System Error: Sheet structure invalid. Please contact admin to run fixSheetIssues()' 
      };
    }
    
    // Validate required fields
    if (!formData.name || !formData.date || !formData.totalHours) {
      Logger.log('ERROR: Missing required fields');
      return { success: false, message: 'Missing required fields' };
    }
    
    // Use OTCalc module to validate if user can apply for requested hours
    const requestedHours = parseFloat(formData.totalHours);
    const userEmail = Session.getActiveUser().getEmail();
    
    Logger.log('Validating OT application: email=' + userEmail + ', requestedHours=' + requestedHours);
    
    const validation = validateOTApplication(userEmail, requestedHours);
    
    if (!validation.valid) {
      Logger.log('Validation failed: ' + validation.reason);
      return { 
        success: false, 
        message: validation.reason,
        remainingHours: validation.remainingHours,
        totalHours: validation.totalHours
      };
    }
    
    Logger.log('Validation passed - user can apply for ' + requestedHours + ' hours (remaining: ' + validation.remainingHours + ')');
    
    // Prepare row data matching sheet structure
    // A: Name, B: Team, C: Date, D: StartTime, E: EndTime, F: TotalHours, 
    // G: OT_Type, H: Is_Public_Holiday, I: Proof_Attendance, J: Reason, 
    // K: Status, L: ApprovedBy, M: Approval_Date
    
    // CRITICAL FIX: Trim the name to ensure consistency
    const cleanName = (formData.name || '').toString().trim();
    
    // CRITICAL: Always set Status to 'Pending' for new submissions
    // This was the PRIMARY ISSUE - status must be exact string 'Pending' for dashboard filter
    const newRow = [
      cleanName,                       // A: Name (trimmed)
      formData.team || '',             // B: Team
      formData.date || '',             // C: Date
      formData.startTime || '',        // D: Start Time
      formData.endTime || '',          // E: End Time
      requestedHours,                  // F: Total Hours (validated)
      formData.otType || '',           // G: OT Type
      formData.isPublicHoliday || 'No',// H: Is Public Holiday
      formData.proofAttendance || '',  // I: Proof of Attendance
      formData.reason || '',           // J: Reason
      'Pending',                       // K: Status (ALWAYS 'Pending' for new submissions)
      '',                              // L: Approved By (empty initially)
      ''                               // M: Approval Date (empty initially)
    ];
    
    Logger.log('Appending row to sheet: ' + JSON.stringify(newRow));
    Logger.log('CRITICAL CHECK - Status field (Column K): "' + newRow[10] + '"');
    
    // Append to sheet
    otSheet.appendRow(newRow);
    
    // CRITICAL: Force immediate flush to ensure data is written before redirect
    SpreadsheetApp.flush();
    
    // Verify the row was actually added
    const lastRow = otSheet.getLastRow();
    const savedData = otSheet.getRange(lastRow, 1, 1, 13).getValues()[0];
    const savedStatus = otSheet.getRange(lastRow, 11).getValue(); // Column K
    Logger.log('===== POST-SAVE VERIFICATION =====');
    Logger.log('Row ' + lastRow + ' saved with data:');
    Logger.log('  Name: "' + savedData[0] + '" (trimmed length: ' + (savedData[0] ? savedData[0].toString().trim().length : 0) + ')');
    Logger.log('  Team: "' + savedData[1] + '"');
    Logger.log('  Date: "' + savedData[2] + '"');
    Logger.log('  Hours: ' + savedData[5]);
    Logger.log('  Status: "' + savedData[10] + '" (CRITICAL: Must be "Pending")');
    Logger.log('Verification: Row ' + lastRow + ' Status = "' + savedStatus + '"');
    
    Logger.log('✅ OT Application submitted successfully!');
    Logger.log('  Row Number: ' + lastRow);
    Logger.log('  Name: ' + cleanName);
    Logger.log('  Name length: ' + cleanName.length);
    Logger.log('  Name charCodes: ' + Array.from(cleanName).map(c => c.charCodeAt(0)).join(','));
    Logger.log('  Team: ' + formData.team);
    Logger.log('  Date: ' + formData.date);
    Logger.log('  Hours: ' + formData.totalHours);
    Logger.log('  Status saved: "' + savedStatus + '" (MUST be "Pending")');
    Logger.log('=== WORKFLOW: Data saved → User will be redirected → Dashboard will reload ===');
    
    return { 
      success: true, 
      message: 'OT Application submitted successfully!',
      rowNumber: lastRow,
      status: savedStatus
    };
    
  } catch (error) {
    Logger.log('ERROR in submitOTApplication: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return { success: false, message: 'Error: ' + error.message };
  }
}

// ======== UPDATE APPLICATION STATUS ========
function updateApplicationStatus(employeeName, applicationDate, newStatus, approverName) {
  Logger.log('=== updateApplicationStatus START ===');
  Logger.log('Name: ' + employeeName + ', Date: ' + applicationDate + ', Status: ' + newStatus);
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return { success: false, error: 'Cannot access database' };
    }
    
    const otSheet = ss.getSheetByName(SHEET_OT);
    
    if (!otSheet) {
      Logger.log('ERROR: OT_Applications sheet not found');
      return { success: false, message: 'OT Applications sheet not found' };
    }
    
    const data = otSheet.getDataRange().getValues();
    
    // Find the application row (skip header row)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[0]; // Column A: Name
      const date = row[2]; // Column C: Date
      
      // Convert date to string for comparison
      let dateStr = '';
      if (date instanceof Date) {
        dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        dateStr = date.toString().split('T')[0];
      }
      
      // Match by name and date
      if (name === employeeName && dateStr === applicationDate) {
        const rowNumber = i + 1; // Sheet rows are 1-indexed
        const approvalDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        
        Logger.log('Found match at row ' + rowNumber);
        
        // Update Status (Column K), ApprovedBy (Column L), Approval_Date (Column M)
        otSheet.getRange(rowNumber, 11).setValue(newStatus); // Column K
        otSheet.getRange(rowNumber, 12).setValue(approverName); // Column L
        otSheet.getRange(rowNumber, 13).setValue(approvalDate); // Column M
        
        SpreadsheetApp.flush();
        
        // Verify the update was saved
        const savedStatus = otSheet.getRange(rowNumber, 11).getValue();
        Logger.log('Post-update verification: Status = "' + savedStatus + '" (expected "' + newStatus + '")');
        
        Logger.log(`Application ${newStatus} successfully!`);
        return { 
          success: true, 
          message: `Application ${newStatus.toLowerCase()} successfully!`,
          updatedRow: rowNumber,
          newStatus: savedStatus
        };
      }
    }
    
    Logger.log('ERROR: Application not found');
    return { 
      success: false, 
      message: 'Application not found' 
    };
    
  } catch (error) {
    Logger.log('ERROR in updateApplicationStatus: ' + error.toString());
    return { 
      success: false, 
      message: 'Error: ' + error.message 
    };
  }
}

// ======== DEBUG FUNCTION ========
function debugOTData(userName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return { success: false, error: 'Cannot access database' };
    }
    
    const otSheet = ss.getSheetByName(SHEET_OT);
    const data = otSheet.getDataRange().getValues();
    
    Logger.log('=== DEBUG OT DATA ===');
    Logger.log('Total rows (including header): ' + data.length);
    Logger.log('Looking for user: ' + userName);
    Logger.log('');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      Logger.log('Row ' + (i+1) + ':');
      Logger.log('  Name: ' + row[0]);
      Logger.log('  Team: ' + row[1]);
      Logger.log('  Date: ' + row[2]);
      Logger.log('  Hours: ' + row[5]);
      Logger.log('  Status: ' + row[10]);
      
      if (row[0] === userName) {
        Logger.log('  ✓ MATCH FOUND!');
      }
      Logger.log('');
    }
    
    return { success: true, message: 'Check logs for details' };
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return { success: false, message: error.message };
  }
}

// ======== HELPER FUNCTION ========
function calcTotalHours(startTime, endTime) {
  const start = new Date(startTime);
  const end = new Date(endTime);
  const diffMs = end - start;
  return diffMs / (1000 * 60 * 60);
}

// ======== OT CLAIM SUBMISSION ========
function submitOTClaim(formData) {
  Logger.log('=== submitOTClaim START ===');
  Logger.log('Form data: ' + JSON.stringify(formData));
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return { success: false, error: 'Cannot access database' };
    }
    
    const claimSheet = ss.getSheetByName('OT_Claim');
    
    if (!claimSheet) {
      Logger.log('ERROR: OT_Claim sheet not found');
      return { success: false, message: 'OT_Claim sheet not found. Please create it first.' };
    }
    
    // Validate required fields
    if (!formData.name || !formData.month || !formData.totalOTHours) {
      Logger.log('ERROR: Missing required fields');
      return { success: false, message: 'Missing required fields' };
    }
    
    // Prepare row data matching sheet structure
    // A: Name, B: Month, C: TotalOTHours, D: ClaimAmount, E: ClaimDate, F: Status
    const newRow = [
      formData.name || '',
      formData.month || '',
      formData.totalOTHours || 0,
      formData.claimAmount || 0,
      formData.claimDate || new Date().toISOString().split('T')[0],
      formData.status || 'Pending'
    ];
    
    Logger.log('Appending row to OT_Claim sheet: ' + JSON.stringify(newRow));
    
    claimSheet.appendRow(newRow);
    SpreadsheetApp.flush();
    
    Logger.log('OT Claim submitted successfully!');
    return { success: true, message: 'OT Claim submitted successfully!' };
    
  } catch (error) {
    Logger.log('ERROR in submitOTClaim: ' + error.toString());
    return { success: false, message: 'Error: ' + error.message };
  }
}

// ======== LEAVE REQUEST SUBMISSION ========
function submitLeaveRequest(formData) {
  Logger.log('=== submitLeaveRequest START ===');
  Logger.log('Form data: ' + JSON.stringify(formData));
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return { success: false, error: 'Cannot access database' };
    }
    
    let leaveSheet = ss.getSheetByName('Leave_Applications');
    
    // Create sheet if it doesn't exist
    if (!leaveSheet) {
      Logger.log('Leave_Applications sheet not found. Creating new sheet...');
      leaveSheet = ss.insertSheet('Leave_Applications');
      
      // Add headers
      const headers = ['Name', 'Team', 'Email', 'Leave Type', 'Start Date', 'End Date', 
                       'Total Days', 'Reason', 'Supporting Doc', 'Status', 'Approved By', 
                       'Submitted Date', 'Approval Date'];
      leaveSheet.appendRow(headers);
      
      // Format header row
      const headerRange = leaveSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4a5568');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      
      Logger.log('Leave_Applications sheet created with headers');
    }
    
    // Validate required fields
    if (!formData.name || !formData.leaveType || !formData.startDate || !formData.endDate) {
      Logger.log('ERROR: Missing required fields');
      return { success: false, message: 'Missing required fields' };
    }
    
    // Prepare row data
    const newRow = [
      formData.name || '',
      formData.team || '',
      formData.email || '',
      formData.leaveType || '',
      formData.startDate || '',
      formData.endDate || '',
      formData.totalDays || 0,
      formData.reason || '',
      formData.supportingDoc || '',
      formData.status || 'Pending',
      '', // Approved By (empty initially)
      formData.submittedDate || new Date().toISOString().split('T')[0],
      '' // Approval Date (empty initially)
    ];
    
    Logger.log('Appending row to Leave_Applications sheet: ' + JSON.stringify(newRow));
    
    leaveSheet.appendRow(newRow);
    SpreadsheetApp.flush();
    
    Logger.log('Leave request submitted successfully!');
    return { success: true, message: 'Leave request submitted successfully!' };
    
  } catch (error) {
    Logger.log('ERROR in submitLeaveRequest: ' + error.toString());
    return { success: false, message: 'Error: ' + error.message };
  }
}

// ======== DIAGNOSTIC FUNCTION ========
// Run this function to check the current state of OT_Applications sheet
function diagnoseOTData() {
  Logger.log('====== DIAGNOSTIC CHECK ======');
  Logger.log('Time: ' + new Date().toISOString());
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const otSheet = ss.getSheetByName(SHEET_OT);
  
  Logger.log('Sheet found: ' + (otSheet ? 'YES' : 'NO'));
  
  if (!otSheet) {
    Logger.log('ERROR: OT_Applications sheet not found!');
    return;
  }
  
  const data = otSheet.getDataRange().getValues();
  Logger.log('Total rows (including header): ' + data.length);
  Logger.log('Total columns: ' + data[0].length);
  
  Logger.log('\n--- HEADER ROW ---');
  Logger.log(JSON.stringify(data[0]));
  
  Logger.log('\n--- ALL DATA ROWS ---');
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) {
      Logger.log('Row ' + (i+1) + ': EMPTY');
      continue;
    }
    
    Logger.log('\nRow ' + (i+1) + ':');
    Logger.log('  Name: "' + row[0] + '"');
    Logger.log('  Team: "' + row[1] + '"');
    Logger.log('  Date: "' + row[2] + '"');
    Logger.log('  Hours: ' + row[5]);
    Logger.log('  Status: "' + row[10] + '"');
    Logger.log('  Status === "Pending": ' + (row[10] === 'Pending'));
    Logger.log('  Status trimmed: "' + (row[10] ? row[10].toString().trim() : '') + '"');
  }
  
  Logger.log('\n====== END DIAGNOSTIC ======');
}

// ======== FIX BLANK STATUS VALUES ========
// Run this ONCE to fix all existing rows with blank Status
function fixBlankStatusValues() {
  Logger.log('====== FIXING BLANK STATUS VALUES ======');
  Logger.log('Time: ' + new Date().toISOString());
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const otSheet = ss.getSheetByName(SHEET_OT);
  
  if (!otSheet) {
    Logger.log('ERROR: OT_Applications sheet not found!');
    return;
  }
  
  const data = otSheet.getDataRange().getValues();
  let fixedCount = 0;
  
  Logger.log('Scanning ' + (data.length - 1) + ' data rows...');
  
  // Start from row 2 (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNumber = i + 1;
    
    // Skip empty rows
    if (!row[0]) {
      continue;
    }
    
    // Check if Status (Column K, index 10) is blank
    const currentStatus = row[10];
    
    if (!currentStatus || currentStatus === '' || currentStatus.toString().trim() === '') {
      Logger.log('Row ' + rowNumber + ': Status is BLANK - Setting to "Pending"');
      
      // Update Status column (K = column 11)
      otSheet.getRange(rowNumber, 11).setValue('Pending');
      fixedCount++;
    } else {
      Logger.log('Row ' + rowNumber + ': Status = "' + currentStatus + '" - OK');
    }
  }
  
  if (fixedCount > 0) {
    SpreadsheetApp.flush();
    Logger.log('\n✅ SUCCESS: Fixed ' + fixedCount + ' rows with blank Status');
    Logger.log('All blank Status values have been set to "Pending"');
  } else {
    Logger.log('\n✅ No blank Status values found - all rows are OK');
  }
  
  Logger.log('====== FIX COMPLETE ======');
}

// ======== MONTHLY ARCHIVE FUNCTION ========
/**
 * Archives all OT data for the current month and resets the system for a new month
 * This should be run at the end of each month (manually or via time-based trigger)
 *
 * Process:
 * 1. Get all staff from Users sheet
 * 2. For each staff, collect their OT applications
 * 3. Calculate monthly statistics (total hours, counts by status)
 * 4. Store applications as JSON in OT_History sheet
 * 5. Archive current OT_Applications (move to history)
 * 6. Reset remaining hours to 104 for all staff
 *
 * @returns {Object} Result object with success status and details
 */
function monthlyArchive() {
  Logger.log('=== MONTHLY ARCHIVE START ===');
  Logger.log('Archive Time: ' + new Date().toISOString());

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // Get all required sheets
    const usersSheet = ss.getSheetByName('Users');
    const otSheet = ss.getSheetByName('OT_Applications');
    const historySheet = ss.getSheetByName('OT_History');

    if (!usersSheet || !otSheet) {
      Logger.log('ERROR: Required sheets not found');
      return { success: false, error: 'Missing required sheets (Users or OT_Applications)' };
    }

    // Create history sheet if it doesn't exist
    if (!historySheet) {
      Logger.log('OT_History sheet not found - running setup');
      setupOTHistorySheet(ss);
    }

    const historySheetFinal = ss.getSheetByName('OT_History');

    // Get current month and year
    const now = new Date();
    const currentMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
    const currentYear = now.getFullYear();

    Logger.log('Archiving data for: ' + currentMonth);

    // Get all users (skip header row)
    const usersData = usersSheet.getDataRange().getValues();
    const headers = usersData[0];
    const users = usersData.slice(1);

    Logger.log('Found ' + users.length + ' users to process');

    // Get all OT applications
    const otData = otSheet.getDataRange().getValues();
    const otHeaders = otData[0];
    const applications = otData.slice(1);

    Logger.log('Found ' + applications.length + ' OT applications');

    let archivedCount = 0;

    // Process each staff member
    for (let i = 0; i < users.length; i++) {
      const user = users[i];
      const staffName = user[0];  // Column A: Name
      const staffEmail = user[1]; // Column B: Email
      const role = user[2];        // Column C: Role
      const team = user[3];        // Column D: Team

      // Skip if not staff or if empty row
      if (!staffName || !staffEmail) {
        continue;
      }

      Logger.log('\nProcessing: ' + staffName + ' (' + staffEmail + ')');

      // Find all applications for this staff
      const staffApplications = [];
      let totalHours = 0;
      let approvedCount = 0;
      let rejectedCount = 0;
      let pendingCount = 0;

      for (let j = 0; j < applications.length; j++) {
        const app = applications[j];
        const appName = app[0];  // Name

        // Skip empty rows
        if (!appName) continue;

        // Check if this application belongs to current staff
        if (appName.trim() === staffName.trim()) {
          // Build application object
          const applicationObj = {
            Name: app[0],
            Team: app[1],
            Date: app[2] ? Utilities.formatDate(new Date(app[2]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
            StartTime: app[3] ? (app[3] instanceof Date ? Utilities.formatDate(app[3], Session.getScriptTimeZone(), 'HH:mm') : app[3].toString()) : '',
            EndTime: app[4] ? (app[4] instanceof Date ? Utilities.formatDate(app[4], Session.getScriptTimeZone(), 'HH:mm') : app[4].toString()) : '',
            TotalHours: parseFloat(app[5]) || 0,
            OT_Type: app[6] || '',
            IsPublicHoliday: app[7] || '',
            ProofAttendance: app[8] || '',
            Reason: app[9] || '',
            Status: app[10] || 'Pending',
            ApprovedBy: app[11] || '',
            ApprovalDate: app[12] ? Utilities.formatDate(new Date(app[12]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
          };

          staffApplications.push(applicationObj);

          // Calculate statistics
          const hours = parseFloat(app[5]) || 0;
          totalHours += hours;

          const status = (app[10] || 'Pending').toString().trim();
          if (status === 'Approved') {
            approvedCount++;
          } else if (status === 'Rejected') {
            rejectedCount++;
          } else {
            pendingCount++;
          }
        }
      }

      Logger.log('  Applications: ' + staffApplications.length);
      Logger.log('  Total Hours: ' + totalHours);
      Logger.log('  Approved: ' + approvedCount + ', Rejected: ' + rejectedCount + ', Pending: ' + pendingCount);

      // Skip staff with no OT applications for this month
      if (staffApplications.length === 0) {
        Logger.log('  ⏭️ Skipping - No applications for this month');
        continue;
      }

      // Convert applications to JSON
      const applicationsJSON = JSON.stringify(staffApplications);

      // Prepare row for OT_History sheet
      const historyRow = [
        currentMonth,              // Month (e.g., "2025-11")
        currentYear,               // Year
        staffName,                 // Staff_Name
        staffEmail,                // Staff_Email
        team || '',                // Team
        totalHours,                // Total_OT_Hours
        staffApplications.length,  // Applications_Count
        approvedCount,             // Approved_Count
        rejectedCount,             // Rejected_Count
        pendingCount,              // Pending_Count
        applicationsJSON,          // Applications_JSON
        new Date()                 // Created_Date
      ];

      // Append to history sheet
      historySheetFinal.appendRow(historyRow);
      archivedCount++;

      Logger.log('  ✅ Archived to OT_History');
    }

    Logger.log('\n=== CLEARING OT_Applications SHEET ===');

    // Clear all OT applications (keep header)
    if (applications.length > 0) {
      const lastRow = otSheet.getLastRow();
      if (lastRow > 1) {
        otSheet.deleteRows(2, lastRow - 1);
        Logger.log('Cleared ' + (lastRow - 1) + ' rows from OT_Applications');
      }
    }

    Logger.log('\n=== RESETTING STAFF REMAINING HOURS ===');

    // Reset remaining hours to 104 for all staff
    // Note: The remaining hours are calculated dynamically from applications
    // So clearing applications automatically resets remaining hours
    Logger.log('Remaining hours automatically reset (no applications = 104 hours available)');

    Logger.log('\n=== MONTHLY ARCHIVE COMPLETE ===');
    Logger.log('✅ Successfully archived ' + archivedCount + ' staff records for ' + currentMonth);
    Logger.log('✅ OT_Applications sheet cleared');
    Logger.log('✅ System ready for new month');

    return {
      success: true,
      message: 'Monthly archive completed successfully',
      details: {
        month: currentMonth,
        staffArchived: archivedCount,
        applicationsCleared: applications.length
      }
    };

  } catch (error) {
    Logger.log('ERROR in monthlyArchive: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ======== GET STAFF HISTORY FUNCTION ========
/**
 * Retrieves historical OT data for a specific staff member
 * @param {string} email - Staff email address
 * @returns {Object} Object containing historical data for the staff member
 */
function getStaffHistory(email) {
  Logger.log('=== GET STAFF HISTORY ===');
  Logger.log('Email: ' + email);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const historySheet = ss.getSheetByName('OT_History');

    if (!historySheet) {
      Logger.log('OT_History sheet not found');
      return {
        success: true,
        history: [],
        message: 'No history available yet'
      };
    }

    // Get all history data
    const data = historySheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    Logger.log('Total history records: ' + rows.length);

    // Filter records for this staff member
    const staffHistory = [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const month = row[0];          // Month
      const year = row[1];           // Year
      const staffName = row[2];      // Staff_Name
      const staffEmail = row[3];     // Staff_Email
      const team = row[4];           // Team
      const totalHours = row[5];     // Total_OT_Hours
      const appsCount = row[6];      // Applications_Count
      const approvedCount = row[7];  // Approved_Count
      const rejectedCount = row[8];  // Rejected_Count
      const pendingCount = row[9];   // Pending_Count
      const appsJSON = row[10];      // Applications_JSON
      const createdDate = row[11];   // Created_Date

      // Check if this record belongs to the staff
      if (staffEmail && staffEmail.toString().trim() === email.trim()) {
        // Parse applications JSON
        let applications = [];
        try {
          if (appsJSON) {
            applications = JSON.parse(appsJSON);
          }
        } catch (e) {
          Logger.log('Error parsing JSON for month ' + month + ': ' + e.toString());
          applications = [];
        }

        // Build history record
        const historyRecord = {
          Month: month,
          Year: year,
          StaffName: staffName,
          Team: team,
          TotalHours: parseFloat(totalHours) || 0,
          ApplicationsCount: parseInt(appsCount) || 0,
          ApprovedCount: parseInt(approvedCount) || 0,
          RejectedCount: parseInt(rejectedCount) || 0,
          PendingCount: parseInt(pendingCount) || 0,
          Applications: applications,
          ArchivedDate: createdDate ? Utilities.formatDate(new Date(createdDate), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : ''
        };

        staffHistory.push(historyRecord);
      }
    }

    // Sort by month (newest first)
    staffHistory.sort((a, b) => {
      return b.Month.localeCompare(a.Month);
    });

    Logger.log('Found ' + staffHistory.length + ' historical records for ' + email);

    return {
      success: true,
      history: staffHistory,
      totalRecords: staffHistory.length
    };

  } catch (error) {
    Logger.log('ERROR in getStaffHistory: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      history: []
    };
  }
}

// ======== GET MANAGEMENT HISTORY FUNCTION ========
/**
 * Retrieves historical OT data for management analytics
 * Can filter by date range or return all history
 * @param {string} startMonth - Optional start month in YYYY-MM format
 * @param {string} endMonth - Optional end month in YYYY-MM format
 * @returns {Object} Object containing aggregated historical data
 */
function getManagementHistory(startMonth, endMonth) {
  Logger.log('=== GET MANAGEMENT HISTORY ===');
  Logger.log('Start Month: ' + (startMonth || 'All'));
  Logger.log('End Month: ' + (endMonth || 'All'));

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const historySheet = ss.getSheetByName('OT_History');

    if (!historySheet) {
      Logger.log('OT_History sheet not found');
      return {
        success: true,
        history: [],
        summary: {},
        message: 'No history available yet'
      };
    }

    // Get all history data
    const data = historySheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    Logger.log('Total history records: ' + rows.length);

    // Process all records
    const historyRecords = [];
    let totalOTHours = 0;
    let totalApplications = 0;
    let totalApproved = 0;
    let totalRejected = 0;
    let totalPending = 0;
    const monthlyStats = {};
    const staffStats = {};

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const month = row[0];          // Month
      const year = row[1];           // Year
      const staffName = row[2];      // Staff_Name
      const staffEmail = row[3];     // Staff_Email
      const team = row[4];           // Team
      const totalHours = parseFloat(row[5]) || 0;
      const appsCount = parseInt(row[6]) || 0;
      const approvedCount = parseInt(row[7]) || 0;
      const rejectedCount = parseInt(row[8]) || 0;
      const pendingCount = parseInt(row[9]) || 0;
      const appsJSON = row[10];      // Applications_JSON
      const createdDate = row[11];   // Created_Date

      // Skip empty rows
      if (!month || !staffName) continue;

      // Apply date range filter if provided
      if (startMonth && month < startMonth) continue;
      if (endMonth && month > endMonth) continue;

      // Parse applications JSON
      let applications = [];
      try {
        if (appsJSON) {
          applications = JSON.parse(appsJSON);
        }
      } catch (e) {
        Logger.log('Error parsing JSON for ' + staffName + ' / ' + month + ': ' + e.toString());
      }

      // Build record
      const record = {
        Month: month,
        Year: year,
        StaffName: staffName,
        StaffEmail: staffEmail,
        Team: team,
        TotalHours: totalHours,
        ApplicationsCount: appsCount,
        ApprovedCount: approvedCount,
        RejectedCount: rejectedCount,
        PendingCount: pendingCount,
        Applications: applications,
        ArchivedDate: createdDate ? Utilities.formatDate(new Date(createdDate), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : ''
      };

      historyRecords.push(record);

      // Aggregate totals
      totalOTHours += totalHours;
      totalApplications += appsCount;
      totalApproved += approvedCount;
      totalRejected += rejectedCount;
      totalPending += pendingCount;

      // Monthly aggregation
      if (!monthlyStats[month]) {
        monthlyStats[month] = {
          month: month,
          totalHours: 0,
          totalApplications: 0,
          staffCount: 0,
          approvedCount: 0,
          rejectedCount: 0,
          pendingCount: 0
        };
      }
      monthlyStats[month].totalHours += totalHours;
      monthlyStats[month].totalApplications += appsCount;
      monthlyStats[month].staffCount += 1;
      monthlyStats[month].approvedCount += approvedCount;
      monthlyStats[month].rejectedCount += rejectedCount;
      monthlyStats[month].pendingCount += pendingCount;

      // Staff aggregation
      if (!staffStats[staffEmail]) {
        staffStats[staffEmail] = {
          name: staffName,
          email: staffEmail,
          team: team,
          totalHours: 0,
          totalApplications: 0,
          monthsActive: 0,
          approvedCount: 0,
          rejectedCount: 0,
          pendingCount: 0
        };
      }
      staffStats[staffEmail].totalHours += totalHours;
      staffStats[staffEmail].totalApplications += appsCount;
      staffStats[staffEmail].monthsActive += 1;
      staffStats[staffEmail].approvedCount += approvedCount;
      staffStats[staffEmail].rejectedCount += rejectedCount;
      staffStats[staffEmail].pendingCount += pendingCount;
    }

    // Sort records by month (newest first)
    historyRecords.sort((a, b) => b.Month.localeCompare(a.Month));

    // Convert stats objects to arrays
    const monthlyStatsArray = Object.values(monthlyStats).sort((a, b) => b.month.localeCompare(a.month));
    const staffStatsArray = Object.values(staffStats).sort((a, b) => b.totalHours - a.totalHours);

    Logger.log('Processed ' + historyRecords.length + ' records');
    Logger.log('Total OT Hours: ' + totalOTHours);

    return {
      success: true,
      history: historyRecords,
      summary: {
        totalOTHours: totalOTHours,
        totalApplications: totalApplications,
        totalApproved: totalApproved,
        totalRejected: totalRejected,
        totalPending: totalPending,
        recordCount: historyRecords.length
      },
      monthlyStats: monthlyStatsArray,
      staffStats: staffStatsArray
    };

  } catch (error) {
    Logger.log('ERROR in getManagementHistory: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      history: [],
      summary: {}
    };
  }
}

