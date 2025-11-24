// ============================================
// OT CALCULATION MODULE
// ============================================
// Purpose: Calculate OT hours, remaining hours, and statistics for users
// Used by: Staff Dashboard, OT Application Form, Management Dashboard

/**
 * Calculate OT statistics for a specific user by email
 * @param {string} email - User's email address
 * @returns {Object} OT statistics object
 */
function calculateUserOTStats(email) {
  Logger.log('=== calculateUserOTStats START ===');
  Logger.log('Email: ' + email);
  
  try {
    // Open spreadsheet by ID
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return {
        success: false,
        error: 'Cannot access database',
        totalOTHours: 0,
        remainingHours: 104,
        remainingColor: 'green'
      };
    }
    
    const usersSheet = ss.getSheetByName(CONFIG.SHEET_STAFF);
    const otSheet = ss.getSheetByName(CONFIG.SHEET_OT);
    
    if (!usersSheet || !otSheet) {
      Logger.log('ERROR: Required sheets not found');
      return {
        success: false,
        error: 'Required sheets not found',
        totalOTHours: 0,
        remainingHours: 104,
        remainingColor: 'green'
      };
    }
    
    // 1. Find user's name from email
    const usersData = usersSheet.getDataRange().getValues();
    let userName = null;
    const emailToFind = (email || '').toString().trim().toLowerCase();
    
    Logger.log('Searching for email: "' + emailToFind + '" in Users sheet');
    Logger.log('Total users to scan: ' + (usersData.length - 1));
    
    for (let i = 1; i < usersData.length; i++) {
      const sheetEmail = (usersData[i][1] || '').toString().trim().toLowerCase();
      if (sheetEmail === emailToFind) { // Column B = Email
        userName = usersData[i][0];    // Column A = Name
        Logger.log('✅ Found user at row ' + (i + 1) + ': ' + userName);
        break;
      }
    }
    
    if (!userName) {
      Logger.log('ERROR: User not found for email: ' + email);
      Logger.log('Available emails in sheet:');
      for (let i = 1; i < Math.min(usersData.length, 6); i++) {
        Logger.log('  Row ' + (i + 1) + ': "' + usersData[i][1] + '"');
      }
      return {
        success: false,
        error: 'User not found',
        totalOTHours: 0,
        remainingHours: 104,
        remainingColor: 'green'
      };
    }
    
    // Trim userName to handle any whitespace
    userName = userName.trim();
    Logger.log('User found: "' + userName + '"');
    
    // 2. Read all OT data and calculate total for this user
    const otData = otSheet.getDataRange().getValues();
    let totalHours = 0;
    let rowCount = 0;
    
    Logger.log('Scanning ' + (otData.length - 1) + ' OT rows...');
    
    // Start from row 1 (skip header)
    for (let i = 1; i < otData.length; i++) {
      const row = otData[i];
      
      // Skip empty rows
      if (!row[0]) continue;
      
      const rowName = row[0] ? row[0].toString().trim() : '';
      const hours = parseFloat(row[5]) || 0; // Column F = TotalHours
      
      // Match by name (trimmed comparison)
      if (rowName === userName) {
        Logger.log('Row ' + (i + 1) + ': Found match - Adding ' + hours + 'h');
        
        if (!isNaN(hours) && hours > 0) {
          totalHours += hours;
          rowCount++;
        }
      }
    }
    
    Logger.log('Total hours for ' + userName + ': ' + totalHours + ' (from ' + rowCount + ' applications)');
    
    // 3. Calculate remaining hours (104 - total)
    const remaining = CONFIG.MAX_OT_HOURS - totalHours;
    
    // 4. Determine color coding
    let color = 'green';
    if (remaining < 35) {
      color = 'amber';  // Warning: approaching limit
    }
    if (remaining <= 0) {
      color = 'red';    // Critical: at or over limit
    }
    
    Logger.log('Remaining hours: ' + remaining + ' (' + color + ')');
    Logger.log('=== calculateUserOTStats END ===');
    
    // 5. Return statistics
    return {
      success: true,
      userName: userName,
      email: email,
      totalOTHours: totalHours,
      remainingHours: remaining,
      remainingColor: color,
      applicationCount: rowCount,
      maxHours: CONFIG.MAX_OT_HOURS
    };
    
  } catch (error) {
    Logger.log('ERROR in calculateUserOTStats: ' + error.toString());
    return {
      success: false,
      error: error.message,
      totalOTHours: 0,
      remainingHours: 104,
      remainingColor: 'green'
    };
  }
}

/**
 * Calculate OT statistics for a specific user by name
 * @param {string} userName - User's name
 * @returns {Object} OT statistics object
 */
function calculateUserOTStatsByName(userName) {
  Logger.log('=== calculateUserOTStatsByName START ===');
  Logger.log('User Name: ' + userName);
  
  try {
    // Open spreadsheet by ID
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return {
        success: false,
        error: 'Cannot access database',
        totalOTHours: 0,
        remainingHours: 104,
        remainingColor: 'green'
      };
    }
    
    const otSheet = ss.getSheetByName(CONFIG.SHEET_OT);
    
    if (!otSheet) {
      Logger.log('ERROR: OT_Applications sheet not found');
      return {
        success: false,
        error: 'OT_Applications sheet not found',
        totalOTHours: 0,
        remainingHours: 104,
        remainingColor: 'green'
      };
    }
    
    // Trim userName to handle any whitespace
    userName = userName.trim();
    
    // Read all OT data and calculate total for this user
    const otData = otSheet.getDataRange().getValues();
    let totalHours = 0;
    let rowCount = 0;
    
    Logger.log('Scanning ' + (otData.length - 1) + ' OT rows...');
    
    // Start from row 1 (skip header)
    for (let i = 1; i < otData.length; i++) {
      const row = otData[i];
      
      // Skip empty rows
      if (!row[0]) continue;
      
      const rowName = row[0] ? row[0].toString().trim() : '';
      const hours = parseFloat(row[5]) || 0; // Column F = TotalHours
      
      // Match by name (trimmed comparison)
      if (rowName === userName) {
        Logger.log('Row ' + (i + 1) + ': Found match - Adding ' + hours + 'h');
        
        if (!isNaN(hours) && hours > 0) {
          totalHours += hours;
          rowCount++;
        }
      }
    }
    
    Logger.log('Total hours for ' + userName + ': ' + totalHours + ' (from ' + rowCount + ' applications)');
    
    // Calculate remaining hours (104 - total)
    const remaining = CONFIG.MAX_OT_HOURS - totalHours;
    
    // Determine color coding
    let color = 'green';
    if (remaining < 35) {
      color = 'amber';
    }
    if (remaining <= 0) {
      color = 'red';
    }
    
    Logger.log('Remaining hours: ' + remaining + ' (' + color + ')');
    Logger.log('=== calculateUserOTStatsByName END ===');
    
    return {
      success: true,
      userName: userName,
      totalOTHours: totalHours,
      remainingHours: remaining,
      remainingColor: color,
      applicationCount: rowCount,
      maxHours: CONFIG.MAX_OT_HOURS
    };
    
  } catch (error) {
    Logger.log('ERROR in calculateUserOTStatsByName: ' + error.toString());
    return {
      success: false,
      error: error.message,
      totalOTHours: 0,
      remainingHours: 104,
      remainingColor: 'green'
    };
  }
}

/**
 * Calculate OT statistics for all users (Management Dashboard)
 * @returns {Object} Aggregated statistics
 */
function calculateAllUsersOTStats() {
  Logger.log('=== calculateAllUsersOTStats START ===');
  
  try {
    // Open spreadsheet by ID
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    if (!ss) {
      Logger.log('ERROR: Cannot open spreadsheet by ID');
      return {
        success: false,
        error: 'Cannot access database',
        totalOTHours: 0,
        staffCount: 0,
        exceedCount: 0,
        staffStats: {}
      };
    }
    
    const usersSheet = ss.getSheetByName(CONFIG.SHEET_STAFF);
    const otSheet = ss.getSheetByName(CONFIG.SHEET_OT);
    
    if (!usersSheet || !otSheet) {
      Logger.log('ERROR: Required sheets not found');
      return {
        success: false,
        error: 'Required sheets not found',
        totalOTHours: 0,
        staffCount: 0,
        exceedCount: 0,
        staffStats: {}
      };
    }
    
    // Read all OT data
    const otData = otSheet.getDataRange().getValues();
    const staffHours = {}; // { "Staff Name": totalHours }
    let totalHours = 0;
    
    Logger.log('Processing ' + (otData.length - 1) + ' OT rows...');
    
    // Start from row 1 (skip header)
    for (let i = 1; i < otData.length; i++) {
      const row = otData[i];
      
      // Skip empty rows
      if (!row[0]) continue;
      
      const name = row[0] ? row[0].toString().trim() : '';
      const hours = parseFloat(row[5]) || 0; // Column F = TotalHours
      
      if (name && !isNaN(hours) && hours > 0) {
        if (!staffHours[name]) {
          staffHours[name] = 0;
        }
        staffHours[name] += hours;
        totalHours += hours;
      }
    }
    
    // Count staff exceeding limit
    const exceedCount = Object.values(staffHours).filter(h => h >= CONFIG.MAX_OT_HOURS).length;
    
    // Get total staff count
    const usersData = usersSheet.getDataRange().getValues();
    const staffCount = usersData.slice(1).filter(row => row[0] && row[0] !== '').length;
    
    Logger.log('Total staff: ' + staffCount);
    Logger.log('Total OT hours (all staff): ' + totalHours);
    Logger.log('Staff exceeding limit: ' + exceedCount);
    Logger.log('=== calculateAllUsersOTStats END ===');
    
    return {
      success: true,
      totalOTHours: totalHours,
      staffCount: staffCount,
      exceedCount: exceedCount,
      staffStats: staffHours, // { "Name": hours }
      maxHours: CONFIG.MAX_OT_HOURS
    };
    
  } catch (error) {
    Logger.log('ERROR in calculateAllUsersOTStats: ' + error.toString());
    return {
      success: false,
      error: error.message,
      totalOTHours: 0,
      staffCount: 0,
      exceedCount: 0,
      staffStats: {}
    };
  }
}

/**
 * Check if user can apply for specified hours
 * @param {string} email - User's email
 * @param {number} requestedHours - Hours user wants to apply for
 * @returns {Object} Validation result
 */
function validateOTApplication(email, requestedHours) {
  Logger.log('=== validateOTApplication START ===');
  Logger.log('Email: ' + email + ', Requested: ' + requestedHours + 'h');
  
  const stats = calculateUserOTStats(email);
  
  if (!stats.success) {
    return {
      valid: false,
      reason: stats.error,
      currentTotal: 0,
      remaining: 104,
      requested: requestedHours
    };
  }
  
  const afterApplication = stats.totalOTHours + requestedHours;
  const wouldRemain = CONFIG.MAX_OT_HOURS - afterApplication;
  
  if (wouldRemain < 0) {
    Logger.log('VALIDATION FAILED: Would exceed limit');
    return {
      valid: false,
      reason: 'Exceeds remaining hours',
      currentTotal: stats.totalOTHours,
      remaining: stats.remainingHours,
      requested: requestedHours,
      wouldExceedBy: Math.abs(wouldRemain)
    };
  }
  
  Logger.log('VALIDATION PASSED: Can apply for ' + requestedHours + 'h');
  return {
    valid: true,
    currentTotal: stats.totalOTHours,
    remaining: stats.remainingHours,
    requested: requestedHours,
    afterApplication: afterApplication,
    willRemain: wouldRemain
  };
}
