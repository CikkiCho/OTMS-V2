// ============================================
// QUICK SHEET STRUCTURE CHECKER
// Run this to quickly verify all sheets are correct
// ============================================

function quickCheck() {
  Logger.clear();
  Logger.log('========================================');
  Logger.log('🔍 QUICK SHEET STRUCTURE CHECK');
  Logger.log('========================================\n');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('✅ Connected to: ' + ss.getName());
    Logger.log('📊 URL: ' + ss.getUrl());
    Logger.log('\n');
    
    let allGood = true;
    
    // Check each sheet
    allGood &= checkUsersSheet(ss);
    allGood &= checkOTApplicationsSheet(ss);
    allGood &= checkOTClaimSheet(ss);
    allGood &= checkLeaveSheet(ss);
    
    Logger.log('\n========================================');
    if (allGood) {
      Logger.log('✅ ALL CHECKS PASSED!');
      Logger.log('✅ System is ready for use');
    } else {
      Logger.log('⚠️ SOME ISSUES FOUND');
      Logger.log('⚠️ Review warnings above and run verifyAndSetupSheets() to fix');
    }
    Logger.log('========================================\n');
    
    return allGood;
    
  } catch (error) {
    Logger.log('❌ FATAL ERROR: ' + error.toString());
    Logger.log('❌ Cannot access spreadsheet. Check SPREADSHEET_ID.');
    return false;
  }
}

function checkUsersSheet(ss) {
  Logger.log('📋 Checking: Users');
  Logger.log('─────────────────────────────────────');
  
  const sheet = ss.getSheetByName('Users');
  if (!sheet) {
    Logger.log('❌ MISSING: Users sheet does not exist!');
    Logger.log('');
    return false;
  }
  
  const headers = sheet.getRange(1, 1, 1, 4).getValues()[0];
  const expectedHeaders = ['Name', 'Email', 'Role', 'Team'];
  
  let headerMatch = true;
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (headers[i] !== expectedHeaders[i]) {
      Logger.log('❌ Column ' + String.fromCharCode(65 + i) + ' header mismatch');
      Logger.log('   Expected: "' + expectedHeaders[i] + '"');
      Logger.log('   Found: "' + headers[i] + '"');
      headerMatch = false;
    }
  }
  
  if (headerMatch) {
    Logger.log('✅ Headers: Correct');
  }
  
  const numRows = sheet.getLastRow();
  const numUsers = numRows - 1;
  Logger.log('✅ Records: ' + numUsers + ' user(s)');
  
  if (numUsers === 0) {
    Logger.log('⚠️ WARNING: No users found! Add users to test system.');
  } else if (numUsers > 0) {
    // Check for duplicate emails
    const emails = sheet.getRange(2, 2, numUsers, 1).getValues().flat();
    const uniqueEmails = [...new Set(emails)];
    if (emails.length !== uniqueEmails.length) {
      Logger.log('⚠️ WARNING: Duplicate emails detected!');
    } else {
      Logger.log('✅ Email uniqueness: Verified');
    }
    
    // Show first user as sample
    const firstUser = sheet.getRange(2, 1, 1, 4).getValues()[0];
    Logger.log('✅ Sample user: ' + firstUser[0] + ' (' + firstUser[1] + ')');
  }
  
  Logger.log('');
  return headerMatch;
}

function checkOTApplicationsSheet(ss) {
  Logger.log('📋 Checking: OT_Applications');
  Logger.log('─────────────────────────────────────');
  
  const sheet = ss.getSheetByName('OT_Applications');
  if (!sheet) {
    Logger.log('❌ MISSING: OT_Applications sheet does not exist!');
    Logger.log('');
    return false;
  }
  
  const numCols = sheet.getLastColumn();
  if (numCols < 13) {
    Logger.log('❌ COLUMN COUNT: Expected 13, found ' + numCols);
    Logger.log('   Missing ' + (13 - numCols) + ' column(s)!');
    Logger.log('');
    return false;
  }
  
  const headers = sheet.getRange(1, 1, 1, 13).getValues()[0];
  const expectedHeaders = [
    'Name', 'Team', 'Date', 'Start Time', 'End Time', 'Total Hours',
    'OT Type', 'Is Public Holiday', 'Proof Attendance', 'Reason',
    'Status', 'Approved By', 'Approval Date'
  ];
  
  let headerMatch = true;
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (headers[i] !== expectedHeaders[i]) {
      Logger.log('❌ Column ' + String.fromCharCode(65 + i) + ' header mismatch');
      Logger.log('   Expected: "' + expectedHeaders[i] + '"');
      Logger.log('   Found: "' + headers[i] + '"');
      headerMatch = false;
    }
  }
  
  if (headerMatch) {
    Logger.log('✅ Headers: All 13 columns correct');
  }
  
  const numRows = sheet.getLastRow();
  const numRecords = numRows - 1;
  Logger.log('✅ Records: ' + numRecords + ' application(s)');
  
  if (numRecords > 0) {
    // Check for empty names
    const names = sheet.getRange(2, 1, numRecords, 1).getValues().flat();
    const emptyNames = names.filter(n => !n || n.toString().trim() === '').length;
    if (emptyNames > 0) {
      Logger.log('⚠️ WARNING: ' + emptyNames + ' record(s) with empty name!');
    }
    
    // Count status values
    const statuses = sheet.getRange(2, 11, numRecords, 1).getValues().flat();
    const pending = statuses.filter(s => s === 'Pending').length;
    const approved = statuses.filter(s => s === 'Approved').length;
    const rejected = statuses.filter(s => s === 'Rejected').length;
    
    Logger.log('✅ Status breakdown:');
    Logger.log('   Pending: ' + pending);
    Logger.log('   Approved: ' + approved);
    Logger.log('   Rejected: ' + rejected);
    
    // Show latest application
    const lastRow = sheet.getRange(numRows, 1, 1, 6).getValues()[0];
    Logger.log('✅ Latest: ' + lastRow[0] + ' | ' + lastRow[2] + ' | ' + lastRow[5] + 'h');
  }
  
  Logger.log('');
  return headerMatch;
}

function checkOTClaimSheet(ss) {
  Logger.log('📋 Checking: OT_Claim');
  Logger.log('─────────────────────────────────────');
  
  const sheet = ss.getSheetByName('OT_Claim');
  if (!sheet) {
    Logger.log('⚠️ OPTIONAL: OT_Claim sheet does not exist');
    Logger.log('   Will be created automatically when needed');
    Logger.log('');
    return true; // Not critical
  }
  
  const numCols = sheet.getLastColumn();
  if (numCols < 6) {
    Logger.log('❌ COLUMN COUNT: Expected 6, found ' + numCols);
    Logger.log('');
    return false;
  }
  
  const headers = sheet.getRange(1, 1, 1, 6).getValues()[0];
  const expectedHeaders = ['Name', 'Month', 'Total OT Hours', 'Claim Amount', 'Claim Date', 'Status'];
  
  let headerMatch = true;
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (headers[i] !== expectedHeaders[i]) {
      Logger.log('❌ Column ' + String.fromCharCode(65 + i) + ' header mismatch');
      Logger.log('   Expected: "' + expectedHeaders[i] + '"');
      Logger.log('   Found: "' + headers[i] + '"');
      headerMatch = false;
    }
  }
  
  if (headerMatch) {
    Logger.log('✅ Headers: All 6 columns correct');
  }
  
  const numRecords = sheet.getLastRow() - 1;
  Logger.log('✅ Records: ' + numRecords + ' claim(s)');
  
  Logger.log('');
  return headerMatch;
}

function checkLeaveSheet(ss) {
  Logger.log('📋 Checking: Leave_Applications');
  Logger.log('─────────────────────────────────────');
  
  const sheet = ss.getSheetByName('Leave_Applications');
  if (!sheet) {
    Logger.log('⚠️ OPTIONAL: Leave_Applications sheet does not exist');
    Logger.log('   Will be created automatically when needed');
    Logger.log('');
    return true; // Not critical
  }
  
  const numCols = sheet.getLastColumn();
  if (numCols < 13) {
    Logger.log('❌ COLUMN COUNT: Expected 13, found ' + numCols);
    Logger.log('');
    return false;
  }
  
  const headers = sheet.getRange(1, 1, 1, 13).getValues()[0];
  const expectedHeaders = [
    'Name', 'Team', 'Email', 'Leave Type', 'Start Date', 'End Date',
    'Total Days', 'Reason', 'Supporting Doc', 'Status', 'Approved By',
    'Submitted Date', 'Approval Date'
  ];
  
  let headerMatch = true;
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (headers[i] !== expectedHeaders[i]) {
      Logger.log('❌ Column ' + String.fromCharCode(65 + i) + ' header mismatch');
      Logger.log('   Expected: "' + expectedHeaders[i] + '"');
      Logger.log('   Found: "' + headers[i] + '"');
      headerMatch = false;
    }
  }
  
  if (headerMatch) {
    Logger.log('✅ Headers: All 13 columns correct');
  }
  
  const numRecords = sheet.getLastRow() - 1;
  Logger.log('✅ Records: ' + numRecords + ' leave request(s)');
  
  Logger.log('');
  return headerMatch;
}

// ============================================
// FIX COMMON ISSUES
// ============================================

function fixCommonIssues() {
  Logger.clear();
  Logger.log('========================================');
  Logger.log('🔧 FIXING COMMON ISSUES');
  Logger.log('========================================\n');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Fix 1: Trim all names in Users sheet
    Logger.log('🔧 Fix 1: Trimming names in Users sheet...');
    const usersSheet = ss.getSheetByName('Users');
    if (usersSheet) {
      const numRows = usersSheet.getLastRow();
      if (numRows > 1) {
        const names = usersSheet.getRange(2, 1, numRows - 1, 1).getValues();
        const trimmedNames = names.map(row => [row[0].toString().trim()]);
        usersSheet.getRange(2, 1, numRows - 1, 1).setValues(trimmedNames);
        Logger.log('✅ Trimmed ' + (numRows - 1) + ' user names');
      }
    }
    
    // Fix 2: Trim all names in OT_Applications sheet
    Logger.log('\n🔧 Fix 2: Trimming names in OT_Applications sheet...');
    const otSheet = ss.getSheetByName('OT_Applications');
    if (otSheet) {
      const numRows = otSheet.getLastRow();
      if (numRows > 1) {
        const names = otSheet.getRange(2, 1, numRows - 1, 1).getValues();
        const trimmedNames = names.map(row => [row[0].toString().trim()]);
        otSheet.getRange(2, 1, numRows - 1, 1).setValues(trimmedNames);
        Logger.log('✅ Trimmed ' + (numRows - 1) + ' application names');
      }
    }
    
    // Fix 3: Ensure all Status fields have valid values
    Logger.log('\n🔧 Fix 3: Normalizing Status values...');
    if (otSheet) {
      const numRows = otSheet.getLastRow();
      if (numRows > 1) {
        const statuses = otSheet.getRange(2, 11, numRows - 1, 1).getValues();
        let fixedCount = 0;
        const normalizedStatuses = statuses.map(row => {
          const status = row[0].toString().trim();
          if (status === '' || status.toLowerCase() === 'pending') {
            fixedCount++;
            return ['Pending'];
          } else if (status.toLowerCase() === 'approved') {
            return ['Approved'];
          } else if (status.toLowerCase() === 'rejected') {
            return ['Rejected'];
          }
          return [status];
        });
        otSheet.getRange(2, 11, numRows - 1, 1).setValues(normalizedStatuses);
        Logger.log('✅ Normalized ' + fixedCount + ' status values');
      }
    }
    
    // Fix 4: Freeze header rows
    Logger.log('\n🔧 Fix 4: Freezing header rows...');
    [usersSheet, otSheet, ss.getSheetByName('OT_Claim'), ss.getSheetByName('Leave_Applications')].forEach(sheet => {
      if (sheet && sheet.getFrozenRows() !== 1) {
        sheet.setFrozenRows(1);
        Logger.log('✅ Froze header row in: ' + sheet.getName());
      }
    });
    
    SpreadsheetApp.flush();
    
    Logger.log('\n========================================');
    Logger.log('✅ ALL FIXES APPLIED!');
    Logger.log('✅ Run quickCheck() to verify');
    Logger.log('========================================\n');
    
    return true;
    
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
    return false;
  }
}

// ============================================
// GET SUMMARY REPORT
// ============================================

function getSummaryReport() {
  Logger.clear();
  Logger.log('========================================');
  Logger.log('📊 OTMS SYSTEM SUMMARY REPORT');
  Logger.log('========================================\n');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    Logger.log('📁 Spreadsheet: ' + ss.getName());
    Logger.log('🔗 URL: ' + ss.getUrl());
    Logger.log('📅 Last Updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
    Logger.log('\n');
    
    // Users summary
    const usersSheet = ss.getSheetByName('Users');
    if (usersSheet) {
      const numUsers = usersSheet.getLastRow() - 1;
      Logger.log('👥 Total Users: ' + numUsers);
      if (numUsers > 0) {
        const roles = usersSheet.getRange(2, 3, numUsers, 1).getValues().flat();
        const staff = roles.filter(r => r === 'Staff').length;
        const management = roles.filter(r => r === 'Management').length;
        Logger.log('   • Staff: ' + staff);
        Logger.log('   • Management: ' + management);
      }
    }
    
    // OT Applications summary
    const otSheet = ss.getSheetByName('OT_Applications');
    if (otSheet) {
      const numApps = otSheet.getLastRow() - 1;
      Logger.log('\n📝 Total OT Applications: ' + numApps);
      if (numApps > 0) {
        const statuses = otSheet.getRange(2, 11, numApps, 1).getValues().flat();
        const pending = statuses.filter(s => s === 'Pending').length;
        const approved = statuses.filter(s => s === 'Approved').length;
        const rejected = statuses.filter(s => s === 'Rejected').length;
        Logger.log('   • Pending: ' + pending + ' ⏳');
        Logger.log('   • Approved: ' + approved + ' ✅');
        Logger.log('   • Rejected: ' + rejected + ' ❌');
        
        // Total hours
        const hours = otSheet.getRange(2, 6, numApps, 1).getValues().flat();
        const totalHours = hours.reduce((sum, h) => sum + (parseFloat(h) || 0), 0);
        Logger.log('   • Total Hours: ' + totalHours.toFixed(2) + ' hours');
      }
    }
    
    // OT Claims summary
    const claimSheet = ss.getSheetByName('OT_Claim');
    if (claimSheet) {
      const numClaims = claimSheet.getLastRow() - 1;
      Logger.log('\n💰 Total OT Claims: ' + numClaims);
      if (numClaims > 0) {
        const amounts = claimSheet.getRange(2, 4, numClaims, 1).getValues().flat();
        const totalAmount = amounts.reduce((sum, a) => sum + (parseFloat(a) || 0), 0);
        Logger.log('   • Total Amount: RM ' + totalAmount.toFixed(2));
      }
    }
    
    // Leave Applications summary
    const leaveSheet = ss.getSheetByName('Leave_Applications');
    if (leaveSheet) {
      const numLeaves = leaveSheet.getLastRow() - 1;
      Logger.log('\n🏖️ Total Leave Requests: ' + numLeaves);
      if (numLeaves > 0) {
        const days = leaveSheet.getRange(2, 7, numLeaves, 1).getValues().flat();
        const totalDays = days.reduce((sum, d) => sum + (parseInt(d) || 0), 0);
        Logger.log('   • Total Days: ' + totalDays + ' days');
      }
    }
    
    Logger.log('\n========================================');
    Logger.log('Report generated successfully!');
    Logger.log('========================================\n');
    
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
  }
}
