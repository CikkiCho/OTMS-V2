// ============================================
// SHEET STRUCTURE VERIFICATION & SETUP
// Run this once to verify/create all required sheets
// ============================================

function verifyAndSetupSheets() {
  Logger.log('=== STARTING SHEET VERIFICATION ===');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('✅ Connected to spreadsheet: ' + ss.getName());
    
    // 1. Verify/Setup Users Sheet
    setupUsersSheet(ss);
    
    // 2. Verify/Setup OT_Applications Sheet
    setupOTApplicationsSheet(ss);
    
    // 3. Verify/Setup OT_Claim Sheet
    setupOTClaimSheet(ss);
    
    // 4. Verify/Setup Leave_Applications Sheet
    setupLeaveApplicationsSheet(ss);

    // 5. Verify/Setup OT_History Sheet
    setupOTHistorySheet(ss);

    Logger.log('=== VERIFICATION COMPLETE ===');
    Logger.log('✅ All sheets verified and properly configured');
    
    return { success: true, message: 'All sheets verified successfully!' };
    
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ======== USERS SHEET ========
function setupUsersSheet(ss) {
  Logger.log('\n--- Checking Users Sheet ---');
  
  let usersSheet = ss.getSheetByName('Users');
  
  if (!usersSheet) {
    Logger.log('⚠️ Users sheet not found. Creating...');
    usersSheet = ss.insertSheet('Users');
    
    // Add headers
    const headers = ['Name', 'Email', 'Role', 'Team'];
    usersSheet.appendRow(headers);
    
    // Format header row
    const headerRange = usersSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a202c');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // Add sample data
    const sampleData = [
      ['Ali Hariz', 'alihariz@malaysiaairports.com.my', 'Staff', 'Operations'],
      ['Ahmad Zaki', 'ahmad.zaki@malaysiaairports.com.my', 'Staff', 'Maintenance'],
      ['Siti Nurhaliza', 'siti.n@malaysiaairports.com.my', 'Management', 'Operations'],
      ['Tan Wei Ming', 'tan.wm@malaysiaairports.com.my', 'Management', 'HR']
    ];
    
    sampleData.forEach(row => usersSheet.appendRow(row));
    
    // Auto-resize columns
    usersSheet.autoResizeColumns(1, headers.length);
    
    Logger.log('✅ Users sheet created with sample data');
  } else {
    Logger.log('✅ Users sheet exists');
    
    // Verify structure
    const headers = usersSheet.getRange(1, 1, 1, 4).getValues()[0];
    Logger.log('   Headers: ' + headers.join(', '));
    
    const dataRange = usersSheet.getDataRange();
    const numRows = dataRange.getNumRows();
    Logger.log('   Total rows (including header): ' + numRows);
    Logger.log('   User records: ' + (numRows - 1));
  }
}

// ======== OT_APPLICATIONS SHEET ========
function setupOTApplicationsSheet(ss) {
  Logger.log('\n--- Checking OT_Applications Sheet ---');
  
  let otSheet = ss.getSheetByName('OT_Applications');
  
  if (!otSheet) {
    Logger.log('⚠️ OT_Applications sheet not found. Creating...');
    otSheet = ss.insertSheet('OT_Applications');
    
    // Add headers (13 columns)
    const headers = [
      'Name',              // A
      'Team',              // B
      'Date',              // C
      'Start Time',        // D
      'End Time',          // E
      'Total Hours',       // F
      'OT Type',           // G
      'Is Public Holiday', // H
      'Proof Attendance',  // I
      'Reason',            // J
      'Status',            // K
      'Approved By',       // L
      'Approval Date'      // M
    ];
    otSheet.appendRow(headers);
    
    // Format header row
    const headerRange = otSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a202c');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setWrap(true);
    
    // Set column widths
    otSheet.setColumnWidth(1, 120);  // Name
    otSheet.setColumnWidth(2, 100);  // Team
    otSheet.setColumnWidth(3, 100);  // Date
    otSheet.setColumnWidth(4, 90);   // Start Time
    otSheet.setColumnWidth(5, 90);   // End Time
    otSheet.setColumnWidth(6, 80);   // Total Hours
    otSheet.setColumnWidth(7, 100);  // OT Type
    otSheet.setColumnWidth(8, 120);  // Is Public Holiday
    otSheet.setColumnWidth(9, 120);  // Proof Attendance
    otSheet.setColumnWidth(10, 200); // Reason
    otSheet.setColumnWidth(11, 80);  // Status
    otSheet.setColumnWidth(12, 120); // Approved By
    otSheet.setColumnWidth(13, 100); // Approval Date
    
    // Freeze header row
    otSheet.setFrozenRows(1);
    
    Logger.log('✅ OT_Applications sheet created with proper structure');
  } else {
    Logger.log('✅ OT_Applications sheet exists');
    
    // Verify structure
    const headers = otSheet.getRange(1, 1, 1, 13).getValues()[0];
    Logger.log('   Headers: ' + headers.join(', '));
    
    const dataRange = otSheet.getDataRange();
    const numRows = dataRange.getNumRows();
    const numCols = dataRange.getNumColumns();
    Logger.log('   Total rows (including header): ' + numRows);
    Logger.log('   Application records: ' + (numRows - 1));
    Logger.log('   Total columns: ' + numCols);
    
    // Verify expected column count
    if (numCols < 13) {
      Logger.log('   ⚠️ WARNING: Expected 13 columns, found ' + numCols);
      Logger.log('   Some columns may be missing!');
    }
    
    // Show sample data from last few rows
    if (numRows > 1) {
      const lastRow = Math.min(numRows, 5);
      const sampleData = otSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      Logger.log('   Sample records (Name, Team, Date, Start, End, Hours):');
      sampleData.forEach((row, idx) => {
        Logger.log('     Row ' + (idx + 2) + ': ' + row.join(' | '));
      });
    }
  }
}

// ======== OT_CLAIM SHEET ========
function setupOTClaimSheet(ss) {
  Logger.log('\n--- Checking OT_Claim Sheet ---');
  
  let claimSheet = ss.getSheetByName('OT_Claim');
  
  if (!claimSheet) {
    Logger.log('⚠️ OT_Claim sheet not found. Creating...');
    claimSheet = ss.insertSheet('OT_Claim');
    
    // Add headers (6 columns)
    const headers = [
      'Name',           // A
      'Month',          // B
      'Total OT Hours', // C
      'Claim Amount',   // D
      'Claim Date',     // E
      'Status'          // F
    ];
    claimSheet.appendRow(headers);
    
    // Format header row
    const headerRange = claimSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a202c');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // Set column widths
    claimSheet.setColumnWidth(1, 120);  // Name
    claimSheet.setColumnWidth(2, 100);  // Month
    claimSheet.setColumnWidth(3, 120);  // Total OT Hours
    claimSheet.setColumnWidth(4, 120);  // Claim Amount
    claimSheet.setColumnWidth(5, 100);  // Claim Date
    claimSheet.setColumnWidth(6, 80);   // Status
    
    // Freeze header row
    claimSheet.setFrozenRows(1);
    
    Logger.log('✅ OT_Claim sheet created with proper structure');
  } else {
    Logger.log('✅ OT_Claim sheet exists');
    
    const headers = claimSheet.getRange(1, 1, 1, 6).getValues()[0];
    Logger.log('   Headers: ' + headers.join(', '));
    
    const dataRange = claimSheet.getDataRange();
    const numRows = dataRange.getNumRows();
    Logger.log('   Total rows (including header): ' + numRows);
    Logger.log('   Claim records: ' + (numRows - 1));
  }
}

// ======== LEAVE_APPLICATIONS SHEET ========
function setupLeaveApplicationsSheet(ss) {
  Logger.log('\n--- Checking Leave_Applications Sheet ---');
  
  let leaveSheet = ss.getSheetByName('Leave_Applications');
  
  if (!leaveSheet) {
    Logger.log('⚠️ Leave_Applications sheet not found. Creating...');
    leaveSheet = ss.insertSheet('Leave_Applications');
    
    // Add headers (13 columns)
    const headers = [
      'Name',           // A
      'Team',           // B
      'Email',          // C
      'Leave Type',     // D
      'Start Date',     // E
      'End Date',       // F
      'Total Days',     // G
      'Reason',         // H
      'Supporting Doc', // I
      'Status',         // J
      'Approved By',    // K
      'Submitted Date', // L
      'Approval Date'   // M
    ];
    leaveSheet.appendRow(headers);
    
    // Format header row
    const headerRange = leaveSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a202c');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setWrap(true);
    
    // Set column widths
    leaveSheet.setColumnWidth(1, 120);  // Name
    leaveSheet.setColumnWidth(2, 100);  // Team
    leaveSheet.setColumnWidth(3, 180);  // Email
    leaveSheet.setColumnWidth(4, 150);  // Leave Type
    leaveSheet.setColumnWidth(5, 100);  // Start Date
    leaveSheet.setColumnWidth(6, 100);  // End Date
    leaveSheet.setColumnWidth(7, 80);   // Total Days
    leaveSheet.setColumnWidth(8, 200);  // Reason
    leaveSheet.setColumnWidth(9, 150);  // Supporting Doc
    leaveSheet.setColumnWidth(10, 80);  // Status
    leaveSheet.setColumnWidth(11, 120); // Approved By
    leaveSheet.setColumnWidth(12, 100); // Submitted Date
    leaveSheet.setColumnWidth(13, 100); // Approval Date
    
    // Freeze header row
    leaveSheet.setFrozenRows(1);
    
    Logger.log('✅ Leave_Applications sheet created with proper structure');
  } else {
    Logger.log('✅ Leave_Applications sheet exists');
    
    const headers = leaveSheet.getRange(1, 1, 1, 13).getValues()[0];
    Logger.log('   Headers: ' + headers.join(', '));
    
    const dataRange = leaveSheet.getDataRange();
    const numRows = dataRange.getNumRows();
    Logger.log('   Total rows (including header): ' + numRows);
    Logger.log('   Leave records: ' + (numRows - 1));
  }
}

// ======== OT_HISTORY SHEET ========
function setupOTHistorySheet(ss) {
  Logger.log('\n--- Checking OT_History Sheet ---');

  let historySheet = ss.getSheetByName('OT_History');

  if (!historySheet) {
    Logger.log('⚠️ OT_History sheet not found. Creating...');
    historySheet = ss.insertSheet('OT_History');

    // Add headers (12 columns)
    const headers = [
      'Month',              // A - e.g., "2025-11"
      'Year',               // B - e.g., 2025
      'Staff_Name',         // C
      'Staff_Email',        // D
      'Team',               // E
      'Total_OT_Hours',     // F - sum for the month
      'Applications_Count', // G - total applications
      'Approved_Count',     // H
      'Rejected_Count',     // I
      'Pending_Count',      // J
      'Applications_JSON',  // K - stringified array of all applications
      'Created_Date'        // L - timestamp of archive
    ];
    historySheet.appendRow(headers);

    // Format header row
    const headerRange = historySheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a202c');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setWrap(true);

    // Set column widths
    historySheet.setColumnWidth(1, 100);  // Month
    historySheet.setColumnWidth(2, 60);   // Year
    historySheet.setColumnWidth(3, 120);  // Staff_Name
    historySheet.setColumnWidth(4, 180);  // Staff_Email
    historySheet.setColumnWidth(5, 100);  // Team
    historySheet.setColumnWidth(6, 120);  // Total_OT_Hours
    historySheet.setColumnWidth(7, 150);  // Applications_Count
    historySheet.setColumnWidth(8, 120);  // Approved_Count
    historySheet.setColumnWidth(9, 120);  // Rejected_Count
    historySheet.setColumnWidth(10, 120); // Pending_Count
    historySheet.setColumnWidth(11, 400); // Applications_JSON
    historySheet.setColumnWidth(12, 150); // Created_Date

    // Freeze header row
    historySheet.setFrozenRows(1);

    Logger.log('✅ OT_History sheet created with proper structure');
  } else {
    Logger.log('✅ OT_History sheet exists');

    const headers = historySheet.getRange(1, 1, 1, 12).getValues()[0];
    Logger.log('   Headers: ' + headers.join(', '));

    const dataRange = historySheet.getDataRange();
    const numRows = dataRange.getNumRows();
    Logger.log('   Total rows (including header): ' + numRows);
    Logger.log('   History records: ' + (numRows - 1));
  }
}

// ======== DATA VALIDATION ========
function validateSheetData() {
  Logger.log('\n=== DATA VALIDATION CHECK ===');
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Check for duplicate names in Users
    const usersSheet = ss.getSheetByName('Users');
    if (usersSheet) {
      const userData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 2).getValues();
      const emails = userData.map(row => row[1]);
      const uniqueEmails = [...new Set(emails)];
      
      if (emails.length !== uniqueEmails.length) {
        Logger.log('⚠️ WARNING: Duplicate emails found in Users sheet');
      } else {
        Logger.log('✅ Users sheet: No duplicate emails');
      }
    }
    
    // Check for orphaned OT applications (users not in Users sheet)
    const otSheet = ss.getSheetByName('OT_Applications');
    if (otSheet && usersSheet) {
      const otData = otSheet.getRange(2, 1, Math.max(1, otSheet.getLastRow() - 1), 1).getValues();
      const userNames = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 1).getValues().map(r => r[0].trim());
      
      const orphanedRecords = otData.filter(row => {
        const name = (row[0] || '').toString().trim();
        return name && !userNames.includes(name);
      });
      
      if (orphanedRecords.length > 0) {
        Logger.log('⚠️ WARNING: Found ' + orphanedRecords.length + ' OT applications from users not in Users sheet:');
        orphanedRecords.forEach(row => Logger.log('   - ' + row[0]));
      } else {
        Logger.log('✅ OT Applications: All records linked to valid users');
      }
    }
    
    Logger.log('=== DATA VALIDATION COMPLETE ===');
    
  } catch (error) {
    Logger.log('❌ ERROR in validation: ' + error.toString());
  }
}

// ======== RUN ALL CHECKS ========
function runFullVerification() {
  verifyAndSetupSheets();
  validateSheetData();
  Logger.log('\n✅ Full verification complete. Check logs above for details.');
}
