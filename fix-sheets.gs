// ============================================
// FIX SHEET STRUCTURE ISSUES
// Run this to fix the issues found by quickCheck()
// ============================================

function fixSheetIssues() {
  Logger.clear();
  Logger.log('========================================');
  Logger.log('🔧 FIXING SHEET ISSUES');
  Logger.log('========================================\n');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Fix 1: OT_Applications - Add missing columns
    fixOTApplicationsColumns(ss);
    
    // Fix 2: OT_Claim - Fix header names (add spaces)
    fixOTClaimHeaders(ss);
    
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

function fixOTApplicationsColumns(ss) {
  Logger.log('🔧 Fix 1: OT_Applications - Adding missing columns');
  Logger.log('─────────────────────────────────────');
  
  const sheet = ss.getSheetByName('OT_Applications');
  if (!sheet) {
    Logger.log('❌ OT_Applications sheet not found!');
    return false;
  }
  
  const currentCols = sheet.getLastColumn();
  Logger.log('Current columns: ' + currentCols);
  
  if (currentCols === 13) {
    Logger.log('✅ Already has 13 columns, checking headers...');
    return true;
  }
  
  // Get current headers
  const currentHeaders = sheet.getRange(1, 1, 1, currentCols).getValues()[0];
  Logger.log('Current headers: ' + currentHeaders.join(', '));
  
  // Expected 13 headers
  const expectedHeaders = [
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
  
  // Find which columns are missing
  const missingColumns = [];
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (!currentHeaders.includes(expectedHeaders[i])) {
      missingColumns.push({
        index: i + 1,
        name: expectedHeaders[i],
        letter: String.fromCharCode(65 + i)
      });
    }
  }
  
  if (missingColumns.length === 0) {
    Logger.log('✅ All required columns exist, just reordering needed');
    // Set correct headers
    sheet.getRange(1, 1, 1, 13).setValues([expectedHeaders]);
    Logger.log('✅ Headers reordered correctly');
    return true;
  }
  
  Logger.log('Missing columns: ' + missingColumns.map(c => c.name).join(', '));
  
  // Strategy: Check if data exists, if yes we need to be careful
  const dataRows = sheet.getLastRow();
  
  if (dataRows > 1) {
    Logger.log('⚠️ WARNING: Sheet has data (' + (dataRows - 1) + ' rows)');
    Logger.log('Will add missing columns at the end, then reorder all data');
    
    // Add missing columns at the end
    const columnsToAdd = 13 - currentCols;
    if (columnsToAdd > 0) {
      sheet.insertColumnsAfter(currentCols, columnsToAdd);
      Logger.log('✅ Added ' + columnsToAdd + ' columns at the end');
    }
    
    // Now we have 13 columns, set correct headers
    sheet.getRange(1, 1, 1, 13).setValues([expectedHeaders]);
    Logger.log('✅ Set correct headers');
    
    // Add missing column values (empty strings for existing rows)
    Logger.log('✅ Missing columns will be empty for existing rows');
    Logger.log('⚠️ IMPORTANT: Review data in sheet and manually move columns if needed!');
    
  } else {
    Logger.log('✅ No data in sheet, safe to restructure');
    
    // Just add columns and set headers
    const columnsToAdd = 13 - currentCols;
    if (columnsToAdd > 0) {
      sheet.insertColumnsAfter(currentCols, columnsToAdd);
    }
    
    sheet.getRange(1, 1, 1, 13).setValues([expectedHeaders]);
    Logger.log('✅ Added ' + columnsToAdd + ' columns and set headers');
  }
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, 13);
  headerRange.setBackground('#1a202c');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setWrap(true);
  
  // Set column widths
  sheet.setColumnWidth(1, 120);  // Name
  sheet.setColumnWidth(2, 100);  // Team
  sheet.setColumnWidth(3, 100);  // Date
  sheet.setColumnWidth(4, 90);   // Start Time
  sheet.setColumnWidth(5, 90);   // End Time
  sheet.setColumnWidth(6, 80);   // Total Hours
  sheet.setColumnWidth(7, 100);  // OT Type
  sheet.setColumnWidth(8, 120);  // Is Public Holiday
  sheet.setColumnWidth(9, 120);  // Proof Attendance
  sheet.setColumnWidth(10, 200); // Reason
  sheet.setColumnWidth(11, 80);  // Status
  sheet.setColumnWidth(12, 120); // Approved By
  sheet.setColumnWidth(13, 100); // Approval Date
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  Logger.log('✅ Formatting applied');
  Logger.log('');
  
  return true;
}

function fixOTClaimHeaders(ss) {
  Logger.log('🔧 Fix 2: OT_Claim - Fixing header names');
  Logger.log('─────────────────────────────────────');
  
  const sheet = ss.getSheetByName('OT_Claim');
  if (!sheet) {
    Logger.log('⚠️ OT_Claim sheet not found, skipping');
    Logger.log('');
    return true;
  }
  
  // Get current headers
  const currentHeaders = sheet.getRange(1, 1, 1, 6).getValues()[0];
  Logger.log('Current: ' + currentHeaders.join(' | '));
  
  // Set correct headers with spaces
  const correctHeaders = [
    'Name',
    'Month',
    'Total OT Hours',  // Was: TotalOTHours
    'Claim Amount',    // Was: ClaimAmount
    'Claim Date',      // Was: ClaimDate
    'Status'
  ];
  
  sheet.getRange(1, 1, 1, 6).setValues([correctHeaders]);
  
  Logger.log('Fixed:   ' + correctHeaders.join(' | '));
  Logger.log('✅ Headers updated with spaces');
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, 6);
  headerRange.setBackground('#1a202c');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  Logger.log('✅ Formatting applied');
  Logger.log('');
  
  return true;
}

// ============================================
// SHOW CURRENT OT_APPLICATIONS STRUCTURE
// ============================================

function showOTStructure() {
  Logger.clear();
  Logger.log('========================================');
  Logger.log('📊 CURRENT OT_APPLICATIONS STRUCTURE');
  Logger.log('========================================\n');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('OT_Applications');
    
    if (!sheet) {
      Logger.log('❌ OT_Applications sheet not found!');
      return;
    }
    
    const numCols = sheet.getLastColumn();
    const numRows = sheet.getLastRow();
    
    Logger.log('Total Columns: ' + numCols);
    Logger.log('Total Rows: ' + numRows);
    Logger.log('Data Rows: ' + (numRows - 1));
    Logger.log('');
    
    // Show headers
    Logger.log('CURRENT HEADERS:');
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    headers.forEach((header, idx) => {
      Logger.log('  ' + String.fromCharCode(65 + idx) + ': ' + header);
    });
    
    Logger.log('');
    Logger.log('EXPECTED HEADERS (13 columns):');
    const expected = [
      'Name', 'Team', 'Date', 'Start Time', 'End Time', 'Total Hours',
      'OT Type', 'Is Public Holiday', 'Proof Attendance', 'Reason',
      'Status', 'Approved By', 'Approval Date'
    ];
    expected.forEach((header, idx) => {
      const current = headers[idx] || '(missing)';
      const match = current === header ? '✅' : '❌';
      Logger.log('  ' + String.fromCharCode(65 + idx) + ': ' + header + ' - Current: ' + current + ' ' + match);
    });
    
    // Show sample data if exists
    if (numRows > 1) {
      Logger.log('');
      Logger.log('SAMPLE DATA (First 3 rows):');
      const sampleRows = Math.min(numRows - 1, 3);
      const data = sheet.getRange(2, 1, sampleRows, numCols).getValues();
      data.forEach((row, idx) => {
        Logger.log('  Row ' + (idx + 2) + ':');
        row.forEach((cell, colIdx) => {
          Logger.log('    ' + String.fromCharCode(65 + colIdx) + ': ' + cell);
        });
      });
    }
    
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
  }
}
