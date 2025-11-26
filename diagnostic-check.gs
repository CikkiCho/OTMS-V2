// Quick diagnostic to copy and paste into Apps Script console
function runFullDiagnostic() {
  Logger.log('=== STARTING FULL SYSTEM DIAGNOSTIC ===');
  Logger.log('Time: ' + new Date().toISOString());
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // Check 1: Sheet existence
  Logger.log('\n=== SHEET EXISTENCE CHECK ===');
  const usersSheet = ss.getSheetByName(SHEET_STAFF);
  const otSheet = ss.getSheetByName(SHEET_OT);
  Logger.log('Users sheet exists: ' + (usersSheet ? 'YES' : 'NO'));
  Logger.log('OT_Applications sheet exists: ' + (otSheet ? 'YES' : 'NO'));
  
  if (!otSheet) {
    Logger.log('ERROR: OT_Applications sheet not found!');
    return;
  }
  
  // Check 2: OT_Applications structure
  Logger.log('\n=== OT_APPLICATIONS SHEET STRUCTURE ===');
  const lastCol = otSheet.getLastColumn();
  const lastRow = otSheet.getLastRow();
  Logger.log('Total columns: ' + lastCol);
  Logger.log('Total rows (including header): ' + lastRow);
  Logger.log('Data rows: ' + (lastRow - 1));
  
  if (lastCol < 13) {
    Logger.log('❌ WARNING: Sheet has ' + lastCol + ' columns, needs 13!');
  } else {
    Logger.log('✅ Column count is correct (13+)');
  }
  
  // Check 3: Header row
  Logger.log('\n=== HEADER ROW ===');
  const headerRow = otSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < Math.min(headerRow.length, 13); i++) {
    Logger.log('Col ' + (i+1) + ' (' + String.fromCharCode(65+i) + '): "' + headerRow[i] + '"');
  }
  
  // Check 4: Data rows analysis
  Logger.log('\n=== DATA ROWS ANALYSIS ===');
  const otData = otSheet.getDataRange().getValues();
  let emptyRows = 0;
  let populatedRows = 0;
  let pendingRows = 0;
  let blankStatusRows = 0;
  
  for (let i = 1; i < otData.length; i++) {
    const row = otData[i];
    
    if (!row[0]) {
      emptyRows++;
      continue;
    }
    
    populatedRows++;
    const name = row[0];
    const date = row[2];
    const hours = row[5];
    const status = row[10];
    
    if (!status || status === '') {
      blankStatusRows++;
    }
    
    if (status === 'Pending') {
      pendingRows++;
    }
    
    if (i <= 5 || populatedRows <= 5) {
      Logger.log('\nRow ' + (i+1) + ':');
      Logger.log('  Name: "' + name + '"');
      Logger.log('  Team: "' + row[1] + '"');
      Logger.log('  Date: "' + date + '"');
      Logger.log('  Hours: ' + hours);
      Logger.log('  Status: "' + status + '" (blank=' + (!status) + ')');
    }
  }
  
  Logger.log('\n=== SUMMARY ===');
  Logger.log('Total empty rows: ' + emptyRows);
  Logger.log('Total populated rows: ' + populatedRows);
  Logger.log('Rows with Status="Pending": ' + pendingRows);
  Logger.log('Rows with blank Status: ' + blankStatusRows);
  
  // Check 5: User data
  Logger.log('\n=== USERS SHEET CHECK ===');
  const usersData = usersSheet.getDataRange().getValues();
  Logger.log('Total users (including header): ' + usersData.length);
  Logger.log('Total user rows: ' + (usersData.length - 1));
  
  for (let i = 1; i < Math.min(usersData.length, 6); i++) {
    Logger.log('Row ' + (i+1) + ': Name="' + usersData[i][0] + '" Email="' + usersData[i][1] + '" Role="' + usersData[i][2] + '"');
  }
  
  Logger.log('\n=== END DIAGNOSTIC ===');
}
